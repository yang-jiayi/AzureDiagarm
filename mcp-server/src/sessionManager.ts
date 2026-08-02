// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * SessionStore — cap and expire MCP HTTP sessions.
 *
 * Safety properties enforced here:
 *  • MAX_SESSIONS: hard cap on concurrent live + pending-init sessions (default 100).
 *  • IDLE_TTL_MS:  sessions idle longer than this are reaped (default 30 min).
 *  • ABS_TTL_MS:   sessions older than this are reaped regardless of activity
 *                  (default 2 h), preventing a long-running client from keeping
 *                  server-side memory pinned indefinitely.
 *  • GC runs on a fixed interval (default 60 s) so expired entries are cleaned
 *    without requiring every request to pay a full-map scan.
 *
 * Capacity admission uses a two-phase reservation protocol so that concurrent
 * initialize requests cannot race past the cap:
 *
 *   1. reserve()            — claim a slot before the transport is created.
 *                             Returns false if at cap (caller should 503).
 *   2. commitReservation()  — called from onsessioninitialized once the SDK
 *                             has assigned the session ID; binds ID → transport.
 *   3. releaseReservation() — called if initialization fails or the transport
 *                             closes before onsessioninitialized fires.
 *
 * Env-var overrides (all optional):
 *  MCP_SESSION_MAX          – integer, max concurrent sessions (≥1, ≤10_000)
 *  MCP_SESSION_IDLE_SECONDS – idle TTL in seconds
 *  MCP_SESSION_TTL_SECONDS  – absolute TTL in seconds
 *  MCP_SESSION_GC_SECONDS   – GC sweep interval in seconds
 */

import type { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';

export interface SessionStoreConfig {
  maxSessions?: number;
  idleTtlMs?: number;
  absTtlMs?: number;
  gcIntervalMs?: number;
}

interface SessionEntry {
  transport: StreamableHTTPServerTransport;
  createdAt: number;
  lastUsedAt: number;
}

/** Parse a positive integer env var, returning `defaultValue` on invalid input.
 * Rejects values with trailing non-numeric junk (e.g. '10oops') by comparing
 * the re-stringified parsed integer against the trimmed raw string. */
function envInt(name: string, defaultValue: number, min = 1, max = Number.MAX_SAFE_INTEGER): number {
  const raw = process.env[name];
  if (!raw) return defaultValue;
  const trimmed = raw.trim();
  const n = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(n) || n < min || n > max || String(n) !== trimmed) {
    console.error(`[session-store] invalid ${name}=${JSON.stringify(raw)}, using default ${defaultValue}`);
    return defaultValue;
  }
  return n;
}

export class SessionStore {
  // Defaults — readable by tests.
  static readonly DEFAULT_MAX_SESSIONS = 100;
  static readonly DEFAULT_IDLE_TTL_MS = 30 * 60 * 1000;    // 30 min
  static readonly DEFAULT_ABS_TTL_MS = 2 * 60 * 60 * 1000; // 2 h
  static readonly DEFAULT_GC_INTERVAL_MS = 60 * 1000;       // 1 min

  readonly maxSessions: number;
  readonly idleTtlMs: number;
  readonly absTtlMs: number;
  readonly gcIntervalMs: number;

  private readonly _sessions = new Map<string, SessionEntry>();
  /** Slots claimed by in-flight initialize requests (reservation phase). */
  private _pendingReservations = 0;
  private _gcTimer: ReturnType<typeof setInterval> | null = null;

  constructor(cfg: SessionStoreConfig = {}) {
    const idleFromEnv = envInt('MCP_SESSION_IDLE_SECONDS', SessionStore.DEFAULT_IDLE_TTL_MS / 1000, 1) * 1000;
    const absFromEnv  = envInt('MCP_SESSION_TTL_SECONDS',  SessionStore.DEFAULT_ABS_TTL_MS  / 1000, 1) * 1000;
    const gcFromEnv   = envInt('MCP_SESSION_GC_SECONDS',   SessionStore.DEFAULT_GC_INTERVAL_MS / 1000, 1) * 1000;

    this.maxSessions  = Math.max(1, cfg.maxSessions ?? envInt('MCP_SESSION_MAX', SessionStore.DEFAULT_MAX_SESSIONS, 1, 10_000));
    // Explicit cfg values are used as-is (minimum 1 ms) so tests can pass small
    // durations. Env-var-derived values are floored at 1 000 ms to protect
    // against accidental misconfiguration in production.
    this.idleTtlMs    = cfg.idleTtlMs  != null ? Math.max(1, cfg.idleTtlMs)  : Math.max(1000, idleFromEnv);
    // Explicit absTtlMs is used as-is (minimum 1 ms) — in production it should
    // be >= idleTtlMs but tests may set a very short abs TTL independently.
    this.absTtlMs     = cfg.absTtlMs   != null ? Math.max(1, cfg.absTtlMs)   : Math.max(this.idleTtlMs, absFromEnv);
    this.gcIntervalMs = cfg.gcIntervalMs != null ? Math.max(1, cfg.gcIntervalMs) : Math.max(1000, gcFromEnv);
  }

  /** Number of committed (fully initialized) live sessions. */
  get size(): number {
    return this._sessions.size;
  }

  /** Number of slots reserved for in-flight initialize requests. */
  get pendingReservations(): number {
    return this._pendingReservations;
  }

  /**
   * Start the background GC timer.
   * Call this once after construction.
   */
  startGc(): void {
    if (this._gcTimer) return;
    this._gcTimer = setInterval(() => this.evictExpired(), this.gcIntervalMs);
    // unref so the timer won't keep the Node process alive when idle
    if (typeof this._gcTimer.unref === 'function') this._gcTimer.unref();
  }

  // ── Two-phase reservation API ─────────────────────────────────────────────

  /**
   * Phase 1: Atomically claim a capacity slot for an in-flight initialize.
   *
   * Runs evictExpired() first so recently-freed slots are visible.
   * Returns `true` on success, `false` when live + pending >= maxSessions.
   */
  reserve(): boolean {
    this.evictExpired();
    if (this._sessions.size + this._pendingReservations >= this.maxSessions) {
      return false;
    }
    this._pendingReservations++;
    return true;
  }

  /**
   * Phase 2a: Bind the reserved slot to a real session ID + transport.
   * Call from onsessioninitialized once the SDK assigns the session ID.
   */
  commitReservation(sid: string, transport: StreamableHTTPServerTransport): void {
    if (this._pendingReservations > 0) this._pendingReservations--;
    const now = Date.now();
    this._sessions.set(sid, { transport, createdAt: now, lastUsedAt: now });
  }

  /**
   * Phase 2b: Release a reserved slot without committing (initialization
   * failed or transport closed before onsessioninitialized fired).
   */
  releaseReservation(): void {
    if (this._pendingReservations > 0) this._pendingReservations--;
  }

  // ── Session lookup / lifecycle ────────────────────────────────────────────

  /**
   * Look up a committed session.
   * Returns the transport if found AND not expired, otherwise `undefined`.
   * Side-effects: evicts the session on expiry.
   */
  get(sid: string): StreamableHTTPServerTransport | undefined {
    const entry = this._sessions.get(sid);
    if (!entry) return undefined;
    if (this._isExpired(entry)) {
      this._evictOne(sid, entry, 'get-expired');
      return undefined;
    }
    return entry.transport;
  }

  /**
   * Update the last-used timestamp (call after each successful request).
   */
  touch(sid: string): void {
    const entry = this._sessions.get(sid);
    if (entry) entry.lastUsedAt = Date.now();
  }

  /**
   * Remove a committed session by id (e.g. on transport.onclose).
   */
  delete(sid: string): void {
    this._sessions.delete(sid);
  }

  /**
   * Walk the whole map and evict any expired entries.
   * Called by the GC timer and on every `reserve()`.
   */
  evictExpired(): void {
    for (const [sid, entry] of this._sessions) {
      if (this._isExpired(entry)) {
        this._evictOne(sid, entry, 'gc');
      }
    }
  }

  /**
   * Close all sessions and stop the GC timer (call during graceful shutdown).
   */
  async closeAll(): Promise<void> {
    if (this._gcTimer) {
      clearInterval(this._gcTimer);
      this._gcTimer = null;
    }
    this._pendingReservations = 0;
    const closures: Promise<void>[] = [];
    for (const [sid, entry] of this._sessions) {
      this._sessions.delete(sid);
      closures.push(
        entry.transport.close().catch((err) =>
          console.error(`[session-store] error closing session ${sid}:`, err),
        ),
      );
    }
    await Promise.allSettled(closures);
  }

  // ── private helpers ────────────────────────────────────────────────────────

  private _isExpired(entry: SessionEntry): boolean {
    const now = Date.now();
    return (
      now - entry.lastUsedAt > this.idleTtlMs ||
      now - entry.createdAt > this.absTtlMs
    );
  }

  private _evictOne(sid: string, entry: SessionEntry, reason: string): void {
    this._sessions.delete(sid);
    console.error(`[session-store] evicted session ${sid} (${reason})`);
    entry.transport.close().catch((err) =>
      console.error(`[session-store] error closing session ${sid}:`, err),
    );
  }
}
