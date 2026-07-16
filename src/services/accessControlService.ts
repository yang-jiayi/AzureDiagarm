export interface AccessIdentity {
  enabled: boolean;
  authenticated: boolean;
  email: string | null;
  isAdmin: boolean;
  allowed: boolean;
}

export interface AllowedUser {
  email: string;
  addedAt: string;
  addedBy: string;
  isAdmin: boolean;
  immutable: boolean;
}

async function parseError(response: Response): Promise<Error> {
  try {
    const payload = await response.json();
    if (typeof payload?.error === 'string' && payload.error.length > 0) {
      return new Error(payload.error);
    }
  } catch {
    // The generic status-based message below is sufficient.
  }
  return new Error(`Access service request failed (${response.status}).`);
}

async function requestJson<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(path, {
    credentials: 'same-origin',
    ...init,
    headers: {
      Accept: 'application/json',
      ...(init?.headers || {}),
    },
  });
  if (!response.ok) throw await parseError(response);
  return response.json() as Promise<T>;
}

export function getAccessIdentity(): Promise<AccessIdentity> {
  return requestJson<AccessIdentity>('/api/access/me');
}

export async function listAllowedUsers(): Promise<AllowedUser[]> {
  const payload = await requestJson<{ users: AllowedUser[] }>('/api/access/users');
  return payload.users;
}

export async function addAllowedUser(email: string): Promise<AllowedUser> {
  const payload = await requestJson<{ user: AllowedUser }>('/api/access/users', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email }),
  });
  return payload.user;
}

export async function removeAllowedUser(email: string): Promise<void> {
  const response = await fetch('/api/access/users', {
    method: 'DELETE',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email }),
  });
  if (!response.ok) throw await parseError(response);
}
