// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export function getCategoryColor(category: string): string {
  const colors: Record<string, string> = {
    compute: '#0078d4', containers: '#0078d4',
    databases: '#10b981', storage: '#10b981', 'data layer': '#10b981',
    'ai + machine learning': '#f59e0b', analytics: '#8b5cf6',
    networking: '#06b6d4', identity: '#ec4899', security: '#ef4444',
    monitor: '#6366f1', integration: '#14b8a6', iot: '#f97316',
    'app services': '#3b82f6', web: '#3b82f6', devops: '#8b5cf6',
  };
  return colors[category?.toLowerCase() || ''] || '#6b7280';
}

export interface GroupColors {
  bg: string;
  border: string;
  header: string;
}

export function getGroupColors(label: string): GroupColors {
  const rules: Array<[string[], string]> = [
    [['web', 'frontend', 'ingress', 'edge'], '#6b7280'],
    [['compute', 'processing', 'microservices', 'api'], '#0078d4'],
    [['data', 'storage', 'database', 'persistence'], '#10b981'],
    [['ai', 'intelligence', 'analytics', 'ml', 'cognitive'], '#f59e0b'],
    [['iot', 'device', 'telemetry'], '#f97316'],
    [['security', 'auth', 'identity', 'vault'], '#ef4444'],
    [['monitor', 'ops', 'observability', 'logging'], '#8b5cf6'],
    [['network', 'integration', 'messaging', 'event', 'ingestion'], '#06b6d4'],
    [['container', 'registry', 'runtime'], '#0078d4'],
  ];
  const lower = label.toLowerCase();
  const match = rules.find(([keywords]) => keywords.some(keyword => lower.includes(keyword)));
  const color = match?.[1] ?? '#6b7280';
  const rgb = [1, 3, 5].map(offset => parseInt(color.slice(offset, offset + 2), 16));
  return { bg: `rgba(${rgb.join(', ')}, ${match ? '0.25' : '0.20'})`, border: color, header: color };
}
