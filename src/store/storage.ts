const SCHEMA_VERSION = 1;

// Arrays can't be spread-merged with {_v} — we box them under _arr instead.
type Envelope = { _v: number; _arr?: unknown[] } & Record<string, unknown>;

function pack(value: unknown): Envelope {
  if (Array.isArray(value)) return { _v: SCHEMA_VERSION, _arr: value };
  return { ...(value as Record<string, unknown>), _v: SCHEMA_VERSION };
}

function unpack<T>(env: Envelope): T {
  if ('_arr' in env) return env._arr as T;
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const { _v, ...rest } = env;
  return rest as T;
}

function getStorage(): Storage | null {
  if (typeof localStorage === 'undefined') return null;
  return localStorage;
}

export const storage = {
  get<T>(key: string): T | null {
    try {
      const store = getStorage();
      if (!store) return null;
      const raw = store.getItem(key);
      if (!raw) return null;
      const env = JSON.parse(raw) as Envelope;
      if (env._v !== SCHEMA_VERSION) return null;
      return unpack<T>(env);
    } catch {
      return null;
    }
  },

  put<T>(key: string, value: T): void {
    const serialized = JSON.stringify(pack(value));
    try {
      const store = getStorage();
      if (!store) return;
      store.setItem(key, serialized);
    } catch (e) {
      if (typeof DOMException !== 'undefined' && e instanceof DOMException && e.name === 'QuotaExceededError') {
        // Evict oldest usage day buckets then retry once
        evictOldestUsageBucket();
        try {
          const store = getStorage();
          if (!store) return;
          store.setItem(key, serialized);
        } catch {
          if (typeof window !== 'undefined') {
            window.dispatchEvent(new CustomEvent('xl:quota-warning', { detail: { key } }));
          }
        }
      }
    }
  },

  remove(key: string): void {
    const store = getStorage();
    if (!store) return;
    store.removeItem(key);
  },
};

function evictOldestUsageBucket(): void {
  const store = getStorage();
  if (!store) return;
  const keys: string[] = [];
  for (let i = 0; i < store.length; i++) {
    const k = store.key(i);
    if (k?.startsWith('xl.usage.day.')) keys.push(k);
  }
  keys.sort(); // ISO dates sort lexicographically
  if (keys.length > 0) store.removeItem(keys[0]);
}
