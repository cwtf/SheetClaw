import type { StateCreator } from 'zustand';
import type { UsageRecord } from '../../types';
import { storage } from '../storage';

const ROLLING_DAYS = 30;

export interface UsageIndex {
  days: Record<string, number>; // YYYY-MM-DD → record count
  oldest: string;
  newest: string;
  totalBytesApprox: number;
}

function pruneOldBuckets(today: string): void {
  const cutoff = new Date(today);
  cutoff.setDate(cutoff.getDate() - ROLLING_DAYS);
  const cutoffStr = cutoff.toISOString().slice(0, 10);
  const keysToRemove: string[] = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k?.startsWith('xl.usage.day.') && k.slice('xl.usage.day.'.length) < cutoffStr) {
      keysToRemove.push(k);
    }
  }
  for (const k of keysToRemove) storage.remove(k);
}

export interface SessionTotals {
  sessionId: string;
  inputTokens: number;
  outputTokens: number;
  costUsd: number;
  turns: number;
}

export interface UsageSlice {
  sessionTotals: SessionTotals | null;
  recordUsage(record: UsageRecord): void;
  resetSessionTotals(sessionId: string, totals?: Partial<Omit<SessionTotals, 'sessionId'>>): void;
  clearSessionTotals(): void;
}

export const createUsageSlice: StateCreator<UsageSlice> = set => ({
  sessionTotals: null,

  recordUsage(record) {
    const day = record.timestamp.slice(0, 10);

    // Prune buckets older than 30 days
    pruneOldBuckets(day);

    // Persist to daily bucket
    const key = `xl.usage.day.${day}`;
    const existing = storage.get<UsageRecord[]>(key);
    const updated = [...(Array.isArray(existing) ? existing : []), record];
    storage.put(key, updated);

    // Update index
    const idx = storage.get<UsageIndex>('xl.usage.index') ?? { days: {}, oldest: day, newest: day, totalBytesApprox: 0 };
    idx.days[day] = updated.length;
    if (day < idx.oldest || !idx.oldest) idx.oldest = day;
    if (day > idx.newest || !idx.newest) idx.newest = day;
    storage.put('xl.usage.index', idx);

    // Update in-memory session totals
    set(state => {
      const prev = state.sessionTotals;
      if (!prev || prev.sessionId !== record.sessionId) return {};
      return {
        sessionTotals: {
          ...prev,
          inputTokens: prev.inputTokens + record.inputTokens,
          outputTokens: prev.outputTokens + record.outputTokens,
          costUsd: prev.costUsd + record.estimatedCostUsd,
          turns: prev.turns + 1,
        },
      };
    });
  },

  resetSessionTotals(sessionId, totals) {
    set({
      sessionTotals: {
        sessionId,
        inputTokens: totals?.inputTokens ?? 0,
        outputTokens: totals?.outputTokens ?? 0,
        costUsd: totals?.costUsd ?? 0,
        turns: totals?.turns ?? 0,
      },
    });
  },

  clearSessionTotals() {
    set({ sessionTotals: null });
  },
});
