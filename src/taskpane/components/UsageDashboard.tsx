import { useState, useCallback, type CSSProperties } from 'react';
import {
  Body1Strong,
  Button,
  Caption1,
  Card,
  CardHeader,
  Select,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  tokens,
} from '@fluentui/react-components';
import {
  queryTotals,
  queryByProvider,
  queryByModel,
  queryByDay,
  type TimeRange,
} from '../../usage/queries';
import { exportCsv } from '../../usage/export';
import { storage } from '../../store/storage';

function fmtTokens(n: number): string {
  if (n >= 1_000_000) return (n / 1_000_000).toFixed(2) + 'M';
  if (n >= 1_000) return (n / 1_000).toFixed(1) + 'k';
  return String(n);
}

const breakdownTableStyle: CSSProperties = {
  width: '100%',
  tableLayout: 'fixed',
};

const nameColumnStyle: CSSProperties = {
  width: '68%',
  minWidth: 0,
};

const numericColumnStyle: CSSProperties = {
  width: '32%',
  textAlign: 'right',
  whiteSpace: 'nowrap',
};

const modelNameStyle: CSSProperties = {
  display: 'block',
  fontFamily: 'monospace',
  overflowWrap: 'anywhere',
  wordBreak: 'break-word',
  lineHeight: 1.25,
  minWidth: 0,
};

export default function UsageDashboard() {
  const [range, setRange] = useState<TimeRange>('week');
  const [tick, setTick] = useState(0);

  const refresh = useCallback(() => setTick(t => t + 1), []);

  const totals = queryTotals(range);
  const byProvider = queryByProvider(range);
  const byModel = queryByModel(range);
  const byDay = queryByDay(range);

  void tick; // consumed to trigger re-render on refresh

  function resetHistory() {
    if (!confirm('Delete all usage history? This cannot be undone.')) return;
    for (let i = localStorage.length - 1; i >= 0; i--) {
      const k = localStorage.key(i);
      if (k?.startsWith('xl.usage.')) storage.remove(k);
    }
    refresh();
  }

  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      gap: 12,
      padding: 12,
      overflowY: 'auto',
      height: '100%',
      boxSizing: 'border-box',
    }}>
      {/* Controls */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <Body1Strong>Usage</Body1Strong>
        <Select
          size="small"
          value={range}
          onChange={(_, d) => setRange(d.value as TimeRange)}
          style={{ flex: 1 }}
        >
          <option value="today">Today</option>
          <option value="week">Last 7 days</option>
          <option value="month">Last 30 days</option>
          <option value="all">All time</option>
        </Select>
        <Button size="small" appearance="subtle" onClick={() => exportCsv(range)}>Export CSV</Button>
      </div>

      {totals.turns === 0 ? (
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>No usage yet — start a chat.</Caption1>
        </div>
      ) : (
        <>
          {/* Totals cards */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 8 }}>
            <TotalsCard label="Tokens" value={fmtTokens(totals.totalTokens)} />
          </div>

          {/* By day (sparkline) */}
          {byDay.length > 1 && <DaySparkline rows={byDay} />}

          {/* By model */}
          <BreakdownTable title="By model" rows={byModel} />

          {/* By provider */}
          {byProvider.length > 1 && <BreakdownTable title="By provider" rows={byProvider} />}
        </>
      )}

      {/* Actions */}
      <div style={{
        marginTop: 'auto',
        paddingTop: 10,
        borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
        display: 'flex',
        justifyContent: 'flex-start',
        flexShrink: 0,
      }}>
        <Button
          size="small"
          appearance="secondary"
          style={{
            color: tokens.colorPaletteRedForeground1,
            borderColor: tokens.colorPaletteRedBorder2,
            minWidth: 96,
          }}
          onClick={resetHistory}
        >
          Reset history
        </Button>
      </div>
    </div>
  );
}

// ── Sub-components ─────────────────────────────────────────────────────────

function TotalsCard({ label, value, muted }: { label: string; value: string; muted?: boolean }) {
  return (
    <Card size="small">
      <CardHeader
        header={<Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{label}</Caption1>}
        description={
          <Body1Strong style={{ color: muted ? tokens.colorNeutralForeground3 : tokens.colorNeutralForeground1 }}>
            {value}
          </Body1Strong>
        }
      />
    </Card>
  );
}

function BreakdownTable({ title, rows }: { title: string; rows: Array<{ key: string; totalTokens: number }> }) {
  return (
    <div>
      <Body1Strong style={{ display: 'block', marginBottom: 4 }}>{title}</Body1Strong>
      <Table size="extra-small" style={breakdownTableStyle}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell style={nameColumnStyle}>Name</TableHeaderCell>
            <TableHeaderCell style={numericColumnStyle}>Tokens</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {rows.map(r => (
            <TableRow key={r.key}>
              <TableCell style={nameColumnStyle}><Caption1 style={modelNameStyle}>{r.key}</Caption1></TableCell>
              <TableCell style={numericColumnStyle}><Caption1>{fmtTokens(r.totalTokens)}</Caption1></TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
}

function DaySparkline({ rows }: { rows: Array<{ day: string; totalTokens: number }> }) {
  const maxTokens = Math.max(...rows.map(r => r.totalTokens), 1);
  return (
    <div>
      <Body1Strong style={{ display: 'block', marginBottom: 4 }}>By day</Body1Strong>
      <div style={{ display: 'flex', alignItems: 'flex-end', gap: 2, height: 40 }}>
        {rows.map(r => (
          <div
            key={r.day}
            title={`${r.day}: ${fmtTokens(r.totalTokens)} tok`}
            style={{
              flex: 1,
              background: tokens.colorBrandBackground,
              height: `${Math.max(4, (r.totalTokens / maxTokens) * 100)}%`,
              borderRadius: 2,
              opacity: 0.8,
            }}
          />
        ))}
      </div>
    </div>
  );
}
