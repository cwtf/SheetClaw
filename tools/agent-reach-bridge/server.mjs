#!/usr/bin/env node
// Agent-Reach bridge: a localhost HTTP endpoint the SheetClaw task pane can call.
//
// Why this exists: the task pane is a WebView inside Excel. It has no child
// processes and no filesystem, so it cannot invoke a CLI. Agent-Reach is
// CLI-only - and it is not even a single CLI. `agent-reach install` provisions
// a set of upstream readers (yt-dlp, gh, ...) and its SKILL.md tells an agent
// which one handles which platform; the reading itself is done by those tools.
//
// So this bridge is a peer of that skill file rather than a wrapper around
// Agent-Reach: it implements the same routing table over HTTP and returns
// normalized JSON. `agent-reach doctor` still drives /health, because knowing
// which backends Agent-Reach actually provisioned is exactly what it reports.
//
//   node tools/agent-reach-bridge/server.mjs --port 8788 --origin https://localhost:3000
//
// Binds to 127.0.0.1 only. Every dispatch uses execFile with an argument array,
// never a shell string, so a hostile URL cannot become a shell injection.

import { createServer } from 'node:http';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const run = promisify(execFile);

const DEFAULTS = {
  port: 8788,
  origin: 'https://localhost:3000',
  timeoutMs: 30_000,
  maxChars: 20_000,
};

// ── Platform routing ───────────────────────────────────────────────────────
// Mirrors the table Agent-Reach's SKILL.md describes. `read`/`search` return an
// argv array, or null when that platform has no handler for that verb. Tools
// marked `implemented: false` are routed and reported but not dispatched: the
// flag surface of those readers is not stable enough to hardcode blind, so they
// return 501 with the tool name rather than guessing and failing obscurely.

const READER_PROXY = 'https://r.jina.ai/';

const PLATFORMS = [
  {
    id: 'youtube',
    tool: 'yt-dlp',
    implemented: true,
    hosts: ['youtube.com', 'www.youtube.com', 'm.youtube.com', 'youtu.be'],
    read: url => ['--dump-json', '--skip-download', '--no-warnings', url],
    search: (q, n) => ['--dump-json', '--flat-playlist', '--skip-download', '--no-warnings', `ytsearch${n}:${q}`],
    parseRead: stdout => {
      const v = JSON.parse(stdout);
      return {
        title: v.title ?? '',
        text: [
          v.title && `Title: ${v.title}`,
          v.uploader && `Channel: ${v.uploader}`,
          v.upload_date && `Uploaded: ${v.upload_date}`,
          Number.isFinite(v.duration) && `Duration: ${v.duration}s`,
          Number.isFinite(v.view_count) && `Views: ${v.view_count}`,
          v.description && `\n${v.description}`,
        ].filter(Boolean).join('\n'),
      };
    },
    parseSearch: stdout => ndjson(stdout).map(v => ({
      title: v.title ?? v.id ?? '',
      url: v.webpage_url ?? (v.id ? `https://www.youtube.com/watch?v=${v.id}` : ''),
      snippet: v.description ?? v.uploader ?? undefined,
    })),
  },
  {
    id: 'github',
    tool: 'gh',
    implemented: true,
    hosts: ['github.com', 'www.github.com'],
    read: url => {
      const path = new URL(url).pathname.split('/').filter(Boolean);
      if (path[2] === 'issues' && path[3]) return ['issue', 'view', url, '--comments'];
      if (path[2] === 'pull' && path[3]) return ['pr', 'view', url, '--comments'];
      if (path.length >= 2) return ['repo', 'view', `${path[0]}/${path[1]}`];
      throw new BadRequest('GitHub URL must point at a repo, issue, or pull request.');
    },
    search: (q, n) => ['search', 'repos', q, '--limit', String(n), '--json', 'fullName,url,description,updatedAt'],
    parseRead: stdout => ({ title: stdout.split('\n')[0] ?? '', text: stdout }),
    parseSearch: stdout => JSON.parse(stdout).map(r => ({
      title: r.fullName ?? '',
      url: r.url ?? '',
      snippet: r.description ?? undefined,
      publishedAt: r.updatedAt ?? undefined,
    })),
  },
  // Routed and health-checked, but not dispatched - see `implemented` above.
  { id: 'reddit',      tool: 'rdt',  implemented: false, hosts: ['reddit.com', 'www.reddit.com', 'old.reddit.com'] },
  { id: 'twitter',     tool: 'twitter-backend', implemented: false, hosts: ['x.com', 'twitter.com', 'www.x.com', 'www.twitter.com'] },
  { id: 'bilibili',    tool: 'bili', implemented: false, hosts: ['bilibili.com', 'www.bilibili.com', 'b23.tv'] },
  { id: 'xiaohongshu', tool: 'xhs',  implemented: false, hosts: ['xiaohongshu.com', 'www.xiaohongshu.com'] },
];

class BadRequest extends Error {}
class Unsupported extends Error {}

function ndjson(stdout) {
  return stdout.split('\n').filter(l => l.trim()).map(l => JSON.parse(l));
}

function platformFor(hostname) {
  const host = hostname.toLowerCase().replace(/^www\./, '');
  return PLATFORMS.find(p => p.hosts.some(h => h.replace(/^www\./, '') === host)) ?? null;
}

// ── Dispatch ───────────────────────────────────────────────────────────────

async function dispatch(tool, args, cfg) {
  try {
    const { stdout } = await run(tool, args, {
      timeout: cfg.timeoutMs,
      maxBuffer: 8 * 1024 * 1024,
      windowsHide: true,
    });
    return stdout;
  } catch (e) {
    if (e.code === 'ENOENT') {
      throw new Unsupported(`${tool} is not on PATH. Run \`agent-reach install\` to provision it.`);
    }
    const detail = (e.stderr || e.message || '').toString().trim().split('\n').slice(0, 4).join(' ');
    throw new Error(`${tool} failed: ${detail}`);
  }
}

/** Generic web reads go through the keyless reader proxy - no CLI involved. */
async function readViaProxy(url, cfg) {
  const res = await fetch(`${READER_PROXY}${url}`, { headers: { Accept: 'text/plain' } });
  if (!res.ok) throw new Error(`reader proxy returned HTTP ${res.status}`);
  const text = await res.text();
  return { title: text.split('\n').find(l => l.trim()) ?? '', text, tool: 'jina-reader' };
}

async function handleRead(params, cfg) {
  const raw = params.get('url');
  if (!raw) throw new BadRequest('Missing "url".');

  let parsed;
  try {
    parsed = new URL(raw);
  } catch {
    throw new BadRequest('"url" must be an absolute URL.');
  }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    throw new BadRequest('"url" must be http or https.');
  }

  const platform = platformFor(parsed.hostname);

  if (!platform) {
    const out = await readViaProxy(raw, cfg);
    return { url: raw, platform: 'web', tool: out.tool, title: out.title, ...cap(out.text, cfg) };
  }
  if (!platform.implemented) {
    throw new Unsupported(
      `Reading ${platform.id} needs the ${platform.tool} backend, which this bridge routes but does not dispatch. ` +
      `Run \`agent-reach doctor\` to check it, then add its argv to PLATFORMS in server.mjs.`
    );
  }

  const stdout = await dispatch(platform.tool, platform.read(raw), cfg);
  const parsedOut = platform.parseRead(stdout);
  return { url: raw, platform: platform.id, tool: platform.tool, title: parsedOut.title, ...cap(parsedOut.text, cfg) };
}

async function handleSearch(params, cfg) {
  const query = (params.get('q') ?? '').trim();
  if (!query) throw new BadRequest('Missing "q".');
  const limit = Math.min(Math.max(parseInt(params.get('limit') ?? '5', 10) || 5, 1), 25);

  const requested = params.get('platform');
  const platform = requested
    ? PLATFORMS.find(p => p.id === requested)
    : PLATFORMS.find(p => p.id === 'youtube');

  if (!platform) {
    throw new BadRequest(`Unknown platform "${requested}". Known: ${PLATFORMS.map(p => p.id).join(', ')}.`);
  }
  if (!platform.implemented || !platform.search) {
    throw new Unsupported(`Searching ${platform.id} is routed but not dispatched by this bridge.`);
  }

  const stdout = await dispatch(platform.tool, platform.search(query, limit), cfg);
  return { query, platform: platform.id, results: platform.parseSearch(stdout).slice(0, limit) };
}

async function handleHealth(cfg) {
  let doctor = '';
  let agentReach = true;
  try {
    doctor = await dispatch('agent-reach', ['doctor'], cfg);
  } catch (e) {
    agentReach = false;
    doctor = e.message;
  }

  const tools = {};
  for (const tool of [...new Set(PLATFORMS.filter(p => p.implemented).map(p => p.tool))]) {
    tools[tool] = await onPath(tool, cfg);
  }

  return {
    ok: true,
    agentReach,
    doctor: doctor.slice(0, 4000),
    tools,
    platforms: PLATFORMS.map(p => ({ id: p.id, tool: p.tool, dispatched: !!p.implemented })),
  };
}

async function onPath(tool, cfg) {
  try {
    await dispatch(tool, ['--version'], { ...cfg, timeoutMs: 5_000 });
    return true;
  } catch {
    return false;
  }
}

function cap(text, cfg) {
  const value = text ?? '';
  if (value.length <= cfg.maxChars) return { text: value, truncated: false };
  return { text: value.slice(0, cfg.maxChars), truncated: true };
}

// ── HTTP ───────────────────────────────────────────────────────────────────

function parseArgs(argv) {
  const cfg = { ...DEFAULTS };
  for (let i = 0; i < argv.length; i += 2) {
    const flag = argv[i];
    const value = argv[i + 1];
    if (flag === '--port') cfg.port = Number(value);
    else if (flag === '--origin') cfg.origin = value;
    else if (flag === '--timeout') cfg.timeoutMs = Number(value);
  }
  return cfg;
}

const cfg = parseArgs(process.argv.slice(2));

const server = createServer((req, res) => {
  const send = (status, body) => {
    res.writeHead(status, {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': cfg.origin,
      'Access-Control-Allow-Headers': 'Content-Type',
      'Cache-Control': 'no-store',
      Vary: 'Origin',
    });
    res.end(JSON.stringify(body));
  };

  if (req.method === 'OPTIONS') return send(204, {});
  if (req.method !== 'GET') return send(405, { error: 'Only GET is supported.' });

  const url = new URL(req.url, `http://127.0.0.1:${cfg.port}`);
  const route =
    url.pathname === '/health' ? handleHealth(cfg)
    : url.pathname === '/read' ? handleRead(url.searchParams, cfg)
    : url.pathname === '/search' ? handleSearch(url.searchParams, cfg)
    : null;

  if (!route) return send(404, { error: 'Not found. Routes: /health, /read?url=, /search?q=' });

  route.then(
    body => send(200, body),
    e => {
      if (e instanceof BadRequest) return send(400, { error: e.message });
      if (e instanceof Unsupported) return send(501, { error: e.message });
      return send(502, { error: e.message });
    }
  );
});

server.listen(cfg.port, '127.0.0.1', () => {
  process.stdout.write(
    `agent-reach bridge on http://127.0.0.1:${cfg.port} (allowing origin ${cfg.origin})\n`
  );
});
