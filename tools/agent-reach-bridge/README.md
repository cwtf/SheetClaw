# Agent-Reach bridge

A localhost HTTP endpoint that lets the SheetClaw task pane read pages a direct
`fetch` cannot see — X, Reddit, YouTube, GitHub, Bilibili, XiaoHongShu.

```bash
npm run agent-reach-bridge
```

Then set **Settings → Search → Bridge URL** to `http://localhost:8788`.

## Why a bridge exists

The task pane is a WebView inside Excel: no child processes, no filesystem. It
cannot invoke a CLI. [Agent-Reach](https://github.com/Panniantong/Agent-Reach)
is CLI-only.

More than that, Agent-Reach is not a single reader CLI. `agent-reach install`
provisions a set of upstream tools (`yt-dlp`, `gh`, `rdt`, `bili`, …) and its
`SKILL.md` tells an agent which tool handles which platform — the reading itself
is done by those tools, called directly.

So this bridge is a **peer of that skill file**, not a wrapper around Agent-Reach:
it implements the same routing table over HTTP and returns normalized JSON.
`agent-reach doctor` still drives `/health`, since reporting which backends were
actually provisioned is exactly what that command does.

Run `agent-reach install` first if you want the platform readers available; the
generic web route needs nothing installed.

## Routes

| Route | Returns |
|---|---|
| `GET /health` | `agent-reach doctor` output, which tools are on PATH, and the routing table |
| `GET /read?url=` | `{ url, platform, tool, title, text, truncated }` |
| `GET /search?q=&limit=&platform=` | `{ query, platform, results: [{ title, url, snippet?, publishedAt? }] }` |

Status codes: `400` bad input, `501` platform routed but not dispatched, `502`
upstream tool failed.

## Platform coverage

| Platform | Tool | Dispatched |
|---|---|---|
| generic web | `r.jina.ai` (no CLI) | yes |
| youtube | `yt-dlp` | yes |
| github | `gh` | yes |
| reddit | `rdt` | no |
| twitter / x | cookie backend | no |
| bilibili | `bili` | no |
| xiaohongshu | `xhs` | no |

**Dispatched: no** means the platform is routed and health-checked but returns
`501` instead of running a command. Those readers' flag surfaces are not stable
enough to hardcode without verifying against a live install — guessing would
produce obscure failures rather than a clear "not wired up" message. To enable
one, add its `read`/`search` argv to `PLATFORMS` in `server.mjs` and flip
`implemented: true`.

## Flags

| Flag | Default | Meaning |
|---|---|---|
| `--port` | `8788` | Listen port |
| `--origin` | `https://localhost:3000` | Value sent in `Access-Control-Allow-Origin` |
| `--timeout` | `30000` | Per-command timeout in ms |

```bash
npm run agent-reach-bridge -- --port 9000 --origin https://localhost:3000
```

## Notes on safety

- Binds `127.0.0.1` only — never a public interface.
- Every dispatch uses `execFile` with an argument array, never a shell string,
  so a hostile URL cannot become shell injection.
- Only `http`/`https` URLs are accepted, and only GET is served.
- Output is capped at 20k characters per read.

It still shells out to tools that fetch attacker-influenced content, so treat it
as a dev-machine tool. Do not expose the port beyond localhost.
