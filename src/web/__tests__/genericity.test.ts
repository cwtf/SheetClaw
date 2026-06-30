import { describe, expect, it } from 'vitest';
import { PROVIDER_URL_HOST_ALLOWLIST } from '../providers';

declare const process: { cwd(): string };
declare function require(id: string): {
  readFileSync?: (path: string, encoding: string) => string;
  readdirSync?: (path: string) => string[];
  statSync?: (path: string) => { isDirectory(): boolean };
  join?: (...parts: string[]) => string;
};

const fs = require('fs');
const path = require('path');
const readFileSync = fs.readFileSync as (path: string, encoding: string) => string;
const readdirSync = fs.readdirSync as (path: string) => string[];
const statSync = fs.statSync as (path: string) => { isDirectory(): boolean };
const join = path.join as (...parts: string[]) => string;

const SRC_ROOT = join(process.cwd(), 'src');

const DEMO_SCENARIO_STRINGS = [
  'data.gov.my',
  'dosm.gov.my',
  'malaysia',
  'population dataset',
  'by state',
  'by age/sex',
  'national trend',
];

const PRE_EXISTING_CONFIG_HOSTS = [
  'localhost',
  'api.openai.com',
  'platform.openai.com',
  'api.anthropic.com',
  'console.anthropic.com',
  'openrouter.ai',
  'api.deepseek.com',
  'platform.deepseek.com',
  'api.groq.com',
  'console.groq.com',
  'api.mistral.ai',
  'console.mistral.ai',
  'api.together.ai',
  'api.moonshot.ai',
  'platform.moonshot.ai',
  'api.z.ai',
  'z.ai',
  'dashscope-intl.aliyuncs.com',
  'bailian.console.aliyun.com',
  'api.llama.com',
  'llama.developer.meta.com',
  'generativelanguage.googleapis.com',
  'aistudio.google.com',
  'api.cerebras.ai',
  'cloud.cerebras.ai',
  'api.cloudflare.com',
  'dash.cloudflare.com',
  'router.huggingface.co',
  'huggingface.co',
  'www.w3.org',
  'cwtf.github.io',
  'github.com',
  'iconlearning.com.my',
  'public.example',
  'finance.example',
  'weather.example',
  'sports.example',
  '127.0.0.1',
  '192.168.1.5',
];

describe('web access genericity guard', () => {
  it('does not contain failed-demo domain, dataset, geography, or canned menu strings', () => {
    const haystack = sourceFiles().map(file => readFileSync(file, 'utf8').toLowerCase()).join('\n');
    for (const needle of DEMO_SCENARIO_STRINGS) {
      expect(haystack, `unexpected demo-specific string: ${needle}`).not.toContain(needle);
    }
  });

  it('does not add undeclared hostnames to source string literals', () => {
    const allowed = new Set([...PROVIDER_URL_HOST_ALLOWLIST, ...PRE_EXISTING_CONFIG_HOSTS]);
    const hits: string[] = [];
    for (const file of sourceFiles()) {
      const text = readFileSync(file, 'utf8');
      for (const host of findHostnames(text)) {
        if (!allowed.has(host)) hits.push(`${file}: ${host}`);
      }
    }
    expect(hits).toEqual([]);
  });
});

function sourceFiles(dir = SRC_ROOT): string[] {
  const out: string[] = [];
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const stat = statSync(full);
    if (stat.isDirectory()) {
      out.push(...sourceFiles(full));
    } else if (/\.(ts|tsx)$/.test(entry) && entry !== 'genericity.test.ts') {
      out.push(full);
    }
  }
  return out;
}

function findHostnames(text: string): string[] {
  const hosts = new Set<string>();
  const urlPattern = /https?:\/\/([^/ "'`<>)]+)/g;
  let match: RegExpExecArray | null;
  while ((match = urlPattern.exec(text))) {
    try {
      hosts.add(new URL(`${text.slice(match.index, urlPattern.lastIndex)}`).hostname.toLowerCase());
    } catch {
      hosts.add(match[1].split(':')[0].toLowerCase());
    }
  }
  return [...hosts];
}
