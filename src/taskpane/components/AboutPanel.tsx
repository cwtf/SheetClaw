import { Body1, Caption1, Title3, tokens } from '@fluentui/react-components';

const GitHubIcon = () => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    width="20"
    height="20"
    viewBox="0 0 24 24"
    aria-label="GitHub"
    fill="currentColor"
  >
    <path d="M12 0C5.373 0 0 5.373 0 12c0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.305 3.492.998.108-.776.417-1.305.76-1.605-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576C20.562 21.8 24 17.302 24 12 24 5.373 18.627 0 12 0z" />
  </svg>
);

export default function AboutPanel() {
  return (
    <div style={{
      padding: '24px 16px',
      display: 'flex',
      flexDirection: 'column',
      boxSizing: 'border-box',
      gap: 16,
      height: '100%',
      overflowY: 'auto',
    }}>
      <Title3>
        <span aria-hidden="true" style={{ marginRight: 8 }}>🦞</span>
        SheetClaw
      </Title3>

      <Body1 style={{ color: tokens.colorNeutralForeground2, lineHeight: '1.5' }}>
        SheetClaw is an agentic chat interface for Excel workbooks. Ask questions, read and
        write cells, create charts and pivot tables — all through a conversational
        interface backed by your choice of LLM provider.
      </Body1>

      <Body1 style={{ color: tokens.colorNeutralForeground2, lineHeight: '1.5' }}>
        Supports Ollama (local), OpenAI, Anthropic, DeepSeek, Groq, and any
        OpenAI-compatible endpoint including OpenRouter.
      </Body1>

      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        Created by Christopher Wong ·{' '}
        <a
          href="https://iconlearning.com.my"
          target="_blank"
          rel="noreferrer"
          style={{ color: 'inherit' }}
        >
          Icon Learning &amp; Development Sdn Bhd
        </a>
      </Caption1>

      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        Licensed under{' '}
        <a
          href="https://github.com/cwtf/SheetClaw/blob/master/LICENSE.md"
          target="_blank"
          rel="noreferrer"
          style={{ color: 'inherit' }}
        >
          PolyForm Noncommercial 1.0.0
        </a>
        {' '}— free for personal and noncommercial use.
      </Caption1>

      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        Commercial use requires a separate license:{' '}
        <a
          href="mailto:christopher.wong@iconlearning.com.my"
          style={{ color: 'inherit' }}
        >
          christopher.wong@iconlearning.com.my
        </a>
      </Caption1>

      <a
        href="https://cwtf.github.io/SheetClaw/privacy.html"
        target="_blank"
        rel="noreferrer"
        style={{
          display: 'block',
          fontSize: 12,
          color: tokens.colorNeutralForeground3,
          textDecoration: 'none',
          marginTop: 4,
        }}
      >
        Privacy Policy
      </a>

      <a
        href="https://github.com/cwtf/SheetClaw"
        target="_blank"
        rel="noreferrer"
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: 8,
          color: tokens.colorNeutralForeground1,
          textDecoration: 'none',
          marginTop: 8,
        }}
        onMouseEnter={e => (e.currentTarget.style.color = tokens.colorBrandForeground1)}
        onMouseLeave={e => (e.currentTarget.style.color = tokens.colorNeutralForeground1)}
      >
        <GitHubIcon />
        <Caption1>github.com/cwtf/SheetClaw</Caption1>
      </a>

      <Caption1
        style={{
          color: tokens.colorNeutralForeground3,
          marginTop: 'auto',
          paddingTop: 8,
        }}
      >
        Version {__APP_VERSION__}
      </Caption1>
    </div>
  );
}
