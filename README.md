# SheetClaw

SheetClaw is a personal-use Excel task pane add-in that brings an agentic chat interface into a workbook. It can inspect workbook context, read and write ranges, create charts and pivot tables, and track estimated LLM usage while routing requests through your chosen model provider.

The add-in is built for local sideloading in Excel with Office.js, React, TypeScript, Vite, Fluent UI, and Zustand.

## Features

- Chat-first workflow for asking questions about the active Excel workbook.
- Office.js tool layer for reading workbook state, selections, sheets, ranges, named ranges, charts, and pivot tables.
- Confirm-before-write flow for mutating tools such as range writes, clears, chart changes, and pivot operations.
- Snapshot support before mutating operations so applied changes can be undone where supported.
- Provider adapters for Ollama, OpenAI, Anthropic, OpenRouter, DeepSeek, Groq, Mistral, Together AI, Kimi, GLM, Qwen, Llama, and generic OpenAI-compatible endpoints.
- OpenRouter OAuth sign-in flow plus API-key fallback.
- Usage tracking with bundled pricing data, rolling local history, dashboard summaries, and CSV export support.
- Host-workbook scoped operation, matching the Office task pane runtime model.

## Sideloading

### Method 0: Automated script (recommended for development)

With the dev server running, execute:

```powershell
npm run sideload
```

This runs `office-addin-debugging start manifest.xml desktop --app excel` and handles registration automatically.

### Method 1: Shared Folder Catalog

The standard manual approach for Excel 2019 and later when the automated script is not available.

**Step 1 — Set up a shared folder**

1. Create a folder on your PC, e.g. `C:\MyAddins`.
2. Right-click the folder → **Properties** → **Sharing** tab → **Share** → set permissions and note the network path (e.g. `\\YourPC\MyAddins`).
3. Copy `manifest.xml` from this repo into that folder.

**Step 2 — Register the folder as a trusted catalog in Excel**

1. Open Excel → **File** → **Options**.
2. Go to **Trust Center** → **Trust Center Settings**.
3. Click **Trusted Add-in Catalogs**.
4. Paste the network path (e.g. `\\YourPC\MyAddins`) into the **Catalog Url** field.
5. Click **Add catalog** → check **Show in Menu** → click **OK**.
6. Restart Excel.

**Step 3 — Load the add-in**

1. In Excel, go to **Insert** → **Get Add-ins** (or **My Add-ins**).
2. Select the **Shared Folder** tab, then select **SheetClaw** and click **Add**.

### Method 2: Upload manifest directly

If your Excel build exposes an upload option:

1. Go to **Insert** → **Get Add-ins**.
2. Click **My Add-ins** → **Upload My Add-in**.
3. Browse to `manifest.xml` in this repo and select it.

## Using The Add-In

1. Open SheetClaw from Excel's ribbon after sideloading.
2. Go to Settings and choose a provider.
3. Configure the provider base URL, model, and API key or OpenRouter sign-in.
4. Test the connection.
5. Return to Chat and ask SheetClaw to inspect or modify the active workbook.

Mutating operations require confirmation before they are applied to the workbook.

## Provider Notes

- Ollama: local provider, no API key required by default.
- OpenRouter: configured under the OpenRouter tab with OAuth sign-in and API-key fallback.
- Other API: API-key based providers including OpenAI, Anthropic, DeepSeek, Groq, Mistral, Together AI, Kimi, GLM, Qwen, and Llama.
- Generic OpenAI-compatible endpoints can be configured by changing the base URL and model.

Credentials are stored locally by the add-in, encrypted at rest with AES-GCM via Web Crypto; the encryption key is a non-extractable `CryptoKey` kept in IndexedDB. This protects keys from storage dumps and disk inspection, but a same-origin script compromise could still use the key, so treat this as a personal development tool unless you harden credential storage further (for example an OS credential vault) for broader use.

## License

SheetClaw is licensed under the [PolyForm Noncommercial License 1.0.0](LICENSE.md) —
free to use, modify, and share for noncommercial purposes: personal productivity,
study, hobby projects, and use by charities, schools, public research bodies, and
government institutions.

Commercial use — including paid training courses and workshops (HRD Corp-claimable
programmes included), consultancy or client-facing engagements, and bundling
SheetClaw into a commercial product or service — is not permitted under this
license and requires a separate commercial license. For the avoidance of doubt,
we do not consider for-profit training providers to be "educational institutions"
under the license.

Commercial licensing enquiries: christopher.wong@iconlearning.com.my
