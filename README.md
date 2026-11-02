# Spårbar demo

A runnable demo of Spårbar, the climate-report portal for transport customers, with AI features powered by
Claude (Anthropic). It runs on your laptop; the Anthropic API key stays on that laptop.

## Start (2 minutes)

You need [Node.js](https://nodejs.org) 20 or newer. There is nothing to install with npm.

```bash
git clone <this repo> && cd sparbar-demo
cp .env.example .env          # then paste the API key into .env
npm start                     # → Spårbar demo on http://127.0.0.1:3000
```

Open http://127.0.0.1:3000. The terminal says `AI: live` when the key works.
Without a key it says `AI: mock mode`: everything runs, and the AI answers are clearly marked demo texts.

## What the demo shows

| Where | What happens | AI? |
| --- | --- | --- |
| Översikt | Key figures, emissions by vehicle type and month, click a bar to trace it to order lines | No: fixed calculation |
| Data | Drop the customer's own Milj.xlsx; the portal reads its Rapport sheet and shows **Matchar Excel** | No |
| Kontroller | 10 automatic data checks; **Förklara med AI** explains one in plain Swedish with actions | Yes |
| Kundrapport | **Skriv sammanfattning med AI** writes the summary from the figures; editable; PDF/Excel/CSV/JSON export | Yes (text only) |
| AI-assistent | Ask questions about the report; answers use only the report's data | Yes |
| Kvitton | Drop a photo or PDF of a fuel receipt; Claude reads date, fuel, litres, amount for review | Yes |

**The AI never calculates the emissions.** The figures come from a fixed calculation that matches the customer's
Excel to the decimal (25,122.88 kg CO₂e for the sample). Claude reads, writes and explains around those figures, and a
person approves every report. Say this in the demo: it is what auditors and EU rules (CountEmissionsEU) require.

The presenter's script, in Swedish, is in [DEMO-SCRIPT.md](DEMO-SCRIPT.md).

## Data

- `data/demo-data.json`: **fictional** orders for three made-up customers (Demokund A, B, C). The vehicle-type
  table and its fuel-cost shares come from MLT's sample; there are no real orders or customers in this repository.
- MLT's own file is dropped in during the meeting. It is read in the browser and never written to disk or Git.
- When AI features are used, the report's **summary figures** (totals per vehicle type and month, checks) are sent to
  Anthropic with your API key; receipt images are sent when you read a receipt. No order lines are sent.
  Before using AI on MLT's real file, check that MLT is fine with that.

## Settings (`.env`)

| Setting | Default | Meaning |
| --- | --- | --- |
| `ANTHROPIC_API_KEY` | empty | Your Claude API key (see below). Empty = mock mode |
| `ANTHROPIC_MODEL` | `claude-sonnet-5-5` | Claude model used for all AI features |
| `PORT` | `3000` | Port |
| `HOST` | `127.0.0.1` | Only this computer. `0.0.0.0` makes it reachable on the network; then set a password |
| `DEMO_PASSWORD` | empty | Browser asks for it (any user name) when set |
| `MAX_AI_CALLS_PER_HOUR` | `120` | Hard cap to protect your API budget |
| `MAX_TOKENS` | `1500` | Longest answer per call |

## Getting a Claude API key

1. Sign in to the Claude Console at [platform.claude.com](https://platform.claude.com) and add credits under
   **Settings → Billing**. (Claude Max, Ultra and Team subscribers get monthly API credits: on claude.ai, go to
   **Settings → Billing → API credits → Link organization**.)
2. Go to **API keys** (platform.claude.com/settings/keys) → **Create key**. Name it `sparbar-demo`, pick an
   expiration (for example 7 days, so it switches itself off after the demo), and **select a workspace** (Default is
   fine). Copy the key straight away: it starts with `sk-ant-` and is shown only once.
3. Paste it into `.env` after `ANTHROPIC_API_KEY=` and restart `npm start`. The terminal should say `AI: live`.

## Keeping the key and the budget safe

- Create a **separate API key just for the demo**, set a monthly spend limit under **Settings → Billing → Spend
  limits** in the Console, and delete the key afterwards.
- Cost, measured with Claude Sonnet 5.5: one round of all four AI features (assistant question, check explanation,
  report summary, receipt photo) used about 10,400 input and 2,600 output tokens, roughly **$0.05**. The terminal
  prints the tokens of every call.
- `.env` is in `.gitignore`. Never commit it, and never paste the key into a chat or e-mail; send it through a
  password manager.
- The browser never sees the key: it calls `api/ai` on this server, which calls Anthropic.

## Check that everything works

```bash
npm test     # server test against a fake Anthropic API: streaming, JSON, wrong key, bad input, mock mode, file safety
```

## Layout

| Path | What |
| --- | --- |
| `server.js` | Web server and Claude proxy, no dependencies |
| `public/index.html` | The portal (built), `public/vendor/` bundled PDF/Excel libraries so it works offline |
| `src/` | Portal sources: calculation engine, report model, exports, app, styles, and `shim.js` (connects the portal to this server) |
| `data/demo-data.json` | Fictional demo data |
| `scripts/build.js` | `npm run build` rebuilds `public/index.html` after changing `src/` |
| `test/` | Smoke test and a fake Anthropic API |

## Troubleshooting

- **"The API key was not accepted"**: check `ANTHROPIC_API_KEY` in `.env`, then restart `npm start`. Expired keys
  also give this error.
- **"This API key is not scoped to a workspace…"**: create a new key with a workspace selected, or put the workspace
  ID (`wrkspc_…`) in `ANTHROPIC_WORKSPACE_ID`.
- **"You have reached your specified API usage limits"**: the spend limit is reached; raise it under Settings → Billing.
- **"Your credit balance is too low"**: add credits under Settings → Billing.
- **"More than N AI calls in the last hour"**: raise `MAX_AI_CALLS_PER_HOUR` or wait.
- **Port in use**: set `PORT=3001` in `.env`.
- **Fonts look plain offline**: the Roboto font loads from Google Fonts; everything else works offline.
