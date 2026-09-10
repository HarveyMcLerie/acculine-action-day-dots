# Action Day Dots — Setup

Standalone from the Sticky Notes add-in: separate repo, separate manifest, separate n8n workflow. The only thing reused is the existing "Microsoft Outlook account 3" OAuth2 credential — not any workflow logic.

## What it does
Adds a ribbon group with 5 one-click buttons (Mon/Tue/Wed/Thu/Fri) when reading an email. Clicking a day prepends a coloured dot emoji to that email's subject line (visible in the inbox list view), replacing any existing dot. Colours: Mon=🔴 Tue=🟠 Wed=🟡 Thu=🟢 Fri=🔵.

## 1. Host the add-in files
Push this folder's contents (`manifest.xml`, `commands.html`, `commands.js`, `icons/`) to a new GitHub repo, e.g. `harveymclerie.github.io/acculine-action-day-dots`, with GitHub Pages enabled — same pattern as the Sticky Notes add-in. If you use a different repo name or path, update the URLs in `manifest.xml` (search for `harveymclerie.github.io/acculine-action-day-dots`) to match.

## 2. Import the n8n workflow
1. In n8n, create a new workflow and import `action-day-dots-workflow.json`.
2. Open the **Update Subject (Graph API)** node and, in the Credential field, select your existing **Microsoft Outlook account 3** credential from the dropdown (replaces the placeholder ID in the file).
3. Activate the workflow. Confirm the webhook URL shown matches `https://acculine135.app.n8n.cloud/webhook/action-day-dots` — if n8n gives a different path, update `WEBHOOK_URL` in `commands.js` to match, and re-publish.

## 3. Sideload the add-in in Outlook
Same process as Sticky Notes — via Outlook Add-ins > My add-ins > Add a custom add-in > From file, pointing to the hosted `manifest.xml` URL, or have Brad Poole push it via Microsoft 365 admin deployment for a permanent install.

## Notes
- No task pane — buttons act instantly and dismiss (`event.completed()`), matching the one-click requirement.
- Only one dot is ever shown; clicking a new day strips any existing dot first.
- Nothing in this project touches the Sticky Notes manifest, add-in, or n8n workflow.
