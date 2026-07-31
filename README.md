# Draft Reminder — Outlook Add-in

Checks your Drafts folder and notifies you (via a browser/OS notification)
when a draft has sat unsent for over an hour. Runs inside Outlook itself,
not on a server — meaning it only works while Outlook is open.

## What's here

| File | Purpose |
|---|---|
| `manifest.xml` | Add-in manifest (classic XML format) |
| `commands.html` / `commands.js` | Background runtime — hourly Drafts check via EWS |
| `taskpane.html` / `taskpane.js` | UI panel — enable notifications, check on demand |
| `assets/icon-*.png` | Placeholder icons (swap for your own branding if you like) |

## Before this will run anywhere

1. **Generate a real GUID** for `<Id>` in `manifest.xml` (currently a placeholder). Any GUID generator works — this just needs to be unique.
2. **Host the files somewhere over HTTPS.** Outlook loads the manifest's URLs live; nothing here works from a local file:// path. Options:
   - Quick local testing: use the official Yeoman generator (`npm install -g yo generator-office`, then `yo office`) which scaffolds a dev server with a trusted local HTTPS cert and can import these files.
   - Real hosting: any static HTTPS host — GitHub Pages, Azure Static Web Apps, Netlify, etc.
3. **Find-and-replace `REPLACE-WITH-YOUR-HOST`** in `manifest.xml` with your actual hosting domain, for every `<AppDomain>`, `IconUrl`, `SourceLocation`, and `bt:Url` entry.

## Sideloading (testing) in New Outlook

1. New Outlook → **Settings (gear icon)** → **General** → **Manage add-ins**, or from the ribbon **Apps** → **More Apps** → **Add custom add-in** → **Add from file**
2. Upload your edited `manifest.xml`
3. Compose a new message once — this fires `OnNewMessageCompose`, which starts the background hourly timer (see caveats below on why it's tied to that event)
4. Open the add-in's task pane (icon in the ribbon) and click **Enable notifications** — this has to happen once per browser/Outlook profile, since notification permission can only be granted from a page with user interaction, not the hidden background page

## Real limitations — please read before relying on this

- **Outlook must be open.** Unlike a cloud-based scheduler, there's no server-side component here. Close Outlook, and checks stop.
- **The hourly timer isn't a platform-guaranteed feature.** I've wired it to Outlook's `lifetime="long"` runtime, which is the documented way to keep a background script alive, tied to a real event (`OnNewMessageCompose`) so it actually starts without you needing to open the task pane first. But Outlook/WebView2 can still suspend or reload that runtime under memory pressure — there's no official SLA that a `setInterval` survives indefinitely. Treat missed checks as possible, not a bug.
- **`ReadWriteMailbox` permission** is required for the EWS call that lists Drafts folder items. This works fine for sideloading/internal org deployment. If you ever wanted to publish this to AppSource, that permission level triggers additional Microsoft review.
- **Notifications require one manual step** (clicking "Enable notifications" in the task pane) because browsers don't allow permission prompts from hidden background pages — this is a browser security rule, not something to work around.

## Honest comparison to the Power Automate version

The cloud flow runs on Microsoft's servers regardless of whether your PC or
Outlook is on, and its hourly trigger is a real guaranteed schedule. This
add-in trades that reliability for running entirely client-side with no
Azure AD app registration or cloud flow to maintain — reasonable if you're
always working with Outlook open, less reliable if you want a guarantee
that doesn't depend on the app being open.
