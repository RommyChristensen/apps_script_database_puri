# AOG Puri Dashboard: launcher

A small installable app (PWA) that opens the Apps Script dashboard full screen on a phone, with the Connect Group icon.

## 1. Update Apps Script (once)

The dashboard must be allowed to show inside the launcher.

1. In Apps Script replace `Code.gs` and `JavaScript.html` with the versions in this repo.
2. **Deploy → Manage deployments → pencil → Version: New version → Deploy.** The `/exec` link stays the same.
3. Copy the **Web app URL** (ends in `/exec`).

> `ALLOWALL` lets other sites embed the dashboard too. It still needs a login, and every data call needs the session token, so the only exposure is clickjacking (a fake page framing the app).

## 2. Paste the URL

Open `docs/config.js` and put the URL between the quotes:

```js
window.AOG_LAUNCHER = {
  APP_URL: 'https://script.google.com/macros/s/AKfycb.../exec',
  NAME: 'AOG Puri Dashboard',
};
```

## 3. Publish on GitHub Pages

Repository **Settings → Pages → Build and deployment**: Source *Deploy from a branch*, Branch `main`, Folder `/docs`, Save.
After a minute the launcher is at `https://<your-user>.github.io/<repo>/`.

## 4. Install on the phone

- **Android (Chrome):** open the link, menu ⋮ → **Install app** (or *Add to Home screen*).
- **iPhone (Safari):** open the link, Share ⬆ → **Add to Home Screen** → Add.

Open it from the home screen icon.

## Good to know

- Google may show a thin banner at the top of Apps Script web apps; the launcher can't remove it.
- On iPhone, Safari can clear the login of framed pages, so you may have to sign in more often. If sign-in doesn't stick, use the **Buka di browser** button on the launcher's problem screen.
- After changing `config.js`, close and reopen the app (the launcher checks for the newest `config.js` whenever it is online).
- To ship a launcher update, change `CACHE` in `sw.js` (e.g. `aog-launcher-v2`).
- The icon is rendered from `docs/icons/logo-source.jpg` by `node dev/build-icons.js`. To change the logo, replace that file and rerun the script.
- The installed icon only updates after the app is reinstalled on some phones (remove it from the home screen and add it again).
