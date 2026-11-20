# Release signing: one key for every Petak update

Android installs an update over an installed app only when both are signed with **the same key**.
Until now the GitHub Actions build had no key of yours, so it signed every APK with the debug key of
the machine that happened to run the build — a new, random key each time. That is why each new APK
had to be installed after uninstalling the old one.

You make one release key, once, and give it to GitHub as four secrets. From then on every build is
signed with it, and every later version installs over the previous one, keeping your settings,
recent files and unsaved session.

> **Keep the key file and its password safe, in two places (for example a USB stick and a password
> manager).** If you lose either, no future Petak can ever be installed over the current one: you
> would have to uninstall, and make a new key. Never put the key in the repository, in an issue or
> in a chat.

It takes about 10 minutes. You need a computer with Java **or** your Android phone with Termux.

---

## Step 1. Make the key

Pick a password of at least 12 characters and write it down now. Use **the same password** for the
key store and the key (newer Java ignores a separate key password anyway).

### On a computer (Windows, macOS or Linux)

`keytool` comes with Java. If `keytool -help` says "not found", install the free Temurin JDK 21 from
adoptium.net (or Android Studio, which contains it) and open a new terminal.

```bash
keytool -genkeypair -v -storetype PKCS12 -keystore petak-release.jks -alias petak \
        -keyalg RSA -keysize 4096 -validity 10000
```

It asks for the password twice, then for your name, organisation, city, and so on — you may press
Enter for every one of those (Android does not check them) — and finally "Is ... correct?": type
`yes`. The file `petak-release.jks` is your key.

### On the phone, with Termux

1. Install **Termux** from F-Droid (the Play Store version is outdated).
2. In Termux:
   ```bash
   pkg update && pkg install openjdk-21
   termux-setup-storage          # allow access to your Downloads folder
   cd ~/storage/downloads
   keytool -genkeypair -v -storetype PKCS12 -keystore petak-release.jks -alias petak \
           -keyalg RSA -keysize 4096 -validity 10000
   ```
   Answer the questions as above. The key is now `Downloads/petak-release.jks`.

### With Android Studio

Build → Generate Signed App Bundle or APK → APK → Next → "Create new…". Key store path:
`petak-release.jks`, the same password twice for the store and the key, alias `petak`, validity
25 years, any name in "First and Last Name". OK, then cancel the wizard: the key file is made.

---

## Step 2. Turn the key into text

GitHub secrets hold text, so the key file is converted to base64 (a long line of letters).

| Where | Command | Result |
|---|---|---|
| Linux, Termux | `base64 -w0 petak-release.jks > petak-release.b64` | `petak-release.b64` |
| macOS | `base64 -i petak-release.jks -o petak-release.b64` | `petak-release.b64` |
| Windows (PowerShell) | `[Convert]::ToBase64String([IO.File]::ReadAllBytes("$PWD\petak-release.jks")) \| Set-Content petak-release.b64` | `petak-release.b64` |

Open `petak-release.b64` in a text editor and copy **all** of it (about 5,000 characters, one line).
In Termux: `termux-clipboard-set < petak-release.b64` needs the Termux:API app; otherwise open the
file from Downloads in any text editor app.

---

## Step 3. Add four secrets on GitHub

On github.com (works in the phone's browser too): open the **Petak** repository → **Settings** →
**Secrets and variables** → **Actions** → **New repository secret**. Add these four, one at a time
(names exactly as written, upper case):

| Name | Value |
|---|---|
| `KEYSTORE_BASE64` | the whole text from `petak-release.b64` |
| `KEYSTORE_PASSWORD` | your password |
| `KEY_ALIAS` | `petak` |
| `KEY_PASSWORD` | your password (the same one) |

Secrets cannot be read back, not even by you; GitHub only uses them inside builds. Delete
`petak-release.b64` afterwards (keep `petak-release.jks` and its backup).

---

## Step 4. Build

Open **Actions** → **Petak Android APK** → **Run workflow**, choose the branch, **Run workflow**
(or push any commit). The build job now:

1. checks the key: a wrong password or alias stops the build with a clear message instead of
   making an APK with the wrong key;
2. signs the release *and* the debug APK with your key;
3. prints the key's fingerprint (SHA-256) in the run's summary and puts it in `SIGNING.txt` next to
   the APKs. **It must be the same for every build from now on.** If it ever changes, the secrets
   were changed.

When it is green, download the `Petak-apks` artifact (bottom of the run page), unzip it, and install
`Petak-v<version>.apk`.

A build without the secrets still works (forks, pull requests), but it says so in red in the
summary, and a version tag (`v1.3.0`) refuses to publish a release without the key.

---

## Step 5. The one-time switch on the phone

The Petak on your phone today is signed with one of the old random keys, so Android refuses the
first APK with your key ("App not installed" / "package conflicts with an existing package").
This happens **once**:

1. In Petak, save every open workbook to a file (File → Save as) — uninstalling deletes the app's
   own storage: recent files list, autosaved session, settings. Your saved `.xlsx` files stay.
2. Uninstall Petak (long-press the icon → App info → Uninstall).
3. Install the new APK.

Every later version installs over this one directly, keeping everything.

---

## Checks

* Fingerprint of an APK on a computer:
  `apksigner verify --print-certs Petak-v1.3.0.apk` (in the Android SDK's `build-tools`), or
  `keytool -printcert -jarfile Petak-v1.3.0.apk`. Compare with the SHA-256 in the run summary.
* Fingerprint of your key file: `keytool -list -v -keystore petak-release.jks -alias petak`.
* Local builds: `android/keystore.properties` (see README) uses the same key file; it is ignored by
  git, as are `*.jks`, `*.keystore` and `*.b64`.
