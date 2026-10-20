# OKV Technology Consults back office - setup (Google Apps Script only)

Everything (forms, emails, alerts, customers, payments, activation codes, admin panel) runs inside one Google Sheet and one Apps Script project. No servers, no other services (Paystack is optional).

## 1. Create the project (10 minutes)
1. Go to https://script.google.com > **New project**. Delete the default code. The data spreadsheet is created by the setup function (no manual sheet needed).
2. Create two files and paste the contents:
   - `Code.gs`  (script file)
   - `Admin.html`  (**+ > HTML**, name it exactly `Admin`)
3. In `Code.gs` change these two lines near the top, then save:
   - `INITIAL_ADMIN_PASSWORD` - your first admin-panel password (you are forced to change it at first login).
   - `DEFAULT_CODE_SECRET` - a **long secret you invent** (30+ characters). It must be identical to `OKV_LICENSE_SECRET` when you build the Setup.exe. Never change it after you have sent codes.
4. Select the function **setup** in the toolbar and press **Run**. Accept the permission prompts (Sheets, Drive, Gmail/Mail, external requests). This creates all sheets, generates **1,200 activation codes**, and installs a 10-minute trigger that confirms Paystack payments.

## 2. Publish it
1. **Deploy > New deployment > Web app**.
2. Execute as: **Me**. Who has access: **Anyone**.
3. Copy the **Web app URL** (ends in `/exec`). This is your `endpoint`.
4. Open the URL in a browser: you see the **OKV admin panel**. Sign in and change the password.

## 3. Connect everything
| Where | What to paste |
|---|---|
| `Website/assets/js/config.js` | `endpoint: "<web app URL>"` |
| Build kit (`set OKV_ENDPOINT=...`) | the same URL (in-app activation form) |
| Build kit (`set OKV_LICENSE_SECRET=...`) | the same secret as `DEFAULT_CODE_SECRET` |

In the admin panel > **Settings & integrations** set: price, trial download link (upload `OKV-CBT-Setup.zip` to Google Drive > share "Anyone with the link" > use a direct link, or host it on your website), your **bank account details** (shown on the Pricing page for manual transfers), alert email, signature and templates.

## 4. Paystack (optional but recommended)
1. Create a Paystack business account, get the **Secret key** (Settings > API Keys).
2. Admin panel > Settings > Paystack: paste the secret key, set **Callback URL** to `https://yourdomain/pricing.html`, enable, press **Test Paystack key**.
3. How it works: customer pays on Paystack and returns to pricing.html; the page asks the script to verify the reference with Paystack; on success the activation code is emailed instantly. If the customer closes the browser, a background check every 10 minutes completes it. Use **Payments > Verify** for anything stuck.

## 5. Daily use
- Every submission (free trial, purchase, activation request, installation request, contact) creates an **alert**: emailed to you, shown with a bell in the panel, optionally sent to Telegram and to your webhook.
- **Submissions**: open, reply, mark handled, send code. **Payments**: approve a bank-transfer payment and the activation code is emailed at once. **Activation codes**: 1,000+ ready; send one to anyone, generate 1,000 more at any time.
- **Messages**: email many customers (Google limit: 100/day on free Gmail, 1,500 on Workspace) or open WhatsApp / SMS / call links one by one.
- **Tools**: generate the forgot-password reset code for an organisation (they send you their Install ID).

## Limits and notes
- Receipts are saved in your Google Drive folder "OKV CBT Receipts" (private).
- Activation codes are checked offline by the app with the shared secret, so they work without internet. One code is meant for one organisation; the sheet records who got it and when it was activated.
- Apps Script web apps cannot read Paystack webhook signatures, so this design verifies by calling Paystack directly (safer).
