# AI Daily Newsletter

Free, Gmail-first, invite-only AI news digest built with Google Sheets, Google Apps Script, and Gmail.

This project is designed for a small beta before you own a domain. It drafts a daily AI newsletter from RSS feeds, emails you a preview, waits for approval in the sheet, and only then sends personalized Gmail messages with unsubscribe links.

## What You Get

- Google Sheet-backed subscriber list, sources list, candidate feed, issues table, send log, and config table.
- Daily source fetch and draft generation from AI news, product, research, and tools feeds.
- Approval-gated sending: no subscriber email is sent unless an issue row is marked `approved`.
- Personalized sends through `GmailApp.sendEmail` so every subscriber gets a unique unsubscribe link.
- Apps Script web app endpoint for unsubscribe tokens.
- Gmail-friendly HTML plus plain-text fallback.
- Quota guard using `MailApp.getRemainingDailyQuota()`.

## Files

- `Code.gs`: the Google Apps Script implementation.
- `appsscript.json`: Apps Script manifest with timezone, runtime, and required scopes.

## Setup

1. Create a new Google Sheet named `AI Daily Newsletter`.
2. Open `Extensions > Apps Script`.
3. Paste the contents of `Code.gs` into the Apps Script editor.
4. In Apps Script, open `Project Settings` and enable `Show "appsscript.json" manifest file in editor`.
5. Replace the manifest with the contents of `appsscript.json`.
6. Save the Apps Script project.
7. Run `setupNewsletter` from the Apps Script editor and approve the requested permissions.
8. Return to the Google Sheet and reload it. You should see a `Newsletter` menu.

## First Configuration

After running `setupNewsletter`, open the `Config` tab and fill these rows:

- `admin_email`: your email address for previews and error alerts.
- `reply_to`: the address subscribers can reply to.
- `newsletter_name`: default is `AI Daily`.
- `unsubscribe_base_url`: leave blank until you deploy the web app in the next step.

Add invite-only subscribers in the `Subscribers` tab:

- `email`: required.
- `name`: optional.
- `status`: use `active`.
- `source`: use `invite`.
- `consent_date`: date they agreed to receive the beta.
- `confirmed`: use `TRUE`.

If you leave `status`, `source`, `confirmed`, or `unsubscribe_token` blank, `setupNewsletter` fills safe defaults.

## Deploy Unsubscribe Web App

1. In Apps Script, click `Deploy > New deployment`.
2. Choose `Web app`.
3. Set `Execute as` to `Me`.
4. Set `Who has access` to `Anyone`.
5. Deploy and approve permissions.
6. Copy the web app URL.
7. Paste it into the `Config` tab as `unsubscribe_base_url`.

Subscribers will receive links like:

```text
https://script.google.com/macros/s/.../exec?token=SUBSCRIBER_TOKEN
```

When clicked, the script marks that subscriber as `unsubscribed`.

## Daily Automation

Run `installDailyTriggers` once from Apps Script, or use the Google Sheet menu:

```text
Newsletter > Install daily triggers
```

This installs two free Apps Script time triggers:

- Around 7 AM Eastern: `runMorningDraft` fetches sources, creates a draft issue, and emails you a preview.
- Around 8 AM Eastern: `sendApprovedIssue` sends only if the latest issue is marked `approved`.

Google may run time-based triggers within a window around the requested time, so treat the timing as approximate.

## Daily Workflow

1. Check your preview email.
2. Open the `Issues` tab.
3. Review the latest row's `subject`, `html_body`, and `text_body`.
4. If it looks good, change `status` from `draft` to `approved`.
5. Let the 8 AM trigger send it, or run:

```text
Newsletter > Send approved issue now
```

If the issue is not approved, the send function logs `held_not_approved` in `SendLog` and exits.

## Testing Checklist

Before using real subscribers:

1. Add 2-3 of your own Gmail addresses to `Subscribers`.
2. Run `Newsletter > Fetch sources + draft preview`.
3. Confirm you receive the preview.
4. Leave the issue as `draft`, run `Newsletter > Send approved issue now`, and confirm no subscriber email sends.
5. Change the issue status to `approved`.
6. Run `Newsletter > Send approved issue now`.
7. Confirm each test Gmail account receives a clean email.
8. Click an unsubscribe link and confirm the matching subscriber becomes `unsubscribed`.
9. Check `SendLog` for sent rows and errors.

## Quotas And Limits

This v1 is intentionally free and small-list only.

- Consumer Gmail / Apps Script accounts commonly have a 100-recipient daily send quota.
- Google Workspace accounts usually have higher quotas, but you should still start small.
- The script sends one personalized email per subscriber to support unique unsubscribe links.
- The script refuses to send if the needed recipient count exceeds `MailApp.getRemainingDailyQuota()`.
- Keep the beta under roughly 75-90 active subscribers so you have quota buffer for previews and retries.

## Upgrade Path

When you buy a domain or want a public signup flow:

1. Buy a domain and use a dedicated subdomain such as `news.example.com`.
2. Move sending to Resend, Buttondown, Beehiiv, Mailchimp, or another newsletter provider.
3. Configure SPF, DKIM, and DMARC.
4. Import subscribers with consent records.
5. Add Google Postmaster Tools after domain-based sending starts.
6. Keep this sheet as the editorial/source pipeline, or replace it with a small web app.

