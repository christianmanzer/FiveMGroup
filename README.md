# FiveM Group website

Static site for Cloudflare Pages. No build step.

- `index.html`, `employment/index.html`, `styles.css`, `assets/` : the site
- `functions/api/apply.js` : emails Employment form submissions to hello@fivem.group (Cloudflare Pages Function)

## Deploy (GitHub)

1. Put this folder in a new GitHub repository.
2. Cloudflare: Workers & Pages > Create > Pages > Connect to Git, and pick the repo.
3. Leave the build command empty and set the build output directory to `/`.

The `functions/` folder is picked up automatically, and every push redeploys the site.

## Make the Employment form send email (Microsoft 365)

Applications are sent from your own Microsoft 365 mailbox through Microsoft Graph, so there is no extra service or fee. One-time setup by a Microsoft 365 admin:

1. Go to entra.microsoft.com > Identity > Applications > App registrations > New registration. Name it "FiveM website form", leave it single-tenant, and register it.
2. On the app's Overview page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
3. API permissions > Add a permission > Microsoft Graph > **Application permissions** > `Mail.Send`. Then click **Grant admin consent**.
4. Certificates & secrets > New client secret. Copy the secret **Value** right away (it is only shown once). Secrets expire (24 months at most), so put a reminder on the calendar to replace it. When it expires, the form stops sending.
5. Recommended: limit the app to the hello@ mailbox only. By default, `Mail.Send` as an application lets the app send as anyone in the company. Use Exchange Online's application access policy (or RBAC for Applications) to restrict it to hello@fivem.group.
6. In Cloudflare, your Pages project > Settings > Variables and Secrets, add:
   - `MS_TENANT_ID`: the Directory (tenant) ID
   - `MS_CLIENT_ID`: the Application (client) ID
   - `MS_CLIENT_SECRET` (secret): the secret Value from step 4
   - `APPLY_MAILBOX`: `hello@fivem.group`
7. Redeploy, then submit a test application with a resume.

Resumes are limited to 3 MB (PDF or Word). If Microsoft rejects an attachment, the application is still emailed with a note saying the resume needs to be requested. Until step 6 is done, the form shows a message asking people to email hello@fivem.group.
