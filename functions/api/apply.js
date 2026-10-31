// Cloudflare Pages Function: POST /api/apply
// Receives the Employment form and emails it (resume attached) from your Microsoft 365
// mailbox using the Microsoft Graph API.
//
// Settings (Cloudflare dashboard > your Pages project > Settings > Variables and Secrets):
//   MS_TENANT_ID       text    Directory (tenant) ID of your Microsoft 365 / Entra tenant
//   MS_CLIENT_ID       text    Application (client) ID of the app registration
//   MS_CLIENT_SECRET   secret  Client secret value of that app registration
//   APPLY_MAILBOX      text    Mailbox the application is sent from, e.g. hello@fivem.group
//   APPLY_TO_EMAIL     text    (optional) where it is sent; defaults to APPLY_MAILBOX.
//                              Separate several addresses with commas.

// Graph accepts attachments up to about 3 MB in a single sendMail request.
const MAX_RESUME_BYTES = 3 * 1024 * 1024;
const ALLOWED_RESUME = /\.(pdf|docx?|rtf|txt)$/i;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const MIME = {
  pdf: "application/pdf",
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  rtf: "application/rtf",
  txt: "text/plain",
};

const POSITIONS = [
  "Back Of House (cook, prep, dish, bake, etc.)",
  "Front Of House (cashier, bartender, server, host, buss, runner, etc.)",
  "Administration (marketing, accounting, design, programming, etc)",
  "Leadership (supervisor, trainer, manager, etc.)",
];
const TIMELINES = ["ASAP", "Within 2-4 weeks", "Within 1-2 months", "Open to needs"];
const AVAILABILITY = [
  "Weekday mornings",
  "Weekday afternoons / evenings",
  "Weekend mornings",
  "Weekend afternoons / evenings",
  "Open availability",
];
const HOURS = ["Full-time", "Part-time", "Either"];
const YES_NO = ["Yes", "No"];
const EXPERIENCE = ["No experience yet", "Less than 1 year", "1-2 years", "3-5 years", "More than 5 years"];
const HEARD = [
  "Referral from a team member or friend",
  "Indeed or another job board",
  "Social media",
  "Walked in / saw a sign",
  "Other",
];

const json = (body, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });

const clean = (value, max) =>
  String(value ?? "")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .trim()
    .slice(0, max);
const oneLine = (value, max) => clean(value, max).replace(/[\r\n]+/g, " ");

function toBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

// Access tokens last about an hour, so reuse one while this instance stays warm.
const tokenCache = new Map();

async function getToken(env) {
  const key = `${env.MS_TENANT_ID}:${env.MS_CLIENT_ID}`;
  const cached = tokenCache.get(key);
  if (cached && cached.expires > Date.now() + 60_000) return cached.token;

  const res = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(env.MS_TENANT_ID)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: env.MS_CLIENT_ID,
        client_secret: env.MS_CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
    },
  );
  let data = null;
  try {
    data = await res.json();
  } catch {
    // Non-JSON response; handled below.
  }
  if (!res.ok || !data?.access_token) {
    throw new Error(`Token request failed (${res.status}): ${data?.error_description || data?.error || "no detail"}`);
  }
  tokenCache.set(key, { token: data.access_token, expires: Date.now() + (Number(data.expires_in) || 3000) * 1000 });
  return data.access_token;
}

async function sendMail(env, token, message) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(env.APPLY_MAILBOX)}/sendMail`,
    {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ message, saveToSentItems: false }),
    },
  );
  if (res.ok) return { ok: true, status: res.status };
  let detail = "";
  try {
    detail = JSON.stringify((await res.json()).error || {});
  } catch {
    // No JSON detail available.
  }
  return { ok: false, status: res.status, detail };
}

export async function onRequestPost({ request, env }) {
  const declared = Number(request.headers.get("content-length") || 0);
  if (declared > MAX_RESUME_BYTES + 512 * 1024) {
    return json({ ok: false, error: "That submission is too large. Please attach a resume under 3 MB." }, 413);
  }

  let form;
  try {
    form = await request.formData();
  } catch {
    return json({ ok: false, error: "We couldn't read that submission. Please try again." }, 400);
  }

  // Hidden field that real people never fill in; bots usually do.
  if (clean(form.get("company_website"), 200)) return json({ ok: true });

  const first = oneLine(form.get("firstName"), 80);
  const last = oneLine(form.get("lastName"), 80);
  const email = oneLine(form.get("email"), 200);
  const phone = oneLine(form.get("phone"), 40);
  const details = clean(form.get("details"), 5000);
  const positions = form.getAll("position").map(String).filter((p) => POSITIONS.includes(p));
  const timeline = form.getAll("timeline").map(String).filter((t) => TIMELINES.includes(t));
  const availability = form.getAll("availability").map(String).filter((a) => AVAILABILITY.includes(a));
  const pick = (name, list) => {
    const value = String(form.get(name) ?? "");
    return list.includes(value) ? value : "";
  };
  const hours = pick("hours", HOURS);
  const workEligible = pick("workEligible", YES_NO);
  const alcoholEligible = form.get("alcoholEligible") === "Yes";
  const experience = pick("experience", EXPERIENCE);
  const heard = pick("heard", HEARD);
  const recentJob = oneLine(form.get("recentJob"), 200);
  const refs = [1, 2]
    .map((n) => ({ name: oneLine(form.get(`ref${n}Name`), 120), contact: oneLine(form.get(`ref${n}Contact`), 120) }))
    .filter((r) => r.name || r.contact);
  const smsConsent = form.get("smsConsent") === "Yes";

  if (!first || !last) return json({ ok: false, error: "Please enter your first and last name." }, 400);
  if (!EMAIL_RE.test(email)) return json({ ok: false, error: "Please enter a valid email address." }, 400);
  if (!phone) return json({ ok: false, error: "Please enter a phone number." }, 400);
  if (!positions.length) return json({ ok: false, error: "Please choose at least one position." }, 400);
  if (!availability.length) return json({ ok: false, error: "Please choose at least one availability option." }, 400);
  if (!hours) return json({ ok: false, error: "Please choose the hours you are looking for." }, 400);
  if (!timeline.length) return json({ ok: false, error: "Please choose a timeline." }, 400);
  if (!workEligible) return json({ ok: false, error: "Please answer the work eligibility question." }, 400);
  if (!details) return json({ ok: false, error: "Please tell us a little about your experience and goals." }, 400);

  const attachments = [];
  let resumeName = "";
  const resume = form.get("resume");
  if (resume && typeof resume === "object" && resume.size > 0) {
    if (resume.size > MAX_RESUME_BYTES) {
      return json({ ok: false, error: "That resume is over 3 MB. Please attach a smaller file." }, 413);
    }
    if (!ALLOWED_RESUME.test(resume.name || "")) {
      return json({ ok: false, error: "Please attach your resume as a PDF or Word document." }, 400);
    }
    resumeName = oneLine(resume.name, 120).replace(/[\\/]/g, "_");
    const extension = resumeName.split(".").pop().toLowerCase();
    attachments.push({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: resumeName,
      contentType: MIME[extension] || "application/octet-stream",
      contentBytes: toBase64(await resume.arrayBuffer()),
    });
  }

  if (!env.MS_TENANT_ID || !env.MS_CLIENT_ID || !env.MS_CLIENT_SECRET || !env.APPLY_MAILBOX) {
    console.error("Applications are not configured: set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET and APPLY_MAILBOX.");
    return json({ ok: false, error: "Applications aren't set up yet. Please email hello@fivem.group." }, 500);
  }

  const buildText = (resumeLine) =>
    [
      "New job application from the FiveM Group website",
      "",
      `Name: ${first} ${last}`,
      `Email: ${email}`,
      `Phone: ${phone}`,
      "",
      "Position(s):",
      ...positions.map((p) => `  - ${p}`),
      "",
      "Availability:",
      ...availability.map((a) => `  - ${a}`),
      `Hours wanted: ${hours}`,
      `Start timeline: ${timeline.join(", ")}`,
      "",
      `Legally eligible to work in the U.S.: ${workEligible}`,
      `Eligible for alcohol server positions (21+ and can take a servers class): ${alcoholEligible ? "Yes" : "No"}`,
      `Experience: ${experience || "not given"}`,
      `Most recent employer and role: ${recentJob || "not given"}`,
      "",
      "References:",
      ...(refs.length ? refs.map((r) => `  - ${r.name || "(no name)"}${r.contact ? `, ${r.contact}` : ""}`) : ["  none given"]),
      "",
      `How they heard about us: ${heard || "not given"}`,
      `OK to text: ${smsConsent ? "Yes" : "No"}`,
      "",
      "Additional details:",
      details,
      "",
      `Resume: ${resumeLine}`,
    ].join("\n");

  const toRecipients = (env.APPLY_TO_EMAIL || env.APPLY_MAILBOX)
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));

  const message = {
    subject: `New application: ${first} ${last}`,
    toRecipients,
    replyTo: [{ emailAddress: { address: email, name: `${first} ${last}` } }],
  };
  const withBody = (text) => ({ contentType: "Text", content: text });

  let token;
  try {
    token = await getToken(env);
  } catch (error) {
    console.error(String(error));
    return json({ ok: false, error: "We couldn't send your application. Please try again or email hello@fivem.group." }, 502);
  }

  let result = await sendMail(
    env,
    token,
    attachments.length
      ? { ...message, body: withBody(buildText(`${resumeName} (attached)`)), attachments }
      : { ...message, body: withBody(buildText("none")) },
  );

  // If Microsoft rejects the attachment, still deliver the application itself so nobody
  // is lost, and say in the email that the resume needs to be requested.
  if (!result.ok && attachments.length) {
    console.error("Send with resume failed; retrying without it", result.status, result.detail);
    result = await sendMail(env, token, {
      ...message,
      body: withBody(buildText(`${resumeName} (could not be attached automatically; please ask the applicant to email it)`)),
    });
  }

  if (!result.ok) {
    console.error("Graph sendMail failed", result.status, result.detail);
    return json({ ok: false, error: "We couldn't send your application. Please try again or email hello@fivem.group." }, 502);
  }
  return json({ ok: true });
}

export const onRequest = () =>
  new Response("Method not allowed", { status: 405, headers: { Allow: "POST" } });
