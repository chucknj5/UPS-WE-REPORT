// Hands the W.E. Report its Google Sheets sync link + token, kept in Netlify
// environment variables (UATTEND_SYNC_URL / UATTEND_SYNC_TOKEN) instead of in the public repo.
export default async () =>
  new Response(JSON.stringify({ url: process.env.UATTEND_SYNC_URL || "", token: process.env.UATTEND_SYNC_TOKEN || "", account: "UPS Time Clock" }),
    { headers: { "content-type": "application/json", "cache-control": "no-store" } });
export const config = { path: "/api/config" };
