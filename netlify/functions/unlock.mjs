// POST /api/unlock {password}  -> {ok, token, exp}  (token good for 12 hours)
// GET  /api/unlock (header x-we-token) -> {ok} if the token is still good
import { password, issue, check, valid } from "../lib/lock.mjs";
const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "content-type": "application/json", "cache-control": "no-store" } });
export default async (req) => {
  const pw = password();
  if (!pw) return json({ ok: false, notset: true, error: "No password set yet. Add WE_PASSWORD in Netlify → Site configuration → Environment variables." }, 503);
  if (req.method === "GET") return json({ ok: valid(req) });
  if (req.method !== "POST") return json({ ok: false, error: "method not allowed" }, 405);
  let b = {}; try { b = await req.json(); } catch {}
  if (!check(pw, b.password)) { await new Promise(r => setTimeout(r, 800)); return json({ ok: false, error: "Wrong password" }, 401); }
  return json({ ok: true, ...issue(pw) });
};
export const config = { path: "/api/unlock" };
