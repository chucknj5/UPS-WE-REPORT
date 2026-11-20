// Shared helpers for the W.E. Report password lock. The password lives only in the
// Netlify environment variable WE_PASSWORD (never in the repo).
import { createHmac, timingSafeEqual } from "node:crypto";
const HOURS = 12;
const sig = (pw, exp) => createHmac("sha256", "we-lock:" + pw).update(String(exp)).digest("hex");
const same = (a, b) => { const x = Buffer.from(String(a)), y = Buffer.from(String(b)); return x.length === y.length && timingSafeEqual(x, y); };
export const password = () => (globalThis.Netlify ? Netlify.env.get("WE_PASSWORD") : process.env.WE_PASSWORD) || "";
export function issue(pw) { const exp = Date.now() + HOURS * 3600e3; return { token: exp + "." + sig(pw, exp), exp }; }
export function check(pw, attempt) { return !!pw && same(pw, attempt || ""); }
export function valid(req) {
  const pw = password(); if (!pw) return false;
  const t = req.headers.get("x-we-token") || ""; const [exp, s] = t.split(".");
  return !!exp && +exp > Date.now() && same(s || "", sig(pw, exp));
}
