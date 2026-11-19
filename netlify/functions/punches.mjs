// Live uAttend Punch Report for the Punch Board.
// GET /api/punches?start=YYYY-MM-DD&end=YYYY-MM-DD  (max 3 days)
// The uAttend key lives in the Netlify env var UATTEND_API_KEY (never in the repo).
const API = "https://api.workwelltech.com";
const HEAD = ["DATE", "ACCOUNT", "PUNCH ID", "USER ID", "NAME", "PAYROLL ID", "DEPT CODE", "DEPARTMENT", "SITE",
  "IN DATE", "IN TIME", "OUT DATE", "OUT TIME", "TOTAL", "IN METHOD", "OUT METHOD", "IN CLOCK", "OUT CLOCK"];
const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "content-type": "application/json", "cache-control": "no-store" } });
const day = v => { const s = String(v || ""); const m = s.match(/^(\d{4}-\d{2}-\d{2})/); return m ? m[1] : ""; };
const hhmm = t => { const m = String(t || "").match(/(\d{1,2}):(\d{2})/); return m ? m[1].padStart(2, "0") + ":" + m[2] : ""; };
const hrs = s => { s = String(s || "").trim(); if (!s) return 0; if (!s.includes(":")) return +s || 0; const neg = s[0] === "-"; const [h, m] = s.replace("-", "").split(":"); const v = (+h || 0) + (+m || 0) / 60; return neg ? -v : v; };

async function page(key, d, n) {
  for (let a = 0; a < 4; a++) {
    const r = await fetch(API + "/reports/punch", { method: "POST", headers: { "content-type": "application/json", "x-api-key": key },
      body: JSON.stringify({ StartDate: d, EndDate: d, UsePaging: true, PageSize: 500, PageNumber: n }) });
    if (r.status === 429 || r.status >= 500) { await new Promise(x => setTimeout(x, 800 * (a + 1))); continue; }
    if (!r.ok) throw new Error("uAttend " + r.status + ": " + (await r.text()).slice(0, 200));
    const j = await r.json(); return j.Body || j;
  }
  throw new Error("uAttend kept failing (rate limit)");
}

export default async (req) => {
  const key = Netlify.env.get("UATTEND_API_KEY"), acct = Netlify.env.get("UATTEND_ACCOUNT") || "UPS Time Clock";
  if (!key) return json({ ok: false, error: "UATTEND_API_KEY is not set on this Netlify site" }, 500);
  const u = new URL(req.url), today = new Date().toISOString().slice(0, 10);
  const end = u.searchParams.get("end") || today, start = u.searchParams.get("start") || end;
  const days = []; for (let d = new Date(start + "T12:00:00Z"); d <= new Date(end + "T12:00:00Z") && days.length < 3; d.setUTCDate(d.getUTCDate() + 1)) days.push(d.toISOString().slice(0, 10));
  try {
    const rows = [];
    await Promise.all(days.map(async d => {
      const first = await page(key, d, 0), items = [...(first.PunchReportLineItems || [])];
      const total = first.Total || items.length, pages = Math.ceil(total / 500);
      const rest = await Promise.all(Array.from({ length: Math.max(0, pages - 1) }, (_, i) => page(key, d, i + 1)));
      rest.forEach(b => items.push(...(b.PunchReportLineItems || [])));
      items.forEach(it => {
        const dept = String(it.DepartmentName || it.DepartmentCode || "").trim(), sm = dept.match(/^(\d{2})/);
        rows.push([day(it.PunchDate || it.InDate) || d, acct, String(it.Id ?? ""), String(it.UserId ?? ""),
          (String(it.FirstName || "").trim() + " " + String(it.LastName || "").trim()).replace(/\s+/g, " "), String(it.PayrollNumber || ""),
          String(it.DepartmentCode || ""), dept, sm ? sm[1] : "", day(it.InDate), hhmm(it.InTime), it.OutTime ? day(it.OutDate) : "", hhmm(it.OutTime),
          (Math.round(hrs(it.Tot) * 100) / 100).toFixed(2), it.InPunchMethod || "", it.OutPunchMethod || "", it.InPunchLocation || "", it.OutPunchLocation || ""]);
      });
    }));
    return json({ ok: true, data: { head: HEAD, rows, pulled: { t: Date.now(), source: "uAttend live" } } });
  } catch (e) { return json({ ok: false, error: String(e.message || e) }, 502); }
};
export const config = { path: "/api/punches" };
