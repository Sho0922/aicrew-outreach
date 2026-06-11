#!/usr/bin/env tsx
/**
 * アタックリスト（data/prospects.xlsx）の各社サイト(URL列)を巡回し、
 * 問い合わせメールアドレス／問い合わせフォームURLを「取れるだけ」抽出して列を追加する。
 *
 * 使い方:
 *   pnpm extract-emails                       # data/prospects.xlsx を更新（メール/問い合わせURL列を追加）
 *   pnpm extract-emails --tier=S,A            # 指定ティアのみ巡回（速い）
 *   pnpm extract-emails --in=data/priority_SA.xlsx --out=data/priority_SA.xlsx
 *   pnpm extract-emails --concurrency=6 --timeout=8000 --limit=200
 *
 * 注意:
 *   - gBizINFO にメールは無いため、各社の公開サイトから抽出する。フォームのみ/サイト無しは空欄。
 *   - 外部サイトへアクセスするため、礼儀として同時実行数・タイムアウトを控えめに。
 *   - 取得したアドレスへの営業送信は特定電子メール法を遵守（差出人明示・配信停止導線）。
 */
import { config } from "dotenv";
import { writeFileSync, readFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";
import * as XLSX from "xlsx";

config({ path: resolve(process.cwd(), ".env.local") });

function parseArgs(argv: string[]) {
  const args: Record<string, string | boolean> = {};
  for (const a of argv.slice(2)) {
    const m = a.match(/^--([^=]+)(?:=(.*))?$/);
    if (!m) continue;
    args[m[1]] = m[2] ?? true;
  }
  return args;
}

type Row = Record<string, string | number | undefined>;

const UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36";

const EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g;
const ROLE_PREFIX = ["info", "contact", "inquiry", "otoiawase", "mail", "sales", "office", "support", "soumu", "general"];
const BAD_DOMAINS = [
  "example.com", "example.co.jp", "sentry.io", "wixpress.com", "w3.org", "schema.org",
  "googleapis.com", "gstatic.com", "cloudflare", "your-domain", "domain.com", "email.com",
  "test.com", "sample.com", "godaddy.com", "wordpress.com", "jimdo.com", "27.media",
];
const BAD_LOCAL_EXT = /\.(png|jpe?g|gif|svg|webp|css|js)$/i;

function decodeEntities(s: string): string {
  return s
    .replace(/&#0*64;/g, "@")
    .replace(/&#x0*40;/gi, "@")
    .replace(/&#0*46;/g, ".")
    .replace(/&#x0*2e;/gi, ".")
    .replace(/\s*\[at\]\s*|\s*\(at\)\s*|\s*＠\s*/gi, "@")
    .replace(/\s*\[dot\]\s*|\s*\(dot\)\s*/gi, ".");
}

function hostOf(url: string): string {
  try {
    return new URL(url).hostname.replace(/^www\./, "");
  } catch {
    return "";
  }
}

function isPlausible(email: string): boolean {
  const lower = email.toLowerCase();
  if (BAD_LOCAL_EXT.test(lower)) return false;
  if (lower.includes("@2x") || lower.includes("@example")) return false;
  const domain = lower.split("@")[1] ?? "";
  if (BAD_DOMAINS.some((b) => domain.includes(b))) return false;
  if (domain.length < 4 || !domain.includes(".")) return false;
  return true;
}

/** HTML から mailto/本文/エンティティ表記のメールを収集。 */
export function emailsFromHtml(html: string): string[] {
  const decoded = decodeEntities(html);
  const out = new Set<string>();
  // mailto: は最も確実
  const mailto = decoded.match(/mailto:([^"'?>\s]+)/gi) ?? [];
  for (const m of mailto) {
    const e = m.replace(/mailto:/i, "").trim();
    if (EMAIL_RE.test(e)) out.add(e);
    EMAIL_RE.lastIndex = 0;
  }
  // 本文中
  for (const e of decoded.match(EMAIL_RE) ?? []) out.add(e);
  return [...out].filter(isPlausible);
}

/** ベストなメールを選ぶ。サイトと同一ドメイン優先・役割アドレス優先。 */
export function pickBest(emails: string[], siteHost: string): string | null {
  if (emails.length === 0) return null;
  const score = (e: string): number => {
    const [local, domain] = e.toLowerCase().split("@");
    let s = 0;
    if (siteHost && (domain === siteHost || domain.endsWith("." + siteHost) || siteHost.endsWith("." + domain))) s += 100;
    if (ROLE_PREFIX.some((p) => local === p || local.startsWith(p))) s += 30;
    if (/(no-?reply|noreply|webmaster|postmaster|abuse)/.test(local)) s -= 50;
    s -= Math.min(20, local.length); // 短い汎用名をやや優先
    return s;
  };
  return [...emails].sort((a, b) => score(b) - score(a))[0];
}

/** 日本の電話番号を「取れるだけ」抽出。FAX番号は除外、TEL/tel: 表記を優先。 */
export function phoneFromHtml(html: string): string {
  // 全角→半角、各種ダッシュ・括弧を統一
  const s = html
    .replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0))
    .replace(/[（]/g, "(")
    .replace(/[）]/g, ")")
    .replace(/[‐‑‒–—―ーｰ−]/g, "-");

  const cands: { num: string; score: number }[] = [];
  const consider = (raw: string, ctx: string) => {
    const clean = raw
      .replace(/^\+81[-\s]?/, "0")
      .replace(/[()]/g, "-")
      .replace(/\s/g, "")
      .replace(/-+/g, "-")
      .replace(/^-|-$/g, "");
    const digits = clean.replace(/\D/g, "");
    if (!/^0\d{9,10}$/.test(digits)) return; // 国内 0始まり 10〜11桁
    if (/fax|ファ?ック|ＦＡＸ/i.test(ctx)) return; // FAXは除外
    let score = 0;
    if (/tel:|TEL|Tel|電話|代表|お問い合わせ|お問合せ/i.test(ctx)) score += 10;
    if (digits.startsWith("0120") || digits.startsWith("0800")) score += 3; // フリーダイヤル
    cands.push({ num: clean, score });
  };

  // tel: リンク（区切り無しもあり得る）
  for (const m of s.matchAll(/tel:\s*([+\d\-()]{9,18})/gi)) consider(m[1], "tel:");
  // 本文中の番号。境界アンカーで先頭0から正しく取る（貪欲truncation回避）
  const re = /(?<![\d-])(?:\+81[-\s]?|0)\d{1,4}[-(]\d{1,4}[-)]?\d{3,4}(?![\d])/g;
  let mm: RegExpExecArray | null;
  while ((mm = re.exec(s)) !== null) {
    consider(mm[0], s.slice(Math.max(0, mm.index - 12), mm.index));
  }

  if (cands.length === 0) return "";
  // 重複を集約し、スコア高い順
  const byNum = new Map<string, number>();
  for (const c of cands) byNum.set(c.num, Math.max(byNum.get(c.num) ?? -99, c.score));
  return [...byNum.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

async function fetchText(url: string, timeoutMs: number): Promise<string | null> {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      headers: { "User-Agent": UA, Accept: "text/html,*/*" },
      redirect: "follow",
      signal: ctrl.signal,
    });
    if (!res.ok) return null;
    const ct = res.headers.get("content-type") ?? "";
    if (!ct.includes("html") && !ct.includes("text")) return null;
    return await res.text();
  } catch {
    return null;
  } finally {
    clearTimeout(t);
  }
}

/** 同一ドメインの問い合わせ系リンクを抽出。 */
export function contactLinks(html: string, base: string): string[] {
  const out = new Set<string>();
  const hrefs = html.match(/href\s*=\s*["']([^"']+)["']/gi) ?? [];
  for (const h of hrefs) {
    const m = h.match(/href\s*=\s*["']([^"']+)["']/i);
    if (!m) continue;
    const href = m[1];
    if (/(contact|inquiry|toiawase|otoiawase|お問い合わせ|問い合わせ|問合せ|company|about)/i.test(href)) {
      try {
        out.add(new URL(href, base).toString());
      } catch {
        /* skip */
      }
    }
  }
  return [...out].slice(0, 3);
}

type Found = { email: string; phone: string; contactUrl: string };

async function findContact(rawUrl: string, timeoutMs: number): Promise<Found> {
  let url = rawUrl.trim();
  if (!/^https?:\/\//i.test(url)) url = "https://" + url;
  const siteHost = hostOf(url);
  const found: Found = { email: "", phone: "", contactUrl: "" };

  const home = await fetchText(url, timeoutMs);
  if (!home) return found;

  let emails = emailsFromHtml(home);
  found.phone = phoneFromHtml(home);
  const links = contactLinks(home, url);
  // 問い合わせページを優先的に巡回（メール/電話が未取得なら深掘り）
  for (const link of links) {
    if (/(contact|inquiry|toiawase|otoiawase|問い合わせ|問合せ)/i.test(link) && !found.contactUrl) {
      found.contactUrl = link;
    }
    if (emails.length === 0 || !found.phone) {
      const page = await fetchText(link, timeoutMs);
      if (page) {
        if (emails.length === 0) emails = emailsFromHtml(page);
        if (!found.phone) found.phone = phoneFromHtml(page);
      }
    }
  }
  const best = pickBest(emails, siteHost);
  if (best) found.email = best;
  return found;
}

/** 簡易並列プール。 */
async function pool<T, R>(items: T[], n: number, worker: (it: T, i: number) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let idx = 0;
  async function run() {
    while (idx < items.length) {
      const i = idx++;
      results[i] = await worker(items[i], i);
    }
  }
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, run));
  return results;
}

function csvEscape(v: unknown): string {
  const s = v === null || v === undefined ? "" : String(v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

async function main() {
  const args = parseArgs(process.argv);
  const inPath = (args.in as string) ?? "data/prospects.xlsx";
  const outBase = (args.out as string)?.replace(/\.(xlsx|csv)$/i, "") ?? inPath.replace(/\.(xlsx|csv)$/i, "");
  const concurrency = Number(args.concurrency ?? 6);
  const timeoutMs = Number(args.timeout ?? 8000);
  const limit = args.limit !== undefined ? Number(args.limit) : undefined;
  const tierFilter = (args.tier as string | undefined)
    ?.split(",").map((t) => t.trim().toUpperCase()).filter(Boolean);

  const buf = readFileSync(resolve(process.cwd(), inPath));
  const wb = XLSX.read(buf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });
  const header = (aoa[0] as string[]) ?? [];
  const rows = XLSX.utils.sheet_to_json<Row>(ws, { defval: "" });

  // 巡回対象（URLあり）を選別
  let targets = rows.map((r, i) => ({ r, i }));
  if (tierFilter) targets = targets.filter(({ r }) => tierFilter.includes(String(r["優先度"]).toUpperCase()));
  if (limit !== undefined) targets = targets.slice(0, limit);
  const withUrl = targets.filter(({ r }) => String(r["URL"] ?? "").trim() !== "");

  console.log(`🌐 巡回対象: ${withUrl.length} 社（URLあり / 全${rows.length}社中・対象${targets.length}社）`);
  let done = 0;
  let hit = 0;
  let phoneHit = 0;
  await pool(withUrl, concurrency, async ({ r }) => {
    const res = await findContact(String(r["URL"]), timeoutMs);
    if (res.email) {
      r["メールアドレス"] = res.email;
      hit++;
    }
    if (res.phone) {
      r["電話番号"] = res.phone;
      phoneHit++;
    }
    if (res.contactUrl) r["問い合わせURL"] = res.contactUrl;
    done++;
    if (done % 25 === 0 || done === withUrl.length)
      console.log(`  巡回 ${done}/${withUrl.length}（メール ${hit} / 電話 ${phoneHit}）`);
    return res;
  });

  // 出力カラム: 既存ヘッダの「URL」直後に メール/電話/問い合わせURL を挿入
  const newCols = ["メールアドレス", "電話番号", "問い合わせURL"];
  const outHeader = [...header];
  const urlIdx = outHeader.indexOf("URL");
  const insertAt = urlIdx >= 0 ? urlIdx + 1 : outHeader.length;
  outHeader.splice(insertAt, 0, ...newCols.filter((c) => !outHeader.includes(c)));

  const outAoa = [outHeader, ...rows.map((r) => outHeader.map((h) => r[h] ?? ""))];
  const csv = "﻿" + outAoa.map((row) => row.map(csvEscape).join(",")).join("\r\n");
  const wsOut = XLSX.utils.aoa_to_sheet(outAoa);
  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, wsOut, "prospects");

  mkdirSync(dirname(resolve(process.cwd(), outBase)), { recursive: true });
  writeFileSync(resolve(process.cwd(), `${outBase}.csv`), csv, "utf8");
  writeFileSync(resolve(process.cwd(), `${outBase}.xlsx`), XLSX.write(wbOut, { type: "buffer", bookType: "xlsx" }));

  console.log(`✅ メール取得 ${hit} 社 / 電話取得 ${phoneHit} 社 / 巡回 ${withUrl.length} 社`);
  console.log(`📦 更新: ${resolve(process.cwd(), outBase)}.csv / .xlsx`);
  console.log("⚠ フォームのみ/サイト無しの会社はメール空欄です。送信は特定電子メール法を遵守してください。");
}

main().catch((e) => {
  console.error(String((e as Error)?.stack ?? e));
  process.exit(1);
});
