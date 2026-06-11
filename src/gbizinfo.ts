/**
 * gBizINFO（経済産業省 法人データ提供サービス）クローラ + アタックリスト生成ロジック。
 *
 * 目的:
 *   aicrew の「生成AI / DX 導入伴走支援」の営業アタック先として、
 *   東京都内・従業員100人以下・年商10〜30億円規模・レガシー（FAX/紙/SaaS乱立）な
 *   中小企業を gBizINFO REST API から抽出し、スコアリングして CSV/xlsx に出力する。
 *
 * データソース:
 *   gBizINFO REST API v1  https://info.gbiz.go.jp/hojin/v1/hojin
 *   認証ヘッダ: X-hojinInfo-api-token: <無料トークン>
 *   ※ トークンは https://info.gbiz.go.jp/ のマイページから無料取得（無期限）。
 *
 * 重要な前提（誠実性のため明記）:
 *   - 「年商（売上高）」と「FAX/SaaS利用などレガシー度」は法人登記オープンデータには
 *     基本的に含まれない。本モジュールは
 *       (1) gBizINFO に財務（売上高）があればそれを使う（revenue_confidence = "reported"）
 *       (2) 無ければ 従業員数 × 業種別の従業員あたり売上 で推定する（"estimated"）
 *       (3) レガシー度は 業種・設立年・自社サイト有無 等のプロキシからの推定スコア
 *     として扱う。最終的な確度はメール送信前に人手で検証する前提。
 */

import { parseJapaneseAddress, normalizePhone } from "./parse-address";
import * as XLSX from "xlsx";

// ───────────────────────────────────────────────────────────────────────────
// 1. API クライアント
// ───────────────────────────────────────────────────────────────────────────

const GBIZ_BASE = "https://info.gbiz.go.jp/hojin/v1/hojin";

/** 都道府県名 → JIS X 0401 コード（今回は東京都=13 が主。拡張用に主要都市も用意） */
export const PREFECTURE_CODE: Record<string, string> = {
  東京都: "13",
  神奈川県: "14",
  埼玉県: "11",
  千葉県: "12",
  大阪府: "27",
  愛知県: "23",
};

/** gBizINFO の検索クエリ（v1）。未指定キーは送らない。 */
export type GbizSearchParams = {
  prefecture?: string; // JISコード "13"
  city?: string;
  name?: string;
  business_item?: string; // 日本標準産業分類コード（カンマ区切り可）
  capital_stock_from?: number;
  capital_stock_to?: number;
  employee_number_from?: number;
  employee_number_to?: number;
  founded_year?: number;
  exist_flg?: boolean; // 廃業除外（true=現存）
  page?: number;
};

/** gBizINFO v1 の hojin-infos 要素（必要なフィールドのみ・防御的に optional） */
export type GbizHojin = {
  corporate_number?: string;
  name?: string;
  location?: string;
  postal_code?: string;
  company_url?: string;
  date_of_establishment?: string; // "YYYY-MM-DD"
  founding_year?: number;
  capital_stock?: number; // 円
  employee_number?: number;
  business_summary?: string;
  business_items?: Array<string | { key?: string; value?: string }>;
  representative_name?: string;
  representative_position?: string;
  close_cause?: string | null;
  // 財務（存在する法人のみ）
  finance?: unknown;
};

export class GbizClient {
  constructor(
    private readonly token: string,
    private readonly opts: { delayMs?: number; fetchImpl?: typeof fetch } = {},
  ) {
    if (!token) throw new Error("gBizINFO API トークンが空です（GBIZINFO_API_TOKEN）");
  }

  private get f(): typeof fetch {
    return this.opts.fetchImpl ?? fetch;
  }

  private async sleep(ms: number) {
    await new Promise((r) => setTimeout(r, ms));
  }

  /** 条件検索（1ページ分）。失敗時は例外。 */
  async searchPage(params: GbizSearchParams): Promise<{ hits: GbizHojin[]; raw: unknown }> {
    const qs = new URLSearchParams();
    for (const [k, v] of Object.entries(params)) {
      if (v === undefined || v === null || v === "") continue;
      qs.set(k, String(v));
    }
    const url = `${GBIZ_BASE}?${qs.toString()}`;
    const res = await this.f(url, {
      headers: { "X-hojinInfo-api-token": this.token, Accept: "application/json" },
    });
    if (res.status === 403 || res.status === 401) {
      throw new Error(`認証エラー(${res.status})。トークンを確認してください。`);
    }
    if (res.status === 429) {
      throw new Error("レート制限(429)。--delay を増やして再実行してください。");
    }
    if (!res.ok) {
      throw new Error(`gBizINFO HTTP ${res.status}: ${url}`);
    }
    const json = (await res.json()) as { "hojin-infos"?: GbizHojin[] };
    return { hits: json["hojin-infos"] ?? [], raw: json };
  }

  /** ページを回して全件取得（maxPages 上限・delay 付き）。 */
  async searchAll(
    params: GbizSearchParams,
    opts: { maxPages?: number; onPage?: (page: number, count: number) => void } = {},
  ): Promise<GbizHojin[]> {
    const maxPages = opts.maxPages ?? 10;
    const delay = this.opts.delayMs ?? 300;
    const out: GbizHojin[] = [];
    for (let page = 1; page <= maxPages; page++) {
      const { hits } = await this.searchPage({ ...params, page });
      opts.onPage?.(page, hits.length);
      if (hits.length === 0) break;
      out.push(...hits);
      await this.sleep(delay);
    }
    return out;
  }
}

// ───────────────────────────────────────────────────────────────────────────
// 2. 業種分類（クライアント側キーワード分類）
//    business_item パラメータの挙動に依存しすぎないよう、取得後に自社バケットへ分類する。
// ───────────────────────────────────────────────────────────────────────────

export type TargetIndustry = "製造" | "卸売・商社" | "建設・設備" | "物流・運輸" | "不動産管理";

const INDUSTRY_KEYWORDS: Record<TargetIndustry, string[]> = {
  製造: ["製造", "製作所", "工業", "金属", "加工", "鋳造", "プレス", "部品", "機械", "電子", "鉄工", "樹脂", "板金", "工場"],
  "卸売・商社": ["卸", "卸売", "商事", "商会", "商社", "問屋", "貿易", "仕入", "販売", "物産"],
  "建設・設備": ["建設", "工務店", "設備", "電気工事", "空調", "管工事", "塗装", "土木", "建築", "施工", "リフォーム"],
  "物流・運輸": ["運輸", "運送", "物流", "倉庫", "配送", "輸送", "ロジ", "トラック", "陸運", "梱包"],
  不動産管理: ["不動産", "賃貸管理", "ビル管理", "管理組合", "プロパティ", "マンション管理", "仲介"],
};

/** name + business_summary + business_items から対象業種を1つ推定。該当なしは null。 */
export function classifyIndustry(h: GbizHojin): TargetIndustry | null {
  const items = (h.business_items ?? [])
    .map((x) => (typeof x === "string" ? x : x?.value ?? x?.key ?? ""))
    .join(" ");
  const text = `${h.name ?? ""} ${h.business_summary ?? ""} ${items}`;
  let best: TargetIndustry | null = null;
  let bestScore = 0;
  for (const [ind, kws] of Object.entries(INDUSTRY_KEYWORDS) as [TargetIndustry, string[]][]) {
    const score = kws.reduce((s, kw) => (text.includes(kw) ? s + 1 : s), 0);
    if (score > bestScore) {
      bestScore = score;
      best = ind;
    }
  }
  return bestScore > 0 ? best : null;
}

// ───────────────────────────────────────────────────────────────────────────
// 3. 年商推定 + レガシー/自動化スコア + ティア付け
// ───────────────────────────────────────────────────────────────────────────

/** 業種別 従業員1人あたり年商（円）のベンチマーク（粗い推定値）。 */
const REVENUE_PER_EMPLOYEE: Record<TargetIndustry, number> = {
  製造: 30_000_000,
  "卸売・商社": 120_000_000, // 卸は回転が大きく一人あたり売上が高い
  "建設・設備": 40_000_000,
  "物流・運輸": 18_000_000,
  不動産管理: 50_000_000,
};

const OKU = 100_000_000; // 1億円
export const TARGET_REVENUE_MIN = 10 * OKU; // 10億
export const TARGET_REVENUE_MAX = 30 * OKU; // 30億

/** 自動化の切り口（aicrew 提案フック）。 */
const AUTOMATION_OPPORTUNITY: Record<TargetIndustry, string> = {
  製造: "FAX/電話受発注・生産管理Excel・在庫照合の自動化（OCR+RPA、生成AIでの図面/仕様問合せ対応）",
  "卸売・商社": "FAX受注・見積/請求の手入力をOCR+生成AIで自動起票、問合せ一次対応の自動化",
  "建設・設備": "見積/工程/原価管理の帳票自動生成、現場写真・日報の生成AI整理",
  "物流・運輸": "配車/伝票/請求処理の自動化、問合せ・追跡対応のAIチャット化",
  不動産管理: "契約更新・入居者問合せ対応の自動化、紙台帳のデータ化",
};

export type RevenueConfidence = "reported" | "estimated";
export type Tier = "S" | "A" | "B" | "C";

export type Prospect = {
  company_name: string;
  corporate_number: string | null;
  industry: TargetIndustry;
  prefecture: string | null;
  city: string | null;
  address: string | null;
  postal_code: string | null;
  representative: string | null;
  website_url: string | null;
  has_website: boolean;
  founded_year: number | null;
  company_age: number | null;
  capital_stock: number | null;
  employee_number: number | null;
  business_summary: string | null;
  est_annual_revenue_oku: number | null; // 億円
  revenue_confidence: RevenueConfidence;
  revenue_in_target_band: boolean;
  legacy_score: number; // 0-100
  automation_opportunity: string;
  priority_tier: Tier;
  tier_reasoning: string;
  recommended_hook: string;
  contact_method: string;
  source: string;
  crawled_at: string;
};

/** gBizINFO の finance 構造から売上高(円)を緩く取り出す。無ければ null。 */
export function extractReportedRevenue(h: GbizHojin): number | null {
  const fin = h.finance as unknown;
  if (!fin) return null;
  const arr = Array.isArray(fin) ? fin : [fin];
  let latest: number | null = null;
  for (const f of arr) {
    if (!f || typeof f !== "object") continue;
    const rec = f as Record<string, unknown>;
    // 代表的なキー名のゆれを吸収
    const v =
      rec.net_sales_summary_of_business_results ??
      rec.net_sales ??
      rec.netSales ??
      rec.sales;
    const n = typeof v === "number" ? v : Number(v);
    if (Number.isFinite(n) && n > 0) latest = n;
  }
  return latest;
}

const CURRENT_YEAR = new Date().getFullYear();

function foundedYearOf(h: GbizHojin): number | null {
  if (h.founding_year && Number.isFinite(h.founding_year)) return h.founding_year;
  if (h.date_of_establishment) {
    const y = Number(String(h.date_of_establishment).slice(0, 4));
    if (Number.isFinite(y) && y > 1800) return y;
  }
  return null;
}

/**
 * レガシー度スコア（0-100）。高いほど「FAX/紙/SaaS乱立」で自動化余地が大きいと推定。
 * プロキシ:
 *  - 設立が古い（社歴長い）           +最大35
 *  - 自社サイトが無い/簡素            +25（デジタル化が遅れている兆候）
 *  - 対象レガシー業種である           +25（業種ベース）
 *  - 規模が中堅（20〜100人）で属人化   +最大15
 */
export function legacyScore(h: GbizHojin, industry: TargetIndustry): number {
  let s = 0;
  const fy = foundedYearOf(h);
  if (fy) {
    const age = CURRENT_YEAR - fy;
    s += Math.min(35, Math.max(0, Math.round((age / 50) * 35))); // 50年で満点
  } else {
    s += 10;
  }
  if (!h.company_url) s += 25;
  // 業種ベース：製造/卸/建設/物流はFAX文化が根強い
  const industryBase: Record<TargetIndustry, number> = {
    製造: 25,
    "卸売・商社": 25,
    "建設・設備": 22,
    "物流・運輸": 20,
    不動産管理: 15,
  };
  s += industryBase[industry];
  const emp = h.employee_number ?? 0;
  if (emp >= 20 && emp <= 100) s += 15;
  else if (emp > 0) s += 8;
  return Math.min(100, s);
}

/** 1社を Prospect に変換。対象外（業種ヒットなし）は null。 */
export function toProspect(h: GbizHojin, opts: { source?: string } = {}): Prospect | null {
  const industry = classifyIndustry(h);
  if (!industry) return null;

  const addr = parseJapaneseAddress(h.location);
  const fy = foundedYearOf(h);
  const emp = h.employee_number ?? null;

  // 年商: 財務があれば採用、無ければ従業員数×業種ベンチで推定
  const reported = extractReportedRevenue(h);
  let revenueYen: number | null = reported;
  let confidence: RevenueConfidence = "reported";
  if (revenueYen === null) {
    confidence = "estimated";
    revenueYen = emp ? emp * REVENUE_PER_EMPLOYEE[industry] : null;
  }
  const revOku = revenueYen !== null ? Math.round((revenueYen / OKU) * 10) / 10 : null;
  const inBand =
    revenueYen !== null && revenueYen >= TARGET_REVENUE_MIN && revenueYen <= TARGET_REVENUE_MAX;

  const legacy = legacyScore(h, industry);
  const empOk = emp !== null && emp <= 100;

  // ティア: 年商バンド一致 × 従業員100人以下 × レガシー度 で決定
  let tier: Tier = "C";
  const reasons: string[] = [];
  if (inBand && empOk && legacy >= 70) {
    tier = "S";
    reasons.push("年商10-30億帯×100人以下×高レガシー");
  } else if (inBand && empOk && legacy >= 55) {
    tier = "A";
    reasons.push("年商帯一致×規模適合×レガシー中〜高");
  } else if (empOk && (inBand || legacy >= 60)) {
    tier = "B";
    reasons.push(inBand ? "規模適合・年商帯一致(レガシー中)" : "規模適合・高レガシー(年商要確認)");
  } else {
    tier = "C";
    reasons.push(!empOk ? "従業員100人超または不明" : "年商帯/レガシーとも条件外");
  }
  if (confidence === "estimated") reasons.push("年商は推定値（要検証）");
  if (!h.company_url) reasons.push("自社サイト未検出（デジタル化遅れの兆候）");

  return {
    company_name: h.name ?? "",
    corporate_number: h.corporate_number ?? null,
    industry,
    prefecture: addr.prefecture,
    city: addr.city,
    address: addr.address_rest || h.location || null,
    postal_code: h.postal_code ?? null,
    representative: h.representative_name ?? null,
    website_url: h.company_url ?? null,
    has_website: Boolean(h.company_url),
    founded_year: fy,
    company_age: fy ? CURRENT_YEAR - fy : null,
    capital_stock: h.capital_stock ?? null,
    employee_number: emp,
    business_summary: h.business_summary ?? null,
    est_annual_revenue_oku: revOku,
    revenue_confidence: confidence,
    revenue_in_target_band: inBand,
    legacy_score: legacy,
    automation_opportunity: AUTOMATION_OPPORTUNITY[industry],
    priority_tier: tier,
    tier_reasoning: reasons.join(" / "),
    recommended_hook: buildHook(industry),
    contact_method: h.company_url ? "問い合わせフォーム/電話" : "電話/郵送",
    source: opts.source ?? "gbizinfo",
    crawled_at: new Date().toISOString(),
  };
}

/** aicrew 資料のメトリクスに紐づく提案フック（営業メールの切り口）。 */
function buildHook(industry: TargetIndustry): string {
  const base = AUTOMATION_OPPORTUNITY[industry];
  return `${base}。PoC止まりにせず現場で動く成果へ（処理時間-88%/3年最大3,675万円コスト削減/AI導入コスト1/10/伴走型）`;
}

// ───────────────────────────────────────────────────────────────────────────
// 4. 重複排除 + 並べ替え
// ───────────────────────────────────────────────────────────────────────────

export function dedupeAndRank(prospects: Prospect[]): Prospect[] {
  const seen = new Set<string>();
  const uniq: Prospect[] = [];
  for (const p of prospects) {
    const key = p.corporate_number ?? `${p.company_name}|${p.address ?? ""}`;
    if (seen.has(key)) continue;
    seen.add(key);
    uniq.push(p);
  }
  const tierRank: Record<Tier, number> = { S: 0, A: 1, B: 2, C: 3 };
  uniq.sort(
    (a, b) =>
      tierRank[a.priority_tier] - tierRank[b.priority_tier] ||
      b.legacy_score - a.legacy_score,
  );
  return uniq;
}

// ───────────────────────────────────────────────────────────────────────────
// 5. 出力（CSV / xlsx）
// ───────────────────────────────────────────────────────────────────────────

/** 出力カラム（順序固定）。日本語ヘッダ。 */
export const OUTPUT_COLUMNS: Array<{ key: keyof Prospect; header: string }> = [
  { key: "priority_tier", header: "優先度" },
  { key: "company_name", header: "法人名" },
  { key: "industry", header: "業種" },
  { key: "prefecture", header: "都道府県" },
  { key: "city", header: "市区町村" },
  { key: "address", header: "住所" },
  { key: "postal_code", header: "郵便番号" },
  { key: "representative", header: "代表者" },
  { key: "website_url", header: "URL" },
  { key: "has_website", header: "自社サイト有無" },
  { key: "founded_year", header: "設立年" },
  { key: "company_age", header: "社歴(年)" },
  { key: "capital_stock", header: "資本金(円)" },
  { key: "employee_number", header: "従業員数" },
  { key: "est_annual_revenue_oku", header: "推定年商(億円)" },
  { key: "revenue_confidence", header: "年商確度" },
  { key: "revenue_in_target_band", header: "10-30億帯" },
  { key: "legacy_score", header: "レガシー度" },
  { key: "automation_opportunity", header: "自動化余地" },
  { key: "recommended_hook", header: "提案フック" },
  { key: "tier_reasoning", header: "優先度の根拠" },
  { key: "contact_method", header: "想定接触手段" },
  { key: "business_summary", header: "事業概要" },
  { key: "corporate_number", header: "法人番号" },
  { key: "source", header: "ソース" },
  { key: "crawled_at", header: "取得日時" },
];

function cell(v: unknown): string | number | boolean {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "○" : "";
  return v as string | number;
}

export function toAoa(prospects: Prospect[]): (string | number | boolean)[][] {
  const header = OUTPUT_COLUMNS.map((c) => c.header);
  const rows = prospects.map((p) => OUTPUT_COLUMNS.map((c) => cell(p[c.key])));
  return [header, ...rows];
}

export function toCsv(prospects: Prospect[]): string {
  const aoa = toAoa(prospects);
  const esc = (v: string | number | boolean) => {
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  // Excel(日本語)で文字化けしないよう UTF-8 BOM を付与
  return "﻿" + aoa.map((r) => r.map(esc).join(",")).join("\r\n");
}

export function toXlsxBuffer(prospects: Prospect[]): Buffer {
  const ws = XLSX.utils.aoa_to_sheet(toAoa(prospects));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "prospects");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

// normalizePhone は将来の電話列正規化用に再エクスポート（現状 gBizINFO は電話を返さない）
export { normalizePhone };
