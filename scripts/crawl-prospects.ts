#!/usr/bin/env tsx
/**
 * 東京都内 レガシー中小企業（製造/卸売・商社/建設・設備/物流・運輸/不動産管理）の
 * aicrew AI・DX 自動化 営業アタックリストを gBizINFO から生成する CLI。
 *
 * 取得戦略:
 *   gBizINFO はサイズ/財務データの欠損が多く、サイズ条件で server 側を絞ると母集団が痩せる。
 *   そこで「レガシー業種が社名に表れるキーワード」で name 検索を横断的に回し（東京都）、
 *   取得後に業種分類・年商推定・スコアリングする。サイズ条件は値がある時だけ適用。
 *
 * 前提:
 *   .env.local に GBIZINFO_API_TOKEN（無料・無期限）。
 *   取得: https://content.info.gbiz.go.jp/api/ → API利用申請 → メールのURLでトークン表示
 *
 * 使い方:
 *   pnpm crawl-prospects                          # 既定キーワードで横断取得 → data/prospects.csv/.xlsx
 *   pnpm crawl-prospects --keywords=製作所,運送,建設   # キーワードを指定
 *   pnpm crawl-prospects --emp-to=100             # 従業員データがある会社を100人以下に制限
 *   pnpm crawl-prospects --max-pages=5 --delay=400
 *   pnpm crawl-prospects --tier=S,A,B             # 出力ティアを絞る
 *   pnpm crawl-prospects:dry                       # 設定だけ表示（API未送信）
 */
import { config } from "dotenv";
import { writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";

config({ path: resolve(process.cwd(), ".env.local") });

import {
  GbizClient,
  PREFECTURE_CODE,
  classifyIndustry,
  toProspect,
  dedupeAndRank,
  toCsv,
  toXlsxBuffer,
  type GbizSearchParams,
  type GbizHojin,
  type Prospect,
  type Tier,
} from "../src/gbizinfo";

/** レガシー業種が社名に表れやすいキーワード（既定）。name 部分一致で横断取得する。 */
const DEFAULT_KEYWORDS = [
  // 製造
  "製作所", "製造", "工業", "工作", "鉄工", "金属", "板金", "精密", "樹脂", "鋳造", "部品",
  "製鋼", "製缶", "メッキ", "鍍金", "溶接", "機工", "ゴム", "硝子", "ガラス", "印刷", "紙器",
  "化成", "食品", "醸造", "木工", "家具", "建具", "繊維", "金型", "電機",
  // 卸売・商社
  "商事", "商会", "商店", "商工", "産業", "物産", "卸", "興業", "鋼材", "資材", "木材", "燃料",
  // 建設・設備
  "建設", "工務店", "設備", "電気工事", "管工事", "空調", "土木", "建材", "塗装", "防水",
  "内装", "鉄筋", "解体", "造園", "サッシ", "水道", "建工", "工営", "興産",
  // 物流・運輸
  "運送", "運輸", "物流", "倉庫", "陸運", "配送", "急便", "通運", "海運", "港運", "流通",
  // 不動産管理
  "不動産", "ビル管理", "管財", "都市開発",
];

function parseArgs(argv: string[]) {
  const args: Record<string, string | boolean> = {};
  for (const a of argv.slice(2)) {
    const m = a.match(/^--([^=]+)(?:=(.*))?$/);
    if (!m) continue;
    args[m[1]] = m[2] ?? true;
  }
  return args;
}

const BAR = "─".repeat(64);
const c = (s: string, code: number) => `\x1b[${code}m${s}\x1b[0m`;

async function main() {
  const args = parseArgs(process.argv);
  const dryRun = Boolean(args["dry-run"]);

  const prefName = (args.pref as string) ?? "東京都";
  const prefecture = PREFECTURE_CODE[prefName] ?? "13";
  const keywords = ((args.keywords as string)?.split(",").map((s) => s.trim()).filter(Boolean)) ?? DEFAULT_KEYWORDS;
  const empTo = args["emp-to"] !== undefined ? Number(args["emp-to"]) : undefined;
  const empFrom = args["emp-from"] !== undefined ? Number(args["emp-from"]) : undefined;
  const capFrom = args["cap-from"] !== undefined ? Number(args["cap-from"]) : undefined;
  const capTo = args["cap-to"] !== undefined ? Number(args["cap-to"]) : undefined;
  const maxPages = Number(args["max-pages"] ?? 5);
  const delay = Number(args["delay"] ?? 350);
  const noEnrich = Boolean(args["no-enrich"]);
  const enrichMax = args["enrich-max"] !== undefined ? Number(args["enrich-max"]) : undefined;
  const outBase = (args.out as string) ?? "data/prospects";
  const tierFilter = (args.tier as string | undefined)
    ?.split(",")
    .map((t) => t.trim().toUpperCase())
    .filter(Boolean) as Tier[] | undefined;

  console.log(BAR);
  console.log("🎯 aicrew 東京中小企業 アタックリスト生成（gBizINFO・キーワード横断）");
  console.log(BAR);
  console.log(`📍 都道府県     : ${prefName} (code ${prefecture})`);
  console.log(`🔤 キーワード   : ${keywords.length} 語（${keywords.slice(0, 8).join(" ")} …）`);
  console.log(`👥 従業員制限   : ${empFrom ?? "—"}〜${empTo ?? "—"}（値がある会社のみ適用）`);
  console.log(`💰 資本金制限   : ${capFrom ?? "—"}〜${capTo ?? "—"}（値がある会社のみ適用）`);
  console.log(`📄 1語あたり    : 最大 ${maxPages} ページ  delay=${delay}ms`);
  console.log(`🔬 詳細補完     : ${noEnrich ? "OFF（サイズ/年商は取得しない）" : `ON${enrichMax ? `（上位${enrichMax}社）` : "（全件・時間がかかります）"}`}`);
  console.log(`🏭 対象業種     : 製造 / 卸売・商社 / 建設・設備 / 物流・運輸 / 不動産管理（取得後に分類）`);
  console.log(`📦 出力         : ${outBase}.csv / ${outBase}.xlsx`);
  if (tierFilter) console.log(`🔎 ティア絞り   : ${tierFilter.join(", ")}`);
  console.log(`🔧 モード       : ${dryRun ? c("DRY-RUN（APIを叩きません）", 33) : c("LIVE", 32)}`);
  console.log(BAR);

  if (dryRun) {
    console.log(c("DRY-RUN: 上記設定で LIVE 実行すると gBizINFO を検索します。", 33));
    return;
  }

  const token = process.env.GBIZINFO_API_TOKEN;
  if (!token) {
    console.error(
      c("ERROR: GBIZINFO_API_TOKEN が .env.local にありません。", 31) +
        "\n  → https://content.info.gbiz.go.jp/api/ で API 利用申請し、メールのURLでトークンを取得してください。",
    );
    process.exit(1);
  }

  const client = new GbizClient(token, { delayMs: delay });
  const base: GbizSearchParams = { prefecture };
  if (empFrom !== undefined) base.employee_number_from = empFrom;
  if (empTo !== undefined) base.employee_number_to = empTo;
  if (capFrom !== undefined) base.capital_stock_from = capFrom;
  if (capTo !== undefined) base.capital_stock_to = capTo;

  console.log("🔍 gBizINFO 横断検索中...");
  const all: GbizHojin[] = [];
  for (const kw of keywords) {
    try {
      const hits = await client.searchAll({ ...base, name: kw }, { maxPages });
      all.push(...hits);
      console.log(`  「${kw}」: ${hits.length} 件`);
    } catch (e) {
      console.log(`  「${kw}」: ${c("失敗 " + String((e as Error).message), 31)}`);
    }
  }
  console.log(`✅ 取得（延べ）: ${all.length} 法人`);

  // 法人番号で重複排除（詳細補完を無駄打ちしないため、ここで先に重複を落とす）
  const seen = new Set<string>();
  const uniqueHits: GbizHojin[] = [];
  for (const h of all) {
    const key = h.corporate_number ?? `${h.name ?? ""}|${h.location ?? ""}`;
    if (seen.has(key)) continue;
    seen.add(key);
    uniqueHits.push(h);
  }

  // 対象業種だけ残す（lean データで分類 → 補完対象を絞る）
  const targetHits = uniqueHits.filter((h) => classifyIndustry(h) !== null);
  console.log(`🧹 重複排除＋対象業種: ${targetHits.length} 法人（補完前）`);

  // 詳細API で従業員数・資本金・財務（売上高）を補完
  let scored: GbizHojin[] = targetHits;
  if (!noEnrich && targetHits.length > 0) {
    const cap = enrichMax !== undefined ? Math.min(enrichMax, targetHits.length) : targetHits.length;
    const toEnrich = targetHits.slice(0, cap);
    console.log(`🔬 詳細補完中... ${cap} 社（残り ${targetHits.length - cap} 社は補完なし）`);
    const enriched = await client.enrich(toEnrich, {
      onProgress: (done, total) => {
        if (done % 50 === 0 || done === total) console.log(`  補完 ${done}/${total}`);
      },
    });
    scored = [...enriched, ...targetHits.slice(cap)];
  }

  const prospects: Prospect[] = [];
  let dropped = 0;
  for (const h of scored) {
    const p = toProspect(h);
    if (p) prospects.push(p);
    else dropped++;
  }
  let ranked = dedupeAndRank(prospects);
  if (tierFilter) ranked = ranked.filter((p) => tierFilter.includes(p.priority_tier));

  mkdirSync(dirname(resolve(process.cwd(), outBase)), { recursive: true });
  const csvPath = resolve(process.cwd(), `${outBase}.csv`);
  const xlsxPath = resolve(process.cwd(), `${outBase}.xlsx`);
  writeFileSync(csvPath, toCsv(ranked), "utf8");
  writeFileSync(xlsxPath, toXlsxBuffer(ranked));

  const byTier = ranked.reduce<Record<string, number>>((m, p) => {
    m[p.priority_tier] = (m[p.priority_tier] ?? 0) + 1;
    return m;
  }, {});
  const reported = ranked.filter((p) => p.revenue_confidence === "reported").length;

  console.log(BAR);
  console.log(`🏭 対象業種ヒット : ${prospects.length} 件（対象外除外: ${dropped} 件）`);
  console.log(`🧹 重複排除後     : ${ranked.length} 件`);
  console.log(`🏅 ティア内訳     : ` + ["S", "A", "B", "C"].map((t) => `${t}=${byTier[t] ?? 0}`).join("  "));
  console.log(`📈 年商実測/推定  : 実測 ${reported} 件 / 推定 ${ranked.length - reported} 件`);
  console.log(`📦 出力完了       : ${csvPath}`);
  console.log(`                    ${xlsxPath}`);
  console.log(BAR);
  console.log(c("⚠ 年商・レガシー度は一部推定値です。メール送信前に必ず人手で検証してください。", 33));
}

main().catch((err) => {
  console.error(c(String(err?.stack ?? err), 31));
  process.exit(1);
});
