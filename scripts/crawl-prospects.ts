#!/usr/bin/env tsx
/**
 * 東京都内 中小企業（製造/卸売・商社/建設・設備/物流・運輸/不動産管理）の
 * aicrew AI・DX 自動化 営業アタックリストを gBizINFO から生成する CLI。
 *
 * 前提:
 *   .env.local に GBIZINFO_API_TOKEN を設定（無料・無期限トークン）。
 *   取得: https://info.gbiz.go.jp/ → ログイン → マイページ → APIトークン発行
 *
 * 使い方:
 *   pnpm crawl-prospects                       # 既定条件で実行 → data/prospects.csv / .xlsx
 *   pnpm crawl-prospects --emp-from=20 --emp-to=100
 *   pnpm crawl-prospects --cap-from=10000000 --cap-to=300000000
 *   pnpm crawl-prospects --max-pages=20 --delay=400
 *   pnpm crawl-prospects --out=data/tokyo_smb   # 拡張子なしで指定（.csv/.xlsx 両方出力）
 *   pnpm crawl-prospects --tier=S,A             # 出力を指定ティアに絞る
 *   pnpm crawl-prospects:dry                    # API を叩かず設定だけ表示
 *
 * 注意:
 *   - 年商(10-30億)は gBizINFO に財務がある法人のみ実測、無い場合は推定（要人手検証）。
 *   - 生成 CSV/xlsx は .gitignore 済み（実アタックリストはコミットしない）。
 */
import { config } from "dotenv";
import { writeFileSync, mkdirSync } from "node:fs";
import { resolve, dirname } from "node:path";

config({ path: resolve(process.cwd(), ".env.local") });

import {
  GbizClient,
  PREFECTURE_CODE,
  toProspect,
  dedupeAndRank,
  toCsv,
  toXlsxBuffer,
  type GbizSearchParams,
  type Prospect,
  type Tier,
} from "../src/gbizinfo";

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
  const empFrom = Number(args["emp-from"] ?? 20);
  const empTo = Number(args["emp-to"] ?? 100);
  const capFrom = Number(args["cap-from"] ?? 10_000_000); // 1千万
  const capTo = Number(args["cap-to"] ?? 300_000_000); // 3億
  const maxPages = Number(args["max-pages"] ?? 15);
  const delay = Number(args["delay"] ?? 350);
  const outBase = (args.out as string) ?? "data/prospects";
  const tierFilter = (args.tier as string | undefined)
    ?.split(",")
    .map((t) => t.trim().toUpperCase())
    .filter(Boolean) as Tier[] | undefined;

  const params: GbizSearchParams = {
    prefecture,
    employee_number_from: empFrom,
    employee_number_to: empTo,
    capital_stock_from: capFrom,
    capital_stock_to: capTo,
    exist_flg: true,
  };

  console.log(BAR);
  console.log("🎯 aicrew 東京中小企業 アタックリスト生成（gBizINFO）");
  console.log(BAR);
  console.log(`📍 都道府県     : ${prefName} (code ${prefecture})`);
  console.log(`👥 従業員数     : ${empFrom}〜${empTo} 人`);
  console.log(`💰 資本金       : ${(capFrom / 1e8).toFixed(2)}〜${(capTo / 1e8).toFixed(2)} 億円（年商10-30億の代理指標）`);
  console.log(`🏭 対象業種     : 製造 / 卸売・商社 / 建設・設備 / 物流・運輸 / 不動産管理（取得後にキーワード分類）`);
  console.log(`📄 最大ページ   : ${maxPages}（1ページ最大100件）  delay=${delay}ms`);
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
        "\n  → https://info.gbiz.go.jp/ でログインし、マイページからトークンを無料発行してください。",
    );
    process.exit(1);
  }

  const client = new GbizClient(token, { delayMs: delay });

  console.log("🔍 gBizINFO 検索中...");
  const hits = await client.searchAll(params, {
    maxPages,
    onPage: (page, count) => console.log(`  page ${page}: ${count} 件`),
  });
  console.log(`✅ 取得: ${hits.length} 法人（業種分類前）`);

  // Prospect 化（対象業種にヒットしないものは除外）
  const prospects: Prospect[] = [];
  let dropped = 0;
  for (const h of hits) {
    const p = toProspect(h);
    if (p) prospects.push(p);
    else dropped++;
  }
  let ranked = dedupeAndRank(prospects);
  if (tierFilter) ranked = ranked.filter((p) => tierFilter.includes(p.priority_tier));

  // 出力
  mkdirSync(dirname(resolve(process.cwd(), outBase)), { recursive: true });
  const csvPath = resolve(process.cwd(), `${outBase}.csv`);
  const xlsxPath = resolve(process.cwd(), `${outBase}.xlsx`);
  writeFileSync(csvPath, toCsv(ranked), "utf8");
  writeFileSync(xlsxPath, toXlsxBuffer(ranked));

  // サマリ
  const byTier = ranked.reduce<Record<string, number>>((m, p) => {
    m[p.priority_tier] = (m[p.priority_tier] ?? 0) + 1;
    return m;
  }, {});
  const reported = ranked.filter((p) => p.revenue_confidence === "reported").length;

  console.log(BAR);
  console.log(`🏭 対象業種ヒット : ${prospects.length} 件（対象外除外: ${dropped} 件）`);
  console.log(`🧹 重複排除後     : ${ranked.length} 件`);
  console.log(
    `🏅 ティア内訳     : ` +
      ["S", "A", "B", "C"].map((t) => `${t}=${byTier[t] ?? 0}`).join("  "),
  );
  console.log(`📈 年商実測/推定  : 実測 ${reported} 件 / 推定 ${ranked.length - reported} 件`);
  console.log(`📦 出力完了       : ${csvPath}`);
  console.log(`                    ${xlsxPath}`);
  console.log(BAR);
  console.log(
    c(
      "⚠ 年商・レガシー度は一部推定値です。メール送信前に必ず人手で検証してください。",
      33,
    ),
  );
}

main().catch((err) => {
  console.error(c(String(err?.stack ?? err), 31));
  process.exit(1);
});
