#!/usr/bin/env tsx
/**
 * アタックリスト（data/prospects.xlsx）の各社に業種別テンプレートを差し込み、
 * 営業メールの下書きを data/emails.csv（件名・本文付き）として一括生成する。
 *
 * 使い方:
 *   pnpm gen-emails                         # data/prospects.xlsx → data/emails.csv
 *   pnpm gen-emails --tier=S,A              # 指定ティアのみ
 *   pnpm gen-emails --in=data/priority_SA.xlsx --out=data/emails_SA.csv
 *   pnpm gen-emails --limit=50
 *
 * 差出人情報（aicrew）は環境変数 or 既定プレースホルダから差し込む。送信前に必ず実値に。
 *   SENDER_NAME / SENDER_COMPANY / SENDER_ADDRESS / SENDER_TEL / SENDER_EMAIL / OPTOUT_CONTACT
 *
 * ⚠ 特定電子メール法: 送信者の氏名・住所の明示と、オプトアウト（配信停止）導線が必須。
 *   本テンプレートはフッターにこれらを含む。実値を必ず設定すること。
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

const SENDER = {
  name: process.env.SENDER_NAME ?? "{{送信者氏名}}",
  company: process.env.SENDER_COMPANY ?? "aicrew（{{会社名}}）",
  address: process.env.SENDER_ADDRESS ?? "{{住所}}",
  tel: process.env.SENDER_TEL ?? "{{電話}}",
  email: process.env.SENDER_EMAIL ?? "{{メール}}",
  optout: process.env.OPTOUT_CONTACT ?? "本メールへのご返信",
};

/** aicrew 資料の実績値（共通の信頼ブロック）。 */
const PROOF = [
  "・処理時間 最大 −88%",
  "・3年で最大 3,675万円のコスト削減",
  "・AI導入コストは従来の約 1/10",
  "・伴走支援による定着率 100%",
].join("\n");

/** 業種別の件名。 */
function subject(industry: string, company: string): string {
  const map: Record<string, string> = {
    製造: `${company} 御中｜FAX受発注・生産管理の手作業をAIで自動化（処理時間−88%の実績）`,
    "卸売・商社": `${company} 御中｜FAX受注・見積/請求の入力をAIで自動化（伴走型・定着率100%）`,
    "建設・設備": `${company} 御中｜見積・原価・日報の作成をAIで自動化（3年で最大3,675万円削減）`,
    "物流・運輸": `${company} 御中｜配車・伝票・請求処理をAIで自動化（処理時間−88%）`,
    不動産管理: `${company} 御中｜入居者対応・契約更新をAIで自動化（伴走型導入）`,
  };
  return map[industry] ?? `${company} 御中｜現場の手作業をAIで自動化（伴走型のDX支援）`;
}

/** 業種別の課題提起（1段落）。 */
function painParagraph(industry: string, opportunity: string): string {
  const map: Record<string, string> = {
    製造:
      "受発注や生産管理がFAX・電話・Excelに依存し、転記や在庫照合に多くの工数が割かれているケースを多く拝見します。",
    "卸売・商社":
      "FAXでの受注や、見積・請求書の手入力・転記に時間が取られ、担当者に業務が属人化しているケースを多く拝見します。",
    "建設・設備":
      "見積・工程・原価管理や日報・現場写真の整理が紙やExcel中心で、事務作業が現場の負担になっているケースを多く拝見します。",
    "物流・運輸":
      "配車・伝票・請求処理や問い合わせ対応に定型業務が多く、人手と時間が逼迫しているケースを多く拝見します。",
    不動産管理:
      "入居者からの問い合わせ対応や契約更新・紙台帳の管理に、定型的な手作業が積み重なっているケースを多く拝見します。",
  };
  const base = map[industry] ?? "FAXや紙・Excel・属人化した手作業に多くの時間が割かれているケースを多く拝見します。";
  // クローラが付与した自動化余地（自動化の具体）を続けて差し込む
  return opportunity ? `${base}\n具体的には「${opportunity}」といった領域で、生成AIとRPA/OCRによる自動化の余地が大きいと考えております。` : base;
}

function body(r: Row): string {
  const company = String(r["法人名"] ?? "");
  const industry = String(r["業種"] ?? "");
  const rep = r["代表者"] ? `${String(r["代表者"])} 様` : "ご担当者様";
  const opportunity = String(r["自動化余地"] ?? "");
  return [
    `${company}`,
    `${rep}`,
    "",
    `突然のご連絡失礼いたします。生成AI・DX導入の伴走支援を行う ${SENDER.company} の ${SENDER.name} と申します。`,
    "",
    painParagraph(industry, opportunity),
    "",
    "aicrew は「PoC止まりのAI・DXを、現場で動く成果へ」をテーマに、導入から定着まで伴走型で支援しております。",
    PROOF,
    "",
    "貴社の業務に合わせた自動化の進め方を、15分ほどのオンラインでご説明できればと存じます。",
    "ご関心をお持ちいただけましたら、ご都合のよい日程を2〜3つご返信いただけますと幸いです（概要資料を添付しております）。",
    "",
    "何卒よろしくお願い申し上げます。",
    "",
    "─────────────────────",
    `${SENDER.company} ${SENDER.name}`,
    `${SENDER.address}`,
    `TEL: ${SENDER.tel} / Email: ${SENDER.email}`,
    `※本メールは貴社の公開情報をもとにお送りしております。配信停止をご希望の場合は、${SENDER.optout}にてお知らせください。以後の送信を停止いたします。`,
  ].join("\n");
}

function csvEscape(v: string): string {
  return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
}

function main() {
  const args = parseArgs(process.argv);
  const inPath = (args.in as string) ?? "data/prospects.xlsx";
  const outPath = (args.out as string) ?? "data/emails.csv";
  const limit = args.limit !== undefined ? Number(args.limit) : undefined;
  const tierFilter = (args.tier as string | undefined)
    ?.split(",").map((t) => t.trim().toUpperCase()).filter(Boolean);

  // ESM ビルドの XLSX には readFile/writeFile（fs依存）が無いため自前で読む
  const buf = readFileSync(resolve(process.cwd(), inPath));
  const wb = XLSX.read(buf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  let rows = XLSX.utils.sheet_to_json<Row>(ws, { defval: "" });
  if (tierFilter) rows = rows.filter((r) => tierFilter.includes(String(r["優先度"]).toUpperCase()));
  if (limit !== undefined) rows = rows.slice(0, limit);

  const header = ["優先度", "法人名", "業種", "宛先", "件名", "本文"];
  const lines = [header.join(",")];
  for (const r of rows) {
    lines.push(
      [
        String(r["優先度"] ?? ""),
        String(r["法人名"] ?? ""),
        String(r["業種"] ?? ""),
        r["代表者"] ? `${String(r["代表者"])} 様` : "ご担当者様",
        subject(String(r["業種"] ?? ""), String(r["法人名"] ?? "")),
        body(r),
      ].map(csvEscape).join(","),
    );
  }
  const csv = "﻿" + lines.join("\r\n");
  mkdirSync(dirname(resolve(process.cwd(), outPath)), { recursive: true });
  writeFileSync(resolve(process.cwd(), outPath), csv, "utf8");

  console.log(`✅ ${rows.length} 件の下書きを生成: ${resolve(process.cwd(), outPath)}`);
  if (SENDER.name.includes("{{")) {
    console.log("⚠ 差出人情報が未設定です。.env.local に SENDER_NAME / SENDER_COMPANY / SENDER_ADDRESS / SENDER_TEL / SENDER_EMAIL を設定してください（特定電子メール法）。");
  }
  console.log("⚠ 送信前に各社の年商・実態を必ず人手で検証してください（年商は推定値です）。");
}

main();
