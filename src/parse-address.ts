/**
 * 日本の住所文字列を都道府県 / 市区町村 / 残り に分解する。
 *
 * 期待入力例:
 *   "日本、〒106-0041 東京都港区麻布台1丁目5−9 1F"
 *   "東京都港区麻布台1丁目5−9 1F"
 *   "〒060-0001 北海道札幌市中央区北1条西2-1-1"
 *
 * パース失敗時は { prefecture: null, city: null, address_rest: raw } を返す。
 *
 * --- サンプルケース (見落としやすいもの中心) ---
 *  1. "日本、〒106-0041 東京都港区麻布台1丁目5−9 1F"
 *      → 東京都 / 港区 / 麻布台1丁目5−9 1F
 *  2. "東京都八王子市横川町1234-5"
 *      → 東京都 / 八王子市 / 横川町1234-5
 *  3. "北海道札幌市中央区北1条西2-1-1"
 *      → 北海道 / 札幌市中央区 / 北1条西2-1-1
 *  4. "大阪府大阪市北区梅田1-1-1"
 *      → 大阪府 / 大阪市北区 / 梅田1-1-1
 *  5. "京都府京都市左京区一乗寺向畑町1-1"
 *      → 京都府 / 京都市左京区 / 一乗寺向畑町1-1
 *  6. "神奈川県横浜市青葉区美しが丘1-1-1"
 *      → 神奈川県 / 横浜市青葉区 / 美しが丘1-1-1
 *  7. "沖縄県中頭郡北谷町美浜2-1"
 *      → 沖縄県 / 中頭郡北谷町 / 美浜2-1
 *  8. "千葉県市原市五井1234"
 *      → 千葉県 / 市原市 / 五井1234
 *  9. "和歌山県西牟婁郡白浜町1-1"
 *      → 和歌山県 / 西牟婁郡白浜町 / 1-1
 * 10. "東京都八丈島八丈町大賀郷1-1"
 *      → 東京都 / 八丈島八丈町 / 大賀郷1-1
 */

const PREFECTURES = [
  "北海道",
  "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
  "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
  "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
  "岐阜県", "静岡県", "愛知県", "三重県",
  "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
  "鳥取県", "島根県", "岡山県", "広島県", "山口県",
  "徳島県", "香川県", "愛媛県", "高知県",
  "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県",
  "沖縄県",
] as const;

const PREFECTURE_RE = new RegExp(`(${PREFECTURES.join("|")})`);
// 「〒XXX-XXXX」「〒XXXXXXX」両対応（全角ハイフン含む）
const POSTAL_RE = /〒\s*\d{3}[\-‐-―－]?\d{4}/;
// 市区町村抽出。
// 優先順:
//   1. 〇〇郡〇〇町(村)  (郡 つき)
//   2. 〇〇市〇〇区      (政令市の区)
//   3. 〇〇[市区町村]    (それ以外。市原市・八丈島八丈町のように先頭に市/町を含むケースに対応するため lazy + .+?)
const CITY_RE = /^(.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村])/;

export type ParsedAddress = {
  prefecture: string | null;
  city: string | null;
  address_rest: string;
};

export function parseJapaneseAddress(raw: string | null | undefined): ParsedAddress {
  const original = String(raw ?? "").trim();
  if (!original) return { prefecture: null, city: null, address_rest: "" };

  // 1) 「日本、」「日本 」プレフィックス除去
  let s = original.replace(/^日本[、,\s]*/, "");
  // 2) 郵便番号除去
  s = s.replace(POSTAL_RE, "").trim();
  // 3) 半角化（必要最低限：全角空白だけ）
  s = s.replace(/　/g, " ").trim();

  const prefMatch = s.match(PREFECTURE_RE);
  if (!prefMatch) {
    return { prefecture: null, city: null, address_rest: original };
  }
  const prefecture = prefMatch[1];
  const afterPref = s.slice(s.indexOf(prefecture) + prefecture.length);

  const cityMatch = afterPref.match(CITY_RE);
  if (!cityMatch) {
    return {
      prefecture,
      city: null,
      address_rest: afterPref.trim(),
    };
  }
  const city = cityMatch[1];
  const address_rest = afterPref.slice(city.length).trim();
  return { prefecture, city, address_rest };
}

/**
 * 電話番号の正規化:
 *  - 全角→半角、各種ダッシュ→ASCIIハイフン
 *  - 括弧・空白除去
 *  - 末尾の余分なハイフン除去
 */
export function normalizePhone(raw: string | null | undefined): string | null {
  const s = String(raw ?? "").trim();
  if (!s) return null;
  const z2h = s.replace(/[０-９]/g, (c) =>
    String.fromCharCode(c.charCodeAt(0) - 0xfee0),
  );
  return z2h
    .replace(/[‐－—–ー―‑]/g, "-")
    .replace(/[（）()\s]/g, "")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-")
    || null;
}
