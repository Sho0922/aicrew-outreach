# aicrew-outreach

aicrew 生成AI/DX導入伴走支援の営業アタック先（東京都内・従業員100人以下・年商10〜30億・レガシー中小企業）を
gBizINFO からクローリングして生成するスタンドアロンツール。hempets-sales から分離。

## 使い方
    npm install
    cp .env.example .env.local
    npm run crawl-prospects:dry
    npm run crawl-prospects

対象業種: 製造 / 卸売・商社 / 建設・設備 / 物流・運輸 / 不動産管理。仕様は docs/spec.md。
