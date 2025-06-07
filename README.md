# excel_management

Excel（.xlsx）ファイルを CSV に変換する Node.js + TypeScript 製ツールです。

## 機能

- 指定ディレクトリ配下のすべての `.xlsx` ファイルを再帰的に探索し、各シートを CSV ファイルとして出力します
- 変換時、セルの値は Excel の型ごとに適切に文字列化されます
- 出力先ディレクトリは変換前に自動削除されます

## ディレクトリ構成

```
excel_management/
├── docs/                # 変換元のExcelファイルを格納
│   └── sample.xlsx
├── diff/                # 変換後のCSVファイルが出力される（実行時に生成）
├── excelUtils.ts        # Excel変換・ユーティリティ
├── fsUtils.ts           # ファイルシステム操作ユーティリティ
├── index.ts             # エントリーポイント
├── package.json
├── tsconfig.json
└── ...
```

# docs ディレクトリについて

このディレクトリには、変換対象となる Excel（.xlsx）ファイルを配置してください。

> **注意:**
> デフォルトでは `sample.xlsx` というサンプルファイルが入っています。ご自身の変換したい Excel ファイルに差し替えてご利用ください。

- 変換元ファイルを追加・削除しても構いません。
- サブディレクトリも再帰的に探索されます。

---

## 例

```
docs/
  ├── your_data1.xlsx
  ├── your_data2.xlsx
  └── ...
```

## 必要要件

- Node.js
- npm
- TypeScript

## セットアップ

```sh
npm install
```

## 使い方

```sh
npx ts-node index.ts [変換元ディレクトリ] [出力先ディレクトリ]
```

- 例:  
  `npx ts-node index.ts docs diff`
- デフォルト値:  
  変換元: `docs`  
  出力先: `diff`

## スクリプト

- `npm run diff`  
  デフォルトの `docs` → `diff` で変換を実行

## ライセンス

ISC
