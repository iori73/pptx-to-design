# PPTX to Figma Plugin

PowerPoint (.pptx) ファイルを Figma に編集可能な状態でインポートする Figma プラグインです。

## 機能

- **テキスト**: フォント、サイズ、色、配置を保持して TextNode として変換
- **図形**: 矩形、楕円、三角形、星形などの基本図形をサポート
- **画像**: PNG, JPEG, GIF などの画像をそのまま取り込み
- **テーブル**: セル構造、背景色、ボーダーを再現
- **グラフ**: チャートをプレースホルダーまたは画像として取り込み
- **グループ**: PowerPoint のグループ階層を維持

## セットアップ

### 必要条件

- Node.js 18+
- npm または yarn
- Figma デスクトップアプリ

### インストール

```bash
# 依存関係のインストール
npm install

# ビルド
npm run build
```

### 開発モード

```bash
# ファイル変更を監視して自動ビルド
npm run watch
```

## Figma での使用方法

1. Figma デスクトップアプリを開く
2. `Plugins` → `Development` → `Import plugin from manifest...` を選択
3. このプロジェクトの `manifest.json` を選択
4. プラグインメニューから「PPTX to Figma」を起動
5. .pptx ファイルをドロップまたは選択
6. 「Import to Figma」をクリック

## プロジェクト構成

```
pptx-to-figma/
├── manifest.json          # Figma プラグイン設定
├── package.json           # 依存関係
├── tsconfig.json          # TypeScript 設定
├── src/
│   ├── code.ts            # メインプラグインロジック
│   ├── ui.html            # UI テンプレート
│   ├── ui.ts              # UI スクリプト
│   ├── parser/
│   │   ├── pptxParser.ts  # .pptx 解析メイン
│   │   ├── slideParser.ts # スライド解析
│   │   └── xmlUtils.ts    # XML 操作ユーティリティ
│   ├── converters/
│   │   ├── textConverter.ts   # テキスト変換
│   │   ├── shapeConverter.ts  # 図形変換
│   │   ├── imageConverter.ts  # 画像変換
│   │   ├── tableConverter.ts  # テーブル変換
│   │   ├── chartConverter.ts  # グラフ変換
│   │   └── groupConverter.ts  # グループ変換
│   └── types/
│       └── pptx.ts        # 型定義
└── dist/                  # ビルド出力
```

## 技術詳細

### PowerPoint ファイル形式

.pptx ファイルは以下の XML ファイルを含む ZIP アーカイブです:

- `ppt/slides/slide*.xml` - スライドコンテンツ
- `ppt/slides/_rels/*.xml.rels` - リレーション情報
- `ppt/media/` - 画像などのメディア
- `ppt/theme/theme*.xml` - テーマ情報

### 座標変換

PowerPoint は EMU (English Metric Units) を使用:
- 1インチ = 914,400 EMU
- Figma は 72 DPI のピクセルを使用
- 変換式: `pixel = emu / 914400 * 72`

## 制限事項

- アニメーション情報は保持されません
- 一部の複雑な図形は矩形に簡略化される場合があります
- EMF/WMF 形式の画像はプレースホルダーに置換されます
- グラフは画像またはプレースホルダーとして取り込まれます

## ライセンス

MIT

