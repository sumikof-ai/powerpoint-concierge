# PowerPoint Concierge

<div align="center">

![PowerPoint Concierge Logo](./assets/logo-filled.png)

**AIを活用したプロフェッショナルなPowerPointプレゼンテーション自動生成ツール**

[![Office Add-in](https://img.shields.io/badge/Office-Add--in-0078D4?logo=microsoft-office&logoColor=white)](https://docs.microsoft.com/en-us/office/dev/add-ins/)
[![React](https://img.shields.io/badge/React-18.0-61DAFB?logo=react&logoColor=white)](https://reactjs.org/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![OpenAI](https://img.shields.io/badge/OpenAI-GPT--3.5--turbo-412991?logo=openai&logoColor=white)](https://openai.com/)

</div>

## 🚀 概要

PowerPoint Conciergeは、**AI技術を活用してプロフェッショナルなPowerPointプレゼンテーションを自動生成**するOffice Add-inです。簡単な要件入力から、説明資料として使える詳細なコンテンツを含むスライドを効率的に作成できます。

### ✨ 主な特徴

- **🧠 AI 3段階生成**: アウトライン → 詳細化 → PowerPoint作成
- **📝 説明資料レベルのコンテンツ**: 各スライドを個別に詳細化し、自立して理解できる内容を生成
- **🎨 美しいデザイン**: 3種類のテーマ（ライト/ダーク/カラフル）とフォントサイズ自動適用
- **📋 テンプレート機能**: 独自のデザインテンプレートを作成・管理・再利用
- **🔧 エラー自動復旧**: 生成エラー時の自動フォールバック機能
- **🎯 進捗管理**: リアルタイムな生成進捗表示とワークフロー管理

## 🎯 使用例

### ビフォー・アフター

**入力例:**
```
「営業戦略についてのプレゼンテーションを作成してください」
```

**AI詳細化後の結果:**
- **アウトライン**: 5分で構造化されたプレゼンテーション構成
- **詳細化**: 各スライドが3-5倍詳細化され、具体例・データ・手順を含む
- **PowerPoint**: 美しくデザインされたスライドとして自動生成

| 項目 | 従来の手作業 | PowerPoint Concierge |
|------|--------------|----------------------|
| **時間** | 3-5時間 | **15-30分** |
| **品質** | スキル依存 | **一定品質を保証** |
| **詳細さ** | 簡潔な箇条書き | **説明資料レベル** |
| **一貫性** | 人による差 | **AI による統一性** |

## 🚀 クイックスタート

### 1. 前提条件
```bash
Node.js >= 16.0.0
npm >= 8.0.0
PowerPoint (Microsoft 365, 2019, 2021)
OpenAI APIキー
```

### 2. インストール・セットアップ
```bash
# リポジトリをクローン
git clone https://github.com/your-username/powerpoint-concierge.git
cd powerpoint-concierge

# 依存関係をインストール
npm install

# 開発用証明書を生成（初回のみ）
npm run setup:certs

# 開発ビルド
npm run build:dev
```

### 3. アドインの起動
```bash
# PowerPointでアドインを起動
npm start

# または、開発サーバーのみ起動
npm run dev-server
```

### 4. 初期設定
1. PowerPointで「挿入」→「アドイン」→「PowerPoint Concierge」を選択
2. 設定タブでOpenAI APIキーを入力
3. 初回利用準備完了！

## 📚 ドキュメント

### 📖 利用者向けガイド
- **[はじめに](./docs/user/getting-started.md)** - 基本的な使い方と初期設定
- **[AI詳細化機能](./docs/user/ai-enhancement.md)** - 詳細化機能の完全ガイド
- **[テンプレート管理](./docs/user/template-management.md)** - テンプレート作成・活用方法
- **[テーマ設定](./docs/user/theme-settings.md)** - 色・フォント設定ガイド
- **[トラブルシューティング](./docs/user/troubleshooting.md)** - よくある問題の解決方法
- **[FAQ](./docs/user/faq.md)** - よくある質問と回答

### 🔧 開発者向けガイド
- **[システムアーキテクチャ](./docs/developer/architecture.md)** - 技術的なシステム設計
- **[開発ガイド](./docs/developer/development-guide.md)** - 開発環境構築から実装まで
- **[APIリファレンス](./docs/developer/api-reference.md)** - 主要APIとサービスクラス

## 🛠️ 技術仕様

### サポート環境
- **PowerPoint**: Microsoft 365, 2019, 2021, PowerPoint Online
- **OS**: Windows 10/11, macOS 10.15+
- **ブラウザ**: Microsoft Edge, Chrome, Safari（最新版推奨）

### 技術スタック
- **フロントエンド**: React 18 + TypeScript 5.0
- **UIライブラリ**: Microsoft Fluent UI v9
- **Office統合**: Office.js (PowerPoint.js API)
- **AIサービス**: OpenAI GPT-3.5-turbo API
- **ビルドツール**: Webpack 5 + Babel

## 💡 主な機能

### 🧠 AI詳細化システム
簡潔なアウトラインを**説明資料として使える詳細なコンテンツ**に変換します。

- **文字数3-5倍拡張**: 簡潔な項目を詳細な説明に変換
- **具体例・データ自動追加**: ビジネス現場で使える実用的な内容
- **前後関係考慮**: スライド間の一貫性とストーリー性を保持

### 🎨 テーマ・テンプレートシステム
- **3種類のテーマ**: ライト/ダーク/カラフル
- **独自テンプレート作成**: 現在のプレゼンテーションから自動抽出
- **スマート推奨**: ユーザー入力に基づく最適テンプレート提案

### 🔧 ワークフロー管理
- **3段階ワークフロー**: チャット → アウトライン編集 → スライド生成
- **リアルタイム進捗**: 各フェーズの進行状況を表示
- **エラー処理**: 問題発生時の自動復旧とユーザー通知

## 🤝 コントリビューション

PowerPoint Conciergeの改善にご協力いただける方を歓迎します！

### 貢献方法
1. **🐛 バグ報告**: [Issues](https://github.com/your-username/powerpoint-concierge/issues) でバグを報告
2. **💡 機能提案**: 新機能のアイデアを Issues で提案
3. **📝 ドキュメント改善**: 誤字脱字の修正、説明の改善
4. **💻 コード貢献**: Pull Request での機能追加・修正

## 📞 サポート

- **技術的な問題**: [GitHub Issues](https://github.com/your-username/powerpoint-concierge/issues)
- **一般的な質問**: [FAQ](./docs/user/faq.md) をまずご確認ください
- **完全なドキュメント**: [docs フォルダ](./docs/) に詳細ガイド

## 📄 ライセンス

このプロジェクトは [MIT License](./LICENSE) の下で公開されています。

---

<div align="center">

**PowerPoint Concierge で、プレゼンテーション作成を革新しましょう！** 🚀

</div>