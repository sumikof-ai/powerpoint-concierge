# PowerPoint自動生成アドイン - テーマ対応・動的生成対応版開発仕様

## プロジェクト概要

https://github.com/sumikof-ai/powerpoint-concierge.git のリポジトリをベースに、チャット入力からOpenAI APIを活用してPowerPointプレゼンテーションを自動生成するOfficeアドインを開発する。

## 核心機能要件

### 1. テーマ・デザイン対応機能

#### 現在のテーマ情報取得
- **テーマ検出**: 現在開いているPowerPointのテーマ名、カラースキーム、フォント設定を取得
- **レイアウト分析**: 利用可能なスライドレイアウト（タイトルスライド、内容スライド等）の構造を解析
- **デザイン要素識別**: 現在のテーマに含まれるプレースホルダーの種類と配置を特定

#### スマートコンテンツ配置
- **プレースホルダー活用**: テキストボックスではなく、既存のプレースホルダー（タイトル、コンテンツ、サブタイトル等）を優先使用
- **レイアウト自動選択**: スライドの内容タイプ（タイトル、箇条書き、比較、まとめ等）に最適なレイアウトを自動選択
- **テーマ準拠**: 現在のテーマのフォント、色、スタイル設定を維持

### 2. 動的コンテンツ生成フロー

```
Step 1: アウトライン生成 → Step 2: 確認・修正 → Step 3: 動的スライド生成 → Step 4: 個別編集
```

#### Step 1: アウトライン自動生成

入力内容から以下の構造化されたアウトラインを生成：

```json
{
  "title": "プレゼンテーションタイトル",
  "audience": "想定聴衆",
  "objective": "プレゼンテーションの目的",
  "slides": [
    {
      "slideNumber": 1,
      "title": "スライドタイトル",
      "contentType": "title|bullets|comparison|conclusion|chart|image_with_text",
      "keyPoints": ["要点1", "要点2", "要点3"],
      "detailLevel": "basic|detailed|comprehensive"
    }
  ],
  "estimatedDuration": "想定発表時間（分）"
}
```

#### Step 2: アウトライン確認・修正UI

- 生成されたアウトラインをツリー表示
- インラインで編集可能（タイトル、要点の追加・削除・並び替え）
- 各スライドの詳細レベル調整（基本/詳細/包括的）
- 「このアウトラインで作成開始」ボタン
- 「AIに再生成を依頼」機能（追加指示入力可能）

#### Step 3: 動的スライド生成

**3.1 スライド別詳細生成**
各スライド作成時に以下の処理を実行：

1. **コンテンツ詳細化API呼び出し**
   ```json
   {
     "slideContext": {
       "presentationTitle": "全体タイトル",
       "slideTitle": "現在のスライドタイトル",
       "slideNumber": 3,
       "totalSlides": 10,
       "previousSlideContent": "前スライドの概要",
       "nextSlidePreview": "次スライドの予定内容"
     },
     "contentRequirements": {
       "detailLevel": "詳細レベル指定",
       "targetAudience": "聴衆レベル",
       "presentationStyle": "formal|casual|technical"
     },
     "designConstraints": {
       "availableLayouts": ["使用可能レイアウト一覧"],
       "themeColors": ["テーマカラー情報"],
       "maxTextLength": "文字数制限"
     }
   }
   ```

2. **動的コンテンツ生成**
   - 箇条書きの場合：各項目の詳細説明を生成
   - 比較スライドの場合：対比表や説明文を生成
   - まとめスライドの場合：全体の流れを踏まえた総括を生成

3. **レイアウト最適化**
   - 生成されたコンテンツ量に応じてレイアウトを調整
   - 長文の場合は複数スライドに分割提案
   - 図表が必要な場合はプレースホルダーと説明を配置

**3.2 プログレッシブ生成**
- 進捗表示（「スライド 3/10 詳細生成中...」）
- 各スライド完成後にプレビュー表示
- ユーザーが途中で生成を停止/修正可能

#### Step 4: スマート編集機能

- **コンテキスト考慮編集**: 特定スライドの編集時に、前後のスライドとの整合性を保持
- **スタイル一貫性**: 編集後もテーマのスタイル設定を維持
- **再生成オプション**: 「もっと詳しく」「簡潔に」「聴衆レベルを変更」等の指示で部分再生成

### 3. テーマ対応PowerPoint操作仕様

#### 3.1 テーマ情報取得機能

```typescript
interface ThemeInfo {
  name: string;
  colorScheme: {
    accent1: string;
    accent2: string;
    background: string;
    text: string;
  };
  fontScheme: {
    major: string;  // 見出し用フォント
    minor: string;  // 本文用フォント
  };
  availableLayouts: LayoutInfo[];
}

interface LayoutInfo {
  name: string;
  type: 'title' | 'content' | 'comparison' | 'blank';
  placeholders: PlaceholderInfo[];
}

interface PlaceholderInfo {
  type: 'title' | 'subtitle' | 'content' | 'footer';
  position: { x: number, y: number, width: number, height: number };
  textFormat: TextFormatInfo;
}
```

#### 3.2 スマートコンテンツ配置

- **プレースホルダー優先**: `slide.placeholders`を使用してコンテンツを配置
- **階層構造保持**: 箇条書きのインデントレベルを適切に設定
- **テーマ色適用**: `theme.colors`を参照して適切な色を自動選択
- **フォント一貫性**: テーマのフォント設定に従って文字スタイルを適用

#### 3.3 動的レイアウト選択

```typescript
function selectOptimalLayout(contentType: string, contentAmount: number): string {
  const layoutMapping = {
    'title': 'Title Slide',
    'bullets': contentAmount > 100 ? 'Content with Caption' : 'Title and Content',
    'comparison': 'Two Content',
    'conclusion': 'Title and Content',
    'chart': 'Content with Caption'
  };
  return layoutMapping[contentType] || 'Title and Content';
}
```

## 技術仕様

### 開発環境

- **フレームワーク**: Yeoman generator + TypeScript + React
- **Officeアドイン**: Office.js API（最新版）
- **スタイリング**: Fluent UI React v9
- **状態管理**: React Context + useReducer + React Query（API状態管理）

### 設定管理

```typescript
interface AppConfig {
  openai: {
    baseUrl: string;
    apiKey: string;
    model: string;
    maxTokens: number;
    temperature: number;        // 創造性レベル
  };
  generation: {
    detailLevel: 'basic' | 'detailed' | 'comprehensive';
    presentationStyle: 'formal' | 'casual' | 'technical';
    useProgressiveGeneration: boolean;
    maxSlideTextLength: number;
  };
  ui: {
    theme: 'light' | 'dark';
    language: 'ja' | 'en';
    showPreview: boolean;
  };
}
```

### API呼び出し戦略

#### 3段階API呼び出し

1. **アウトライン生成API**
   - プロンプト: 「以下の要件から構造化されたプレゼンテーションアウトラインを生成してください」
   - 出力: JSON形式のアウトライン

2. **スライド詳細生成API**（各スライドごと）
   - プロンプト: 「以下のコンテキストでスライドの詳細内容を生成してください」
   - 出力: 具体的なタイトル、箇条書き、説明文

3. **調整・改善API**（オプション）
   - プロンプト: 「以下のスライドを[指示内容]に従って改善してください」
   - 出力: 改善されたコンテンツ

## 実装優先順位

### Phase 1: テーマ対応基盤

1. テーマ情報取得機能
2. プレースホルダー識別・活用機能
3. 基本的なレイアウト選択機能

### Phase 2: 動的生成コア

1. 段階的API呼び出し機能
2. コンテンツ詳細化機能
3. プログレッシブ生成UI

### Phase 3: 高度機能

1. コンテキスト考慮編集
2. スタイル一貫性保持
3. パフォーマンス最適化

## エラーハンドリング・パフォーマンス要件

### API呼び出し最適化

- **並列処理制限**: 同時API呼び出し数を3以下に制限
- **リトライ機能**: 失敗時の指数バックオフ実装
- **キャッシュ機能**: 類似コンテンツの再利用

### ユーザー体験向上

- **リアルタイム進捗**: 各スライド生成の進捗表示
- **中断・再開機能**: 長時間処理の途中停止・再開
- **プレビュー機能**: 生成中のコンテンツをリアルタイム表示

## 開発時重要ポイント

1. **Office.js テーマAPI**: `Office.context.document.theme`を活用
2. **非同期処理**: 各スライド生成を非同期で実行し、UIをブロックしない
3. **メモリ管理**: 大量のスライド生成時のメモリリーク防止
4. **ユーザビリティ**: 生成過程でのユーザーフィードバック重視
5. **エラー回復**: 部分的な失敗時も他のスライドは正常に生成続行