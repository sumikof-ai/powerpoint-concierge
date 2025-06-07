# 開発ガイド

PowerPoint Conciergeの開発環境構築から実装まで、開発者向けの詳細ガイドです。

## 🚀 開発環境セットアップ

### 前提条件
```bash
# 必要なソフトウェア
Node.js >= 16.0.0
npm >= 8.0.0
PowerPoint (Microsoft 365, 2019, 2021)
Git
```

### 環境構築手順

#### 1. リポジトリのクローン
```bash
git clone https://github.com/your-username/powerpoint-concierge.git
cd powerpoint-concierge
```

#### 2. 依存関係のインストール
```bash
# パッケージのインストール
npm install

# 開発者証明書の生成（初回のみ）
npm run setup:certs
```

#### 3. 環境設定
```bash
# .env ファイルの作成（任意）
cp .env.example .env
```

#### 4. 開発サーバーの起動
```bash
# 開発ビルド
npm run build:dev

# Office Add-in デバッグ開始
npm start

# または、別ターミナルで
npm run dev-server
```

### 開発者証明書の設定
```bash
# 証明書の生成（Linux/macOSの場合、管理者権限が必要）
sudo npm run setup:certs

# Windowsの場合
npm run setup:certs
```

## 🏗️ プロジェクト構造詳細

### フォルダ構成
```
src/
├── commands/                   # Ribbon UI commands
│   ├── commands.html          # Commands taskpane HTML
│   └── commands.ts            # Commands functionality
├── services/                   # Business logic services
│   ├── openai.service.ts      # OpenAI API integration
│   └── powerpoint/            # PowerPoint-related services
│       ├── core/              # Core functionality
│       │   ├── ContentRenderer.ts     # Slide content rendering
│       │   ├── SlideContentGenerator.ts # AI content generation
│       │   ├── SlideFactory.ts        # Slide creation
│       │   └── ThemeApplier.ts        # Theme application
│       ├── template/          # Template system
│       │   ├── TemplateAdaptationService.ts
│       │   └── TemplateBasedGenerationService.ts
│       ├── theme/             # Theme management
│       │   ├── ThemeAnalyzer.ts
│       │   └── ThemeService.ts
│       ├── powerpoint.service.ts      # Main PowerPoint service
│       ├── types.ts           # PowerPoint-related types
│       └── index.ts           # Service exports
├── taskpane/                  # Main UI components
│   ├── components/            # React components
│   │   ├── chat/              # Chat interface
│   │   │   ├── ChatInput.tsx         # Main chat interface
│   │   │   ├── ThemeSettings.tsx     # Theme configuration
│   │   │   └── WorkflowManager.tsx   # Workflow state management
│   │   ├── outline/           # Outline editing
│   │   │   ├── OutlineEditor.tsx     # Main outline editor
│   │   │   └── SlideEditor.tsx       # Individual slide editor
│   │   ├── template/          # Template management
│   │   │   ├── TemplateManager.tsx   # Template management UI
│   │   │   └── TemplateSelector.tsx  # Template selection UI
│   │   ├── App.tsx            # Root application component
│   │   ├── Header.tsx         # App header
│   │   ├── Settings.tsx       # Settings panel
│   │   └── types.ts           # UI-related types
│   ├── index.tsx              # React app entry point
│   ├── taskpane.html          # Main taskpane HTML
│   └── taskpane.ts            # Taskpane initialization
├── types/                     # Global type definitions
config/                        # Build configuration
├── webpack.config.js          # Webpack configuration
assets/                        # Static assets (icons, images)
docs/                          # Documentation
manifest.xml                   # Office Add-in manifest
```

## 🛠️ 開発ワークフロー

### 1. 新機能開発の流れ

#### Feature Branch作成
```bash
# 新機能用ブランチの作成
git checkout -b feature/new-feature-name
```

#### 開発・テスト
```bash
# 開発サーバー起動
npm run dev-server

# 別ターミナルでPowerPoint連携テスト
npm start

# TypeScript型チェック
npm run typecheck

# Lint実行
npm run lint

# フォーマット
npm run prettier
```

#### ビルドテスト
```bash
# 本番ビルド
npm run build

# 開発ビルド
npm run build:dev
```

### 2. コード品質管理

#### TypeScript設定
```json
// tsconfig.json の重要な設定
{
  "compilerOptions": {
    "strict": true,
    "noImplicitAny": true,
    "strictNullChecks": true,
    "strictFunctionTypes": true
  }
}
```

#### ESLint設定
```javascript
// .eslintrc.json
{
  "extends": [
    "@microsoft/eslint-config-spfx/lib/profiles/react",
    "@microsoft/eslint-config-spfx/lib/profiles/typescript"
  ]
}
```

### 3. デバッグ方法

#### ブラウザDevTools
```typescript
// デバッグ用ログ出力
console.log('🔍 デバッグ情報:', data);
console.error('❌ エラー:', error);

// 詳細化進捗のデバッグ
onProgress?: (phase: string, current: number, total: number, message: string) => {
  console.log(`📊 進捗: ${phase} ${current}/${total} - ${message}`);
}
```

#### Office.js デバッグ
```typescript
// PowerPoint.run のエラーハンドリング
PowerPoint.run(async (context) => {
  try {
    // PowerPoint操作
    await context.sync();
  } catch (error) {
    console.error('PowerPoint API Error:', error);
    throw error;
  }
});
```

## 📝 実装ガイド

### 1. 新しいサービスの追加

#### サービスクラスの作成
```typescript
// src/services/example.service.ts
export class ExampleService {
  constructor(private config: ExampleConfig) {}

  public async processData(input: InputType): Promise<OutputType> {
    try {
      // 処理ロジック
      return result;
    } catch (error) {
      console.error('ExampleService error:', error);
      throw error;
    }
  }
}
```

#### 型定義の追加
```typescript
// src/types/example.types.ts
export interface ExampleConfig {
  apiKey: string;
  options: ExampleOptions;
}

export interface InputType {
  data: string;
  parameters: Record<string, any>;
}

export interface OutputType {
  result: any;
  metadata: any;
}
```

### 2. 新しいReactコンポーネントの追加

#### コンポーネントの作成
```typescript
// src/taskpane/components/example/ExampleComponent.tsx
import * as React from "react";
import { useState, useEffect } from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";

interface ExampleComponentProps {
  data: any;
  onAction: (result: any) => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  }
});

export const ExampleComponent: React.FC<ExampleComponentProps> = ({
  data,
  onAction
}) => {
  const styles = useStyles();
  const [state, setState] = useState(initialState);

  useEffect(() => {
    // 初期化処理
  }, [data]);

  const handleAction = async () => {
    try {
      const result = await processAction(state);
      onAction(result);
    } catch (error) {
      console.error('Action failed:', error);
    }
  };

  return (
    <div className={styles.container}>
      <Text>Example Component</Text>
      <Button onClick={handleAction}>Action</Button>
    </div>
  );
};

export default ExampleComponent;
```

### 3. API統合の実装

#### OpenAI APIサービスの拡張
```typescript
// src/services/openai.service.ts に追加
export class OpenAIService {
  public async generateCustomContent(
    prompt: string,
    options: CustomOptions
  ): Promise<CustomResponse> {
    try {
      const messages = this.buildCustomPrompt(prompt, options);
      const response = await this.sendRequest(messages);
      return this.parseCustomResponse(response);
    } catch (error) {
      console.error('Custom content generation failed:', error);
      throw error;
    }
  }

  private buildCustomPrompt(prompt: string, options: CustomOptions): any[] {
    return [
      { role: 'system', content: 'カスタムシステムプロンプト' },
      { role: 'user', content: prompt }
    ];
  }

  private parseCustomResponse(response: string): CustomResponse {
    // レスポンスのパース処理
    return parsed;
  }
}
```

## 🧪 テスト戦略

### 1. ユニットテストの設定

#### Jest設定
```javascript
// jest.config.js
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jsdom',
  setupFilesAfterEnv: ['<rootDir>/src/test/setup.ts'],
  moduleNameMapping: {
    '\\.(css|less|scss)$': 'identity-obj-proxy'
  }
};
```

#### テストファイルの作成
```typescript
// src/services/__tests__/openai.service.test.ts
import { OpenAIService } from '../openai.service';

describe('OpenAIService', () => {
  let service: OpenAIService;

  beforeEach(() => {
    service = new OpenAIService({
      apiKey: 'test-key',
      model: 'gpt-3.5-turbo'
    });
  });

  test('should generate structured outline', async () => {
    // テスト実装
    const result = await service.generateStructuredOutline('test input');
    expect(result).toBeDefined();
    expect(result.title).toBeTruthy();
    expect(result.slides).toHaveLength(3);
  });

  test('should handle API errors gracefully', async () => {
    // エラーハンドリングのテスト
    await expect(service.generateStructuredOutline('')).rejects.toThrow();
  });
});
```

### 2. 統合テストの実装

#### PowerPoint APIテスト
```typescript
// src/services/__tests__/powerpoint.service.integration.test.ts
describe('PowerPoint Service Integration', () => {
  test('should create slides in PowerPoint', async () => {
    // Office.js環境でのテスト
    const service = new PowerPointService();
    const testData = createTestSlideData();
    
    await expect(service.generateBulkSlides(testData)).resolves.toBeUndefined();
  });
});
```

### 3. E2Eテストの考慮事項

#### PowerPointアドインのE2Eテスト
```typescript
// Office環境での自動テストは複雑なため、
// 手動テストシナリオを定義
const E2E_TEST_SCENARIOS = [
  'アウトライン生成からスライド作成まで',
  'テンプレート選択と適用',
  'エラー発生時のフォールバック',
  '大量スライドの生成'
];
```

## 🔧 ビルド・デプロイ

### 1. ビルド設定

#### Webpack設定のカスタマイズ
```javascript
// webpack.config.js
const config = {
  entry: {
    taskpane: './src/taskpane/index.tsx',
    commands: './src/commands/commands.ts'
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx'],
    alias: {
      '@': path.resolve(__dirname, 'src')
    }
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      }
    ]
  }
};
```

### 2. 環境別設定

#### 開発環境
```javascript
// webpack.dev.js
module.exports = merge(common, {
  mode: 'development',
  devtool: 'inline-source-map',
  devServer: {
    contentBase: './dist',
    hot: true,
    port: 3000
  }
});
```

#### 本番環境
```javascript
// webpack.prod.js
module.exports = merge(common, {
  mode: 'production',
  optimization: {
    minimize: true,
    splitChunks: {
      chunks: 'all'
    }
  }
});
```

### 3. マニフェスト管理

#### manifest.xml の更新
```xml
<!-- manifest.xml -->
<OfficeApp>
  <Id>12345678-1234-1234-1234-123456789012</Id>
  <Version>1.0.0</Version>
  <ProviderName>Your Organization</ProviderName>
  <DefaultLocale>ja-JP</DefaultLocale>
  <DisplayName DefaultValue="PowerPoint Concierge" />
  <Description DefaultValue="AI-powered presentation generator" />
  
  <Hosts>
    <Host Name="Presentation" />
  </Hosts>
  
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
    </Sets>
  </Requirements>
</OfficeApp>
```

## 📊 パフォーマンス最適化

### 1. バンドルサイズ最適化

#### 動的インポートの活用
```typescript
// 大きなライブラリの遅延読み込み
const loadHeavyFeature = async () => {
  const { HeavyFeature } = await import('./HeavyFeature');
  return HeavyFeature;
};
```

#### Tree Shakingの最適化
```javascript
// webpack.config.js
module.exports = {
  optimization: {
    usedExports: true,
    sideEffects: false
  }
};
```

### 2. APIパフォーマンス

#### リクエスト最適化
```typescript
// バッチ処理によるAPI効率化
class BatchProcessor {
  private queue: ProcessingItem[] = [];
  private batchSize = 5;

  async addToQueue(item: ProcessingItem): Promise<void> {
    this.queue.push(item);
    
    if (this.queue.length >= this.batchSize) {
      await this.processBatch();
    }
  }

  private async processBatch(): Promise<void> {
    const batch = this.queue.splice(0, this.batchSize);
    await Promise.all(batch.map(item => this.processItem(item)));
  }
}
```

## 🔐 セキュリティ考慮事項

### 1. APIキー管理
```typescript
// セキュアなAPIキー処理
class SecureAPIManager {
  private encryptedKey: string = '';

  setAPIKey(key: string): void {
    // メモリ内でのみ保持、永続化しない
    this.encryptedKey = this.encrypt(key);
  }

  private encrypt(data: string): string {
    // 簡易暗号化（本番環境では適切な暗号化を実装）
    return btoa(data);
  }
}
```

### 2. 入力検証
```typescript
// ユーザー入力の検証
class InputValidator {
  static validateUserInput(input: string): ValidationResult {
    if (!input || input.trim().length === 0) {
      return { isValid: false, error: '入力が空です' };
    }

    if (input.length > 5000) {
      return { isValid: false, error: '入力が長すぎます' };
    }

    // XSS対策
    const sanitized = this.sanitizeInput(input);
    return { isValid: true, sanitized };
  }

  private static sanitizeInput(input: string): string {
    return input.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  }
}
```

## 🚀 リリース手順

### 1. バージョン管理
```bash
# バージョンアップ
npm version patch  # パッチリリース
npm version minor  # マイナーリリース  
npm version major  # メジャーリリース
```

### 2. リリースビルド
```bash
# 本番ビルド
npm run build

# マニフェスト検証
npm run validate

# テスト実行
npm test
```

### 3. デプロイメント
```bash
# Office Store用パッケージ作成
npm run package

# または手動配布用
npm run build:production
```

## 📚 参考資料

### Office Add-in開発
- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [PowerPoint JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/powerpoint-add-ins-reference-overview)

### React/TypeScript
- [React Documentation](https://reactjs.org/docs/getting-started.html)
- [TypeScript Handbook](https://www.typescriptlang.org/docs/)
- [Fluent UI Documentation](https://developer.microsoft.com/en-us/fluentui)

### 開発ツール
- [Webpack Documentation](https://webpack.js.org/concepts/)
- [Jest Testing Framework](https://jestjs.io/docs/getting-started)

開発時の質問や問題については、プロジェクトのIssuesまたは開発チームにお問い合わせください。