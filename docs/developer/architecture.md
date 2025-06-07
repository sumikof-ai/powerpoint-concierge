# システムアーキテクチャ

PowerPoint Conciergeの技術的なアーキテクチャとシステム設計について説明します。

## 🏗️ 全体アーキテクチャ

### システム構成概要
```
┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐
│                     │    │                     │    │                     │
│    PowerPoint       │◄──►│   Office Add-in     │◄──►│    OpenAI API       │
│      Client         │    │    (React/TS)       │    │      Service        │
│                     │    │                     │    │                     │
└─────────────────────┘    └─────────────────────┘    └─────────────────────┘
                                       │
                                       ▼
                           ┌─────────────────────┐
                           │                     │
                           │   Browser Storage   │
                           │   (LocalStorage)    │
                           │                     │
                           └─────────────────────┘
```

### 技術スタック
- **フロントエンド**: React 18 + TypeScript
- **UIライブラリ**: Fluent UI v9
- **Office統合**: Office.js (PowerPoint.js)
- **AIサービス**: OpenAI GPT API
- **バンドラー**: Webpack 5
- **開発環境**: Node.js + npm

## 📁 プロジェクト構造

```
powerpoint-concierge/
├── src/
│   ├── commands/               # Ribbon commands
│   │   ├── commands.html      # Commands UI
│   │   └── commands.ts        # Commands logic
│   ├── services/              # Business logic layer
│   │   ├── openai.service.ts  # OpenAI API integration
│   │   └── powerpoint/        # PowerPoint services
│   │       ├── core/          # Core slide generation
│   │       ├── template/      # Template system
│   │       ├── theme/         # Theme management
│   │       └── index.ts       # Service exports
│   ├── taskpane/              # Main UI components
│   │   ├── components/        # React components
│   │   │   ├── chat/          # Chat interface
│   │   │   ├── outline/       # Outline editor
│   │   │   └── template/      # Template management
│   │   ├── index.tsx          # Entry point
│   │   └── taskpane.html      # HTML template
│   └── types/                 # TypeScript type definitions
├── assets/                    # Static assets
├── docs/                      # Documentation
├── config/                    # Build configuration
├── manifest.xml               # Office Add-in manifest
└── package.json               # Dependencies and scripts
```

## 🔧 コアサービス設計

### 1. PowerPointService
```typescript
class PowerPointService {
  // 主要な責務
  - スライド生成の統合管理
  - AI詳細化とテンプレート機能の切り替え
  - エラーハンドリングとフォールバック
  - 進捗管理とユーザーフィードバック
}
```

### 2. OpenAIService
```typescript
class OpenAIService {
  // 主要な責務
  - OpenAI API との通信管理
  - プロンプト生成と応答パース
  - レート制限とエラー処理
  - 設定管理（APIキー、モデル選択）
}
```

### 3. SlideContentGenerator
```typescript
class SlideContentGenerator {
  // 主要な責務
  - 個別スライドの詳細化
  - コンテキスト情報の構築
  - 品質検証とフォールバック
  - バッチ処理とエラー復旧
}
```

## 🎨 UI/UXアーキテクチャ

### Component階層
```
App
├── WorkflowManager           # ワークフロー状態管理
├── ChatInput                 # メインインターフェース
│   ├── ThemeSettings        # テーマ設定
│   ├── TemplateSelector     # テンプレート選択
│   └── TemplateManager      # テンプレート管理
├── OutlineEditor            # アウトライン編集
│   └── SlideEditor         # 個別スライド編集
└── Settings                 # アプリケーション設定
```

### 状態管理パターン
```typescript
// React Hooks を使用したローカル状態管理
interface ChatInputState {
  currentStep: WorkflowStep;
  currentOutline: PresentationOutline | null;
  selectedTemplate: TemplateInfo | null;
  useTemplateGeneration: boolean;
  generationProgress: string;
  // ...その他の状態
}
```

## 🔄 データフロー

### 1. プレゼンテーション生成フロー
```
User Input → OpenAI API → Outline Generation → 
User Review → Template/Theme Selection → 
AI Enhancement/Template Optimization → 
PowerPoint Generation → Complete
```

### 2. AI詳細化フロー
```typescript
// Phase 1: Analysis
analyzeOutline(outline) → contextBuilding

// Phase 2: Enhancement
for each slide {
  generateDetailedContent(slide, context)
  validateContent(content)
  applyFallbackIfNeeded(content)
}

// Phase 3: PowerPoint Creation
createPowerPointSlides(detailedContent, theme)
```

### 3. テンプレートフロー
```typescript
// Template Selection
analyzeUserInput(input) → recommendTemplates()

// Template Application
adaptOutlineToTemplate(outline, template) →
generateTemplateOptimizedContent(adaptedOutline) →
createSlidesFromTemplate(content, template)
```

## 🏛️ アーキテクチャパターン

### レイヤードアーキテクチャ
```
┌─────────────────────────────────────┐
│         Presentation Layer          │  React Components
├─────────────────────────────────────┤
│         Application Layer           │  Workflow Management
├─────────────────────────────────────┤
│          Service Layer              │  Business Logic Services
├─────────────────────────────────────┤
│         Integration Layer           │  Office.js, OpenAI API
└─────────────────────────────────────┘
```

### 関心の分離
- **UI Layer**: ユーザーインターフェース、状態表示
- **Logic Layer**: ビジネスロジック、ワークフロー制御
- **Service Layer**: 外部API統合、データ変換
- **Storage Layer**: ローカルストレージ、設定管理

## 🔌 統合パターン

### Office.js Integration
```typescript
// PowerPoint API 呼び出しパターン
return new Promise((resolve, reject) => {
  PowerPoint.run(async (context) => {
    try {
      // PowerPoint操作の実行
      const slides = context.presentation.slides;
      // ... 操作内容
      await context.sync();
      resolve(result);
    } catch (error) {
      reject(error);
    }
  });
});
```

### OpenAI API Integration
```typescript
// API呼び出しとエラーハンドリング
async callDetailedContentAPI(slide, context, options) {
  try {
    const response = await this.openAIService.sendRequest(messages);
    return this.parseResponse(response);
  } catch (error) {
    console.error('API call failed:', error);
    return this.createFallbackContent(slide);
  }
}
```

## 📊 パフォーマンス最適化

### 1. API呼び出し最適化
```typescript
// レート制限対応
private async delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// バッチ処理での並列制御
for (let i = 0; i < slides.length; i++) {
  await processSlide(slides[i]);
  if (i < slides.length - 1) {
    await this.delay(500); // Rate limiting
  }
}
```

### 2. メモリ管理
```typescript
// 大量データ処理時のメモリ効率化
class SlideContentGenerator {
  private processInBatches(slides: SlideOutline[], batchSize = 5) {
    // バッチ処理による メモリ使用量制御
  }
}
```

### 3. キャッシュ戦略
```typescript
// テンプレートライブラリのキャッシュ
private loadTemplateLibrary(): TemplateLibrary {
  const cached = localStorage.getItem('template-library');
  return cached ? JSON.parse(cached) : this.initializeDefault();
}
```

## 🛡️ エラーハンドリング

### エラー処理の階層
```
1. UI Level Error Handling
   ├── User feedback (MessageBar, error states)
   └── Graceful degradation

2. Service Level Error Handling
   ├── API error recovery
   ├── Fallback mechanisms
   └── Progress tracking

3. Integration Level Error Handling
   ├── Network error handling
   ├── API rate limiting
   └── Office.js compatibility
```

### フォールバック戦略
```typescript
// AI詳細化失敗時のフォールバック
async generateWithErrorHandling(outline, options, onProgress?, onError?) {
  for (const slide of outline.slides) {
    try {
      const detailed = await this.generateDetailedSlideContent(slide);
      results.push(detailed);
    } catch (error) {
      onError?.(slideIndex, error);
      const fallback = this.createFallbackContent(slide);
      results.push(fallback);
    }
  }
}
```

## 🔒 セキュリティ設計

### データプライバシー
- **ローカル処理**: テンプレート情報はブラウザローカルストレージ
- **API通信**: OpenAI APIの利用規約に準拠
- **機密データ**: ユーザーデータをサーバーに保存しない

### API Key管理
```typescript
// セキュアなAPIキー管理
class OpenAIService {
  private apiKey: string = '';
  
  constructor(settings: OpenAISettings) {
    this.apiKey = settings.apiKey; // メモリのみに保持
  }
}
```

## 🚀 拡張性設計

### プラグインアーキテクチャ
```typescript
// 将来的な機能拡張のための設計
interface ServiceProvider {
  initialize(): Promise<void>;
  generateContent(input: any): Promise<any>;
  validateCapabilities(): boolean;
}

// テンプレートシステムの拡張
interface TemplateProvider {
  getTemplates(): TemplateInfo[];
  registerTemplate(template: TemplateInfo): Promise<void>;
  analyzeTemplate(file: File): Promise<TemplateAnalysisResult>;
}
```

### 設定可能なアーキテクチャ
```typescript
// 設定による動作制御
interface SystemConfiguration {
  aiProvider: 'openai' | 'azure' | 'custom';
  templateSource: 'local' | 'remote' | 'hybrid';
  cachingStrategy: 'aggressive' | 'conservative' | 'none';
  errorRecovery: 'immediate' | 'delayed' | 'manual';
}
```

## 📈 モニタリング・分析

### パフォーマンス指標
```typescript
// 処理時間の測定
class PerformanceMonitor {
  private metrics = new Map<string, number>();
  
  startTimer(operation: string): void;
  endTimer(operation: string): number;
  getMetrics(): Record<string, number>;
}
```

### エラー追跡
```typescript
// エラー情報の収集
interface ErrorReport {
  timestamp: Date;
  operation: string;
  errorType: string;
  message: string;
  context: Record<string, any>;
}
```

このアーキテクチャにより、PowerPoint Conciergeは**拡張性、保守性、パフォーマンス**を兼ね備えたシステムとして設計されています。