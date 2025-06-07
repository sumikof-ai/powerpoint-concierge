# API リファレンス

PowerPoint Conciergeの主要なAPIとサービスクラスのリファレンスです。

## 🏗️ Core Services

### PowerPointService

メインのPowerPoint操作サービスクラス。スライド生成、テンプレート管理、テーマ適用を統合管理します。

#### Constructor
```typescript
class PowerPointService {
  constructor()
}
```

#### Main Methods

##### generateSlidesWithDetailedProgress
AI詳細化機能を使用してスライドを生成します。

```typescript
public async generateSlidesWithDetailedProgress(
  outline: PresentationOutline,
  openAIService: OpenAIService,
  options: SlideGenerationOptions = {},
  onDetailProgress?: (phase: 'analyzing' | 'detailing' | 'creating', current: number, total: number, message: string) => void
): Promise<void>
```

**Parameters:**
- `outline`: 生成するプレゼンテーションのアウトライン
- `openAIService`: OpenAI API サービスインスタンス
- `options`: スライド生成オプション
- `onDetailProgress`: 進捗コールバック関数

**Example:**
```typescript
const service = new PowerPointService();
await service.generateSlidesWithDetailedProgress(
  outline,
  openAIService,
  { theme: 'light', fontSize: 'medium' },
  (phase, current, total, message) => {
    console.log(`${phase}: ${current}/${total} - ${message}`);
  }
);
```

##### generateSlidesWithTemplate
テンプレートベースのスライド生成を実行します。

```typescript
public async generateSlidesWithTemplate(
  userInput: string,
  outline: PresentationOutline,
  openAIService: OpenAIService,
  options: SlideGenerationOptions = {},
  onProgress?: (phase: string, current: number, total: number, message: string) => void
): Promise<void>
```

**Parameters:**
- `userInput`: 元のユーザー入力
- `outline`: プレゼンテーションアウトライン
- `openAIService`: OpenAI サービス
- `options`: 生成オプション
- `onProgress`: 進捗コールバック

##### registerCurrentPresentationAsTemplate
現在のプレゼンテーションをテンプレートとして登録します。

```typescript
public async registerCurrentPresentationAsTemplate(
  templateName: string,
  description: string,
  metadata: Partial<TemplateRegistrationRequest['metadata']>
): Promise<TemplateInfo>
```

**Parameters:**
- `templateName`: テンプレート名
- `description`: テンプレートの説明
- `metadata`: テンプレートメタデータ

**Returns:** 登録されたテンプレート情報

##### getTemplateRecommendations
ユーザー入力に基づいてテンプレート推奨を取得します。

```typescript
public async getTemplateRecommendations(
  userInput: string,
  preferences?: {
    categories?: string[];
    maxResults?: number;
  }
): Promise<TemplateRecommendation[]>
```

### OpenAIService

OpenAI APIとの統合を管理するサービスクラス。

#### Constructor
```typescript
class OpenAIService {
  constructor(settings: OpenAISettings)
}
```

**Parameters:**
- `settings.apiKey`: OpenAI APIキー
- `settings.model`: 使用するAIモデル（デフォルト: gpt-3.5-turbo）

#### Methods

##### generateStructuredOutline
ユーザー入力からプレゼンテーションアウトラインを生成します。

```typescript
public async generateStructuredOutline(userInput: string): Promise<PresentationOutline>
```

**Parameters:**
- `userInput`: ユーザーの要求

**Returns:** 構造化されたプレゼンテーションアウトライン

**Example:**
```typescript
const service = new OpenAIService({ apiKey: 'your-api-key' });
const outline = await service.generateStructuredOutline(
  "営業戦略についてのプレゼンテーションを作成してください"
);
```

##### regenerateOutline
既存のアウトラインを指示に基づいて再生成します。

```typescript
public async regenerateOutline(
  currentOutline: PresentationOutline,
  instruction: string
): Promise<PresentationOutline>
```

##### sendRequest
OpenAI APIへの直接リクエスト送信（低レベルAPI）。

```typescript
public async sendRequest(messages: any[]): Promise<string>
```

### SlideContentGenerator

個別スライドのコンテンツ詳細化を担当するサービス。

#### Constructor
```typescript
class SlideContentGenerator {
  constructor(openAIService: OpenAIService)
}
```

#### Methods

##### generateWithErrorHandling
エラーハンドリング付きでスライドコンテンツを生成します。

```typescript
public async generateWithErrorHandling(
  outline: PresentationOutline,
  options: SlideGenerationOptions,
  onProgress?: (current: number, total: number, slideName: string) => void,
  onError?: (slideIndex: number, error: Error) => void
): Promise<SlideContent[]>
```

**Example:**
```typescript
const generator = new SlideContentGenerator(openAIService);
const detailedSlides = await generator.generateWithErrorHandling(
  outline,
  options,
  (current, total, name) => console.log(`Progress: ${current}/${total} - ${name}`),
  (index, error) => console.error(`Slide ${index} failed:`, error)
);
```

##### validateSlideContent
スライドコンテンツの品質を検証します。

```typescript
public validateSlideContent(content: SlideContent): {
  isValid: boolean;
  warnings: string[];
  suggestions: string[];
}
```

## 🎨 Template System

### TemplateBasedGenerationService

テンプレートベースの生成機能を提供するサービス。

#### Methods

##### selectOptimalTemplate
ユーザー入力に最適なテンプレートを選択します。

```typescript
public async selectOptimalTemplate(
  userInput: string,
  criteria?: Partial<TemplateSelectionCriteria>
): Promise<TemplateRecommendation[]>
```

##### registerTemplate
新しいテンプレートを登録します。

```typescript
public async registerTemplate(
  request: TemplateRegistrationRequest
): Promise<TemplateInfo>
```

##### getTemplateLibrary
登録済みテンプレートライブラリを取得します。

```typescript
public getTemplateLibrary(): TemplateLibrary
```

### TemplateAdaptationService

テンプレート検出・適応機能を提供するサービス。

#### Methods

##### detectTemplate
現在のプレゼンテーションからテンプレート情報を検出します。

```typescript
public async detectTemplate(): Promise<TemplateInfo | null>
```

##### adaptOutlineToTemplate
アウトラインをテンプレートに適応させます。

```typescript
public async adaptOutlineToTemplate(
  outline: any,
  template: TemplateInfo
): Promise<AdaptedOutline>
```

## 🎭 Theme System

### ThemeService

テーマ管理とカラー適用を担当するサービス。

#### Methods

##### getCurrentThemeInfo
現在のテーマ情報を取得します。

```typescript
public async getCurrentThemeInfo(): Promise<ThemeInfo>
```

##### applyTheme
指定されたテーマを適用します。

```typescript
public async applyTheme(
  themeName: 'light' | 'dark' | 'colorful',
  options?: ThemeOptions
): Promise<void>
```

### ThemeApplier

テーマカラーとスタイルの適用を実行するクラス。

#### Methods

##### applyThemeColors
指定されたオブジェクトにテーマカラーを適用します。

```typescript
public applyThemeColors(
  shape: PowerPoint.Shape,
  theme: 'light' | 'dark' | 'colorful',
  elementType: 'title' | 'body' | 'accent'
): void
```

## 📊 Type Definitions

### Core Types

#### PresentationOutline
```typescript
interface PresentationOutline {
  title: string;
  estimatedDuration: number;
  slides: SlideOutline[];
}
```

#### SlideOutline
```typescript
interface SlideOutline {
  slideNumber: number;
  title: string;
  content: string[];
  slideType: 'title' | 'content' | 'conclusion';
}
```

#### SlideContent
```typescript
interface SlideContent {
  title: string;
  content: string[];
  slideType: 'title' | 'content' | 'conclusion';
  speakerNotes?: string;
}
```

#### SlideGenerationOptions
```typescript
interface SlideGenerationOptions {
  slideLayout?: 'content' | 'twoContent' | 'comparison' | 'blank';
  theme?: 'light' | 'dark' | 'colorful';
  fontSize?: 'small' | 'medium' | 'large';
  includeTransitions?: boolean;
  useThemeAwareGeneration?: boolean;
}
```

### Template Types

#### TemplateInfo
```typescript
interface TemplateInfo {
  id: string;
  name: string;
  description?: string;
  category: TemplateCategory;
  metadata: TemplateMetadata;
  designPatterns: DesignPattern[];
  structure: TemplateStructure;
  compatibility: TemplateCompatibility;
}
```

#### TemplateMetadata
```typescript
interface TemplateMetadata {
  presentationStyle: 'formal' | 'casual' | 'creative' | 'minimalist';
  targetAudience: 'executive' | 'technical' | 'general' | 'academic';
  slideCount?: number;
  colorSchemeType: 'corporate' | 'academic' | 'creative' | 'minimal' | 'custom';
  layoutComplexity: 'simple' | 'moderate' | 'complex';
  contentDensity: 'low' | 'medium' | 'high';
  purpose: 'pitch' | 'report' | 'training' | 'analysis' | 'overview';
  tags: string[];
  registeredAt: Date;
  usageCount: number;
  industry?: string;
}
```

#### TemplateRecommendation
```typescript
interface TemplateRecommendation {
  template: TemplateInfo;
  score: number;
  reasoning: string[];
  adaptations: any[];
}
```

## 🛠️ Utility Functions

### Content Processing

#### parseContentString
文字列を配列に変換する utility 関数。

```typescript
function parseContentString(content: string): string[]
```

#### adjustContentLength
コンテンツの長さを調整する utility 関数。

```typescript
function adjustContentLength(content: string[], maxLength: number = 200): string[]
```

### Error Handling

#### createFallbackContent
エラー時のフォールバックコンテンツを生成。

```typescript
function createFallbackContent(slide: SlideOutline): SlideContent
```

## 🔧 Configuration

### OpenAISettings
```typescript
interface OpenAISettings {
  apiKey: string;
  model?: string;
  temperature?: number;
  maxTokens?: number;
}
```

### SystemConfiguration
```typescript
interface SystemConfiguration {
  defaultTheme: 'light' | 'dark' | 'colorful';
  defaultFontSize: 'small' | 'medium' | 'large';
  maxSlideCount: number;
  apiTimeout: number;
  enableTemplateFeatures: boolean;
}
```

## 📈 Events and Callbacks

### Progress Callbacks

#### DetailedProgressCallback
AI詳細化の進捗を通知するコールバック。

```typescript
type DetailedProgressCallback = (
  phase: 'analyzing' | 'detailing' | 'creating',
  current: number,
  total: number,
  message: string
) => void;
```

#### TemplateProgressCallback
テンプレート処理の進捗を通知するコールバック。

```typescript
type TemplateProgressCallback = (
  phase: string,
  current: number,
  total: number,
  message: string
) => void;
```

### Error Callbacks

#### ErrorCallback
エラー発生時の通知コールバック。

```typescript
type ErrorCallback = (slideIndex: number, error: Error) => void;
```

## 🔍 Usage Examples

### Complete Workflow Example

```typescript
// 1. サービスの初期化
const openAIService = new OpenAIService({
  apiKey: 'your-api-key',
  model: 'gpt-3.5-turbo'
});

const powerPointService = new PowerPointService();

// 2. アウトライン生成
const outline = await openAIService.generateStructuredOutline(
  "デジタルマーケティング戦略について"
);

// 3. AI詳細化でスライド生成
await powerPointService.generateSlidesWithDetailedProgress(
  outline,
  openAIService,
  {
    theme: 'light',
    fontSize: 'medium',
    slideLayout: 'content'
  },
  (phase, current, total, message) => {
    console.log(`Phase: ${phase}, Progress: ${current}/${total}, Message: ${message}`);
  }
);
```

### Template-based Generation Example

```typescript
// 1. テンプレート推奨の取得
const recommendations = await powerPointService.getTemplateRecommendations(
  "営業提案プレゼンテーション"
);

// 2. テンプレートベース生成
if (recommendations.length > 0) {
  const selectedTemplate = recommendations[0].template;
  
  await powerPointService.generateSlidesWithTemplate(
    "営業提案プレゼンテーション",
    outline,
    openAIService,
    { theme: 'light' },
    (phase, current, total, message) => {
      console.log(`Template phase: ${phase}, Progress: ${current}/${total}`);
    }
  );
}
```

### Error Handling Example

```typescript
try {
  await powerPointService.generateSlidesWithDetailedProgress(
    outline,
    openAIService,
    options,
    progressCallback
  );
} catch (error) {
  if (error instanceof NetworkError) {
    console.error('Network error:', error.message);
    // ネットワークエラーの処理
  } else if (error instanceof APIError) {
    console.error('API error:', error.message);
    // APIエラーの処理
  } else {
    console.error('Unexpected error:', error);
    // その他のエラー処理
  }
}
```

このAPIリファレンスを参考に、PowerPoint Conciergeの機能を効果的に活用してください。詳細な実装例は [開発ガイド](./development-guide.md) をご参照ください。