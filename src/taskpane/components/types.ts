// OpenAI API関連の型定義
export interface OpenAISettings {
  apiKey: string;
  baseUrl: string;
  model: string;
  temperature: number;
  maxTokens: number;
}

// チャットメッセージの型定義
export interface ChatMessage {
  id: string;
  content: string;
  timestamp: Date;
  type: "user" | "assistant";
}

// プレゼンテーション関連の型定義（OutlineEditorからインポートして使用）
export interface PresentationOutline {
  title: string;
  slides: SlideOutline[];
  estimatedDuration: number;
}

export interface SlideOutline {
  slideNumber: number;
  title: string;
  content: string[];
  slideType: "title" | "content" | "conclusion";
  speakerNotes?: string;
}

// PowerPoint操作関連の型定義
export interface SlideInfo {
  id: string;
  title: string;
  content: string;
  index: number;
}

// 段階的ワークフロー管理
export type WorkflowStep = "chat" | "outline" | "generating" | "completed";

export interface WorkflowState {
  currentStep: WorkflowStep;
  outline: PresentationOutline | null;
  generationProgress: string;
  isLoading: boolean;
  error: string;
}

// OpenAI APIレスポンスの型定義
export interface OpenAIResponse {
  id: string;
  object: string;
  created: number;
  model: string;
  choices: OpenAIChoice[];
  usage: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
}

export interface OpenAIChoice {
  index: number;
  message: {
    role: string;
    content: string;
  };
  finish_reason: string;
}

// APIリクエストの型定義
export interface OpenAIRequest {
  model: string;
  messages: {
    role: "system" | "user" | "assistant";
    content: string;
  }[];
  temperature?: number;
  max_tokens?: number;
}

// エラーハンドリングの型定義
export interface APIError {
  message: string;
  code?: string;
  status?: number;
}

// スライド生成オプション
export interface SlideGenerationOptions {
  includeTransitions?: boolean;
  useTemplate?: string;
  slideLayout?: "title" | "content" | "comparison" | "blank";
}

// プレゼンテーション設定
export interface PresentationSettings {
  theme: "light" | "dark" | "colorful";
  fontSize: "small" | "medium" | "large";
  includeSlideNumbers: boolean;
  includeNotes: boolean;
}
