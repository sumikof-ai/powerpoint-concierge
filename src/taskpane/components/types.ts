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
    type: 'user' | 'assistant';
  }
  
  // プレゼンテーション関連の型定義
  export interface PresentationOutline {
    title: string;
    slides: SlideOutline[];
  }
  
  export interface SlideOutline {
    id: string;
    title: string;
    content: string[];
    notes?: string;
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
      role: 'system' | 'user' | 'assistant';
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