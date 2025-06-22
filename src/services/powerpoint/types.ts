// src/services/powerpoint/types.ts - 更新版（実用的なテーマ対応）
/* global PowerPoint */

export interface SlideInfo {
  id: string;
  title: string;
  content: string;
  index: number;
}

export interface SlideGenerationOptions {
  includeTransitions?: boolean;
  useTemplate?: string;
  slideLayout?: "title" | "content" | "twoContent" | "comparison" | "blank";
  theme?: "light" | "dark" | "colorful";
  fontSize?: "small" | "medium" | "large";
  useThemeAwareGeneration?: boolean;
}

export interface SlideContent {
  title: string;
  content: string[];
  slideType: "title" | "content" | "conclusion";
  speakerNotes?: string;
}

export interface BulkSlideData {
  slides: SlideContent[];
  options?: SlideGenerationOptions;
}

export interface FontSizes {
  title: number;
  heading: number;
  subtitle: number;
  body: number;
  accent: number;
}

export interface PresentationStats {
  slideCount: number;
  estimatedDuration: number;
  wordCount: number;
}

export interface ThemeColors {
  backgroundColor: string;
  textColor: string;
  accentColor: string;
  titleColor: string;
  borderColor: string;
}

export interface PresentationInfo {
  slideCount: number;
  title: string;
  existingStyle: {
    commonFontSize: number;
    commonFontColor: string;
    backgroundDetected: boolean;
  } | null;
}

export type SlideLayoutType = "title" | "content" | "twoContent" | "comparison" | "blank";
export type ThemeType = "light" | "dark" | "colorful";
export type FontSizeType = "small" | "medium" | "large";
export type ColorType = "title" | "heading" | "subtitle" | "body" | "accent";

export interface PowerPointContext {
  context: PowerPoint.RequestContext;
  slide: PowerPoint.Slide;
}

export interface ShapeOptions {
  left: number;
  top: number;
  width: number;
  height: number;
}

// 実用的なテーマプリセット用の型
export interface ThemePreset {
  name: string;
  colors: ThemeColors;
  description: string;
}

// スライド作成結果の型
export interface SlideCreationResult {
  success: boolean;
  slideIndex: number;
  title: string;
  error?: string;
}

// 一括スライド生成の進捗情報
export interface GenerationProgress {
  current: number;
  total: number;
  currentSlideName: string;
  percentage: number;
  timeElapsed: number;
  estimatedTimeRemaining?: number;
}

// エラーハンドリング用の型
export interface PowerPointError {
  code: string;
  message: string;
  details?: any;
  recoverable: boolean;
}

// デバッグ・テスト用の型
export interface ThemeTestResult {
  themeName: string;
  slidesCreated: number;
  errors: string[];
  executionTime: number;
}
