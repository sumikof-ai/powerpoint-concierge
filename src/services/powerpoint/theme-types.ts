// src/services/powerpoint/theme-types.ts - テーマ関連の型定義

/**
 * テーマ情報の全体構造
 */
export interface ThemeInfo {
  name: string;
  colorScheme: ColorScheme;
  fontScheme: FontScheme;
  availableLayouts: LayoutInfo[];
}

/**
 * カラースキーム定義
 */
export interface ColorScheme {
  // 主要なアクセントカラー
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;

  // 背景色
  background1: string;
  background2: string;

  // テキスト色
  text1: string;
  text2: string;

  // ハイパーリンク色
  hyperlink: string;
  followedHyperlink: string;
}

/**
 * フォントスキーム定義
 */
export interface FontScheme {
  majorFont: FontSet; // 見出し用フォント
  minorFont: FontSet; // 本文用フォント
}

/**
 * フォントセット（言語別）
 */
export interface FontSet {
  latin: string; // 欧文フォント
  eastAsian: string; // 東アジア言語フォント
  complexScript: string; // 複雑な文字体系用フォント
}

/**
 * スライドレイアウト情報
 */
export interface LayoutInfo {
  name: string;
  type: "title" | "content" | "comparison" | "blank" | "section" | "custom";
  placeholders: PlaceholderInfo[];
  isCustom?: boolean;
}

/**
 * プレースホルダー情報
 */
export interface PlaceholderInfo {
  type: "title" | "subtitle" | "content" | "footer" | "slideNumber" | "date" | "picture" | "chart";
  position: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  textFormat: TextFormatInfo;
  isRequired?: boolean;
}

/**
 * テキストフォーマット情報
 */
export interface TextFormatInfo {
  defaultFontSize: number;
  defaultFontBold: boolean;
  defaultFontItalic?: boolean;
  defaultAlignment: "left" | "center" | "right" | "justify";
  defaultColor?: string;
  defaultLineSpacing?: number;
  bulletStyle?: BulletStyle;
}

/**
 * 箇条書きスタイル
 */
export interface BulletStyle {
  type: "none" | "bullet" | "number" | "custom";
  symbol?: string;
  indentLevel?: number;
  color?: string;
}

/**
 * テーマ適用オプション
 */
export interface ThemeApplicationOptions {
  preserveCustomFormatting?: boolean;
  applyToAllSlides?: boolean;
  includeSlideLayouts?: boolean;
  includeSlideMaster?: boolean;
}

/**
 * スライドデザイン情報（拡張情報）
 */
export interface SlideDesignInfo {
  backgroundStyle?: BackgroundStyle;
  transitionEffect?: TransitionEffect;
  animationScheme?: AnimationScheme;
}

/**
 * 背景スタイル
 */
export interface BackgroundStyle {
  type: "solid" | "gradient" | "pattern" | "picture";
  color?: string;
  gradientStops?: GradientStop[];
  pictureUrl?: string;
  transparency?: number;
}

/**
 * グラデーション停止点
 */
export interface GradientStop {
  color: string;
  position: number; // 0-1
}

/**
 * トランジション効果
 */
export interface TransitionEffect {
  type: "none" | "fade" | "push" | "wipe" | "split" | "reveal" | "random";
  duration: number; // 秒単位
  advanceOnClick: boolean;
  advanceAfterTime?: number; // 秒単位
}

/**
 * アニメーションスキーム
 */
export interface AnimationScheme {
  entranceEffect?: "none" | "appear" | "fade" | "fly" | "float" | "split" | "wipe";
  emphasisEffect?: "none" | "pulse" | "spin" | "grow" | "teeter";
  exitEffect?: "none" | "disappear" | "fade" | "fly" | "float" | "split" | "wipe";
  motionPath?: "none" | "lines" | "arcs" | "turns" | "shapes" | "loops";
}
