// src/services/powerpoint/core/ThemeApplier.ts - テーマ適用サービス
/* global PowerPoint */

import { FontSizes, ThemeType, FontSizeType, ColorType } from '../types';

/**
 * スライドのテーマとスタイルを適用するサービス
 */
export class ThemeApplier {
  
  // 利用可能なテーマ設定
  private themePresets = {
    light: {
      backgroundColor: '#FFFFFF',
      textColor: '#000000',
      accentColor: '#0078D4',
      titleColor: '#323130',
      borderColor: '#D1D1D1'
    },
    dark: {
      backgroundColor: '#1F1F1F',
      textColor: '#FFFFFF',
      accentColor: '#0078D4',
      titleColor: '#FFFFFF',
      borderColor: '#404040'
    },
    colorful: {
      backgroundColor: '#FFFFFF',
      textColor: '#323130',
      accentColor: '#FF6B35',
      titleColor: '#2D3748',
      borderColor: '#E2E8F0'
    }
  };

  /**
   * フォントサイズを取得
   */
  public getFontSize(size?: FontSizeType): FontSizes {
    switch (size) {
      case 'small':
        return { title: 32, heading: 20, subtitle: 16, body: 12, accent: 14 };
      case 'large':
        return { title: 44, heading: 28, subtitle: 22, body: 16, accent: 18 };
      case 'medium':
      default:
        return { title: 38, heading: 24, subtitle: 18, body: 14, accent: 16 };
    }
  }

  /**
   * テーマに基づいて色を適用
   */
  public applyThemeColors(
    shape: PowerPoint.Shape, 
    theme?: ThemeType, 
    type?: ColorType
  ): void {
    try {
      const themeColors = this.themePresets[theme || 'light'];
      
      switch (type) {
        case 'title':
          this.applyTitleStyle(shape, themeColors);
          break;
        case 'heading':
          this.applyHeadingStyle(shape, themeColors);
          break;
        case 'subtitle':
          this.applySubtitleStyle(shape, themeColors);
          break;
        case 'body':
          this.applyBodyStyle(shape, themeColors);
          break;
        case 'accent':
          this.applyAccentStyle(shape, themeColors);
          break;
        default:
          this.applyDefaultStyle(shape, themeColors);
      }
    } catch (error) {
      console.warn('テーマ色の適用に失敗しました:', error);
      this.applyFallbackColors(shape);
    }
  }

  /**
   * タイトルスタイルを適用
   */
  private applyTitleStyle(shape: PowerPoint.Shape, themeColors: any): void {
    shape.fill.setSolidColor(themeColors.backgroundColor);
    shape.textFrame.textRange.font.color = themeColors.titleColor;
    
    // タイトル用の装飾
    if (shape.lineFormat) {
      shape.lineFormat.color = themeColors.accentColor;
      shape.lineFormat.weight = 2;
    }
  }

  /**
   * 見出しスタイルを適用
   */
  private applyHeadingStyle(shape: PowerPoint.Shape, themeColors: any): void {
    shape.fill.setSolidColor(themeColors.backgroundColor);
    shape.textFrame.textRange.font.color = themeColors.titleColor;
    
    if (shape.lineFormat) {
      shape.lineFormat.color = themeColors.borderColor;
      shape.lineFormat.weight = 1;
    }
  }

  /**
   * サブタイトルスタイルを適用
   */
  private applySubtitleStyle(shape: PowerPoint.Shape, themeColors: any): void {
    shape.fill.setSolidColor(themeColors.backgroundColor);
    shape.textFrame.textRange.font.color = themeColors.textColor;
  }

  /**
   * 本文スタイルを適用
   */
  private applyBodyStyle(shape: PowerPoint.Shape, themeColors: any): void {
    shape.fill.setSolidColor(themeColors.backgroundColor);
    shape.textFrame.textRange.font.color = themeColors.textColor;
    
    if (shape.lineFormat) {
      shape.lineFormat.color = themeColors.borderColor;
      shape.lineFormat.weight = 1;
    }
  }

  /**
   * アクセントスタイルを適用
   */
  private applyAccentStyle(shape: PowerPoint.Shape, themeColors: any): void {
    shape.fill.setSolidColor(themeColors.accentColor);
    shape.textFrame.textRange.font.color = themeColors.backgroundColor;
    
    if (shape.lineFormat) {
      shape.lineFormat.color = themeColors.accentColor;
      shape.lineFormat.weight = 2;
    }
  }

  /**
   * デフォルトスタイルを適用
   */
  private applyDefaultStyle(shape: PowerPoint.Shape, themeColors: any): void {
    shape.fill.setSolidColor(themeColors.backgroundColor);
    shape.textFrame.textRange.font.color = themeColors.textColor;
  }

  /**
   * フォールバック色を適用
   */
  private applyFallbackColors(shape: PowerPoint.Shape): void {
    try {
      shape.fill.setSolidColor('#FFFFFF');
      shape.textFrame.textRange.font.color = '#000000';
    } catch (fallbackError) {
      console.warn('フォールバック色の適用も失敗しました:', fallbackError);
    }
  }

  /**
   * アクセントカラーを取得
   */
  public getAccentColor(theme?: ThemeType): string {
    const themeColors = this.themePresets[theme || 'light'];
    return themeColors.accentColor;
  }

  /**
   * 背景色を取得
   */
  public getBackgroundColor(theme?: ThemeType): string {
    const themeColors = this.themePresets[theme || 'light'];
    return themeColors.backgroundColor;
  }

  /**
   * テキスト色を取得
   */
  public getTextColor(theme?: ThemeType): string {
    const themeColors = this.themePresets[theme || 'light'];
    return themeColors.textColor;
  }

  /**
   * ボーダー色を取得
   */
  public getBorderColor(theme?: ThemeType): string {
    const themeColors = this.themePresets[theme || 'light'];
    return themeColors.borderColor;
  }

  /**
   * タイトル色を取得
   */
  public getTitleColor(theme?: ThemeType): string {
    const themeColors = this.themePresets[theme || 'light'];
    return themeColors.titleColor;
  }

  /**
   * テーマプリセットを取得
   */
  public getThemePreset(theme?: ThemeType) {
    return this.themePresets[theme || 'light'];
  }

  /**
   * カスタムテーマを適用
   */
  public applyCustomTheme(
    shape: PowerPoint.Shape,
    customColors: {
      backgroundColor?: string;
      textColor?: string;
      accentColor?: string;
    }
  ): void {
    try {
      if (customColors.backgroundColor) {
        shape.fill.setSolidColor(customColors.backgroundColor);
      }
      if (customColors.textColor) {
        shape.textFrame.textRange.font.color = customColors.textColor;
      }
      if (customColors.accentColor && shape.lineFormat) {
        shape.lineFormat.color = customColors.accentColor;
      }
    } catch (error) {
      console.warn('カスタムテーマの適用に失敗しました:', error);
      this.applyFallbackColors(shape);
    }
  }

  /**
   * テーマの互換性をチェック
   */
  public validateThemeCompatibility(theme: ThemeType): {
    isSupported: boolean;
    warnings: string[];
  } {
    const warnings: string[] = [];
    let isSupported = true;

    // PowerPoint.js APIでサポートされていない機能のチェック
    if (theme === 'dark') {
      warnings.push('ダークテーマは一部の環境で表示が異なる場合があります');
    }

    if (theme === 'colorful') {
      warnings.push('カラフルテーマは印刷時の色味が異なる場合があります');
    }

    return { isSupported, warnings };
  }

  /**
   * グラデーション効果を適用（実験的）
   */
  public applyGradientEffect(
    shape: PowerPoint.Shape,
    startColor: string,
    endColor: string
  ): void {
    try {
      // PowerPoint.js APIでのグラデーション適用は限定的
      // 将来の拡張に備えた実装
      console.log(`グラデーション効果を適用: ${startColor} → ${endColor}`);
      
      // フォールバック: 開始色を背景色として使用
      shape.fill.setSolidColor(startColor);
    } catch (error) {
      console.warn('グラデーション効果の適用に失敗しました:', error);
      shape.fill.setSolidColor(startColor);
    }
  }

  /**
   * 影効果を適用（実験的）
   */
  public applyShadowEffect(
    shape: PowerPoint.Shape,
    shadowColor: string = '#000000',
    opacity: number = 0.3
  ): void {
    try {
      // PowerPoint.js APIでの影効果は限定的
      console.log(`影効果を適用: ${shadowColor}, 透明度: ${opacity}`);
      
      // 現在のAPIでは直接的な影効果の適用は制限される
      // 代替として、ボーダーで擬似的な効果を作成
      if (shape.lineFormat) {
        shape.lineFormat.color = shadowColor;
        shape.lineFormat.weight = 2;
      }
    } catch (error) {
      console.warn('影効果の適用に失敗しました:', error);
    }
  }

  /**
   * アニメーション効果のヒント（将来拡張用）
   */
  public getAnimationSuggestions(slideType: string): string[] {
    const suggestions: string[] = [];
    
    switch (slideType) {
      case 'title':
        suggestions.push('フェードイン効果でタイトルを印象的に');
        break;
      case 'content':
        suggestions.push('箇条書きを順次表示');
        break;
      case 'conclusion':
        suggestions.push('まとめ項目を強調アニメーション');
        break;
    }
    
    return suggestions;
  }

  /**
   * 色のコントラスト比をチェック
   */
  public checkColorContrast(
    foregroundColor: string,
    backgroundColor: string
  ): { ratio: number; isAccessible: boolean } {
    // 簡易的なコントラスト比計算
    // 実際の実装では、より詳細な色空間変換が必要
    
    const getLuminance = (color: string): number => {
      // 16進数カラーをRGBに変換して輝度を計算
      const hex = color.replace('#', '');
      const r = parseInt(hex.substr(0, 2), 16) / 255;
      const g = parseInt(hex.substr(2, 2), 16) / 255;
      const b = parseInt(hex.substr(4, 2), 16) / 255;
      
      return 0.299 * r + 0.587 * g + 0.114 * b;
    };
    
    const fgLuminance = getLuminance(foregroundColor);
    const bgLuminance = getLuminance(backgroundColor);
    
    const ratio = Math.abs(fgLuminance - bgLuminance);
    const isAccessible = ratio > 0.5; // 簡易的な閾値
    
    return { ratio, isAccessible };
  }
}