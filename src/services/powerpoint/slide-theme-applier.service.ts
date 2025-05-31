// src/services/powerpoint/slide-theme-applier.service.ts - テーマ適用サービス
/* global PowerPoint */

import { FontSizes, ThemeType, FontSizeType, ColorType } from './types';

/**
 * スライドのテーマとスタイルを適用するサービス
 */
export class SlideThemeApplier {
  
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
      switch (theme) {
        case 'dark':
          this.applyDarkTheme(shape, type);
          break;
        case 'colorful':
          this.applyColorfulTheme(shape, type);
          break;
        case 'light':
        default:
          this.applyLightTheme(shape, type);
      }
    } catch (error) {
      console.warn('テーマ色の適用に失敗しました:', error);
      this.applyFallbackColors(shape);
    }
  }

  /**
   * ライトテーマを適用
   */
  private applyLightTheme(shape: PowerPoint.Shape, type?: ColorType): void {
    shape.fill.setSolidColor('#FFFFFF');
    shape.textFrame.textRange.font.color = '#333333';
    
    if (type === 'accent' && shape.lineFormat) {
      shape.fill.setSolidColor('#F8F9FA');
      shape.lineFormat.color = '#DEE2E6';
      shape.lineFormat.weight = 1;
    }
  }

  /**
   * ダークテーマを適用
   */
  private applyDarkTheme(shape: PowerPoint.Shape, type?: ColorType): void {
    shape.fill.setSolidColor('#2D2D30');
    shape.textFrame.textRange.font.color = '#FFFFFF';
    
    if (type === 'accent') {
      shape.fill.setSolidColor('#007ACC');
    }
  }

  /**
   * カラフルテーマを適用
   */
  private applyColorfulTheme(shape: PowerPoint.Shape, type?: ColorType): void {
    switch (type) {
      case 'title':
        shape.fill.setSolidColor('#FF6B6B');
        shape.textFrame.textRange.font.color = '#FFFFFF';
        break;
      case 'heading':
        shape.fill.setSolidColor('#4ECDC4');
        shape.textFrame.textRange.font.color = '#FFFFFF';
        break;
      case 'accent':
        shape.fill.setSolidColor('#45B7D1');
        shape.textFrame.textRange.font.color = '#FFFFFF';
        break;
      default:
        shape.fill.setSolidColor('#FFFFFF');
        shape.textFrame.textRange.font.color = '#2C3E50';
    }
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
}