// src/services/powerpoint/theme/ThemeAnalyzer.ts - テーマ分析サービス
/* global PowerPoint */

import { 
    ColorScheme, 
    FontScheme, 
    FontSet, 
    PlaceholderInfo, 
    TextFormatInfo 
  } from '../theme-types';
  
  /**
   * PowerPointのテーマ分析を専門に行うサービス
   */
  export class ThemeAnalyzer {
  
    /**
     * カラースキームを分析
     */
    public async analyzeColorScheme(context: PowerPoint.RequestContext): Promise<ColorScheme> {
      try {
        // 実際のスライドから色情報を抽出する試み
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
  
        if (slides.items.length > 0) {
          const detectedColors = await this.extractColorsFromSlides(context, slides.items.slice(0, 3));
          if (detectedColors) {
            return detectedColors;
          }
        }
      } catch (error) {
        console.warn("カラースキームの分析に失敗しました。デフォルト値を使用します。", error);
      }
  
      // デフォルトのOfficeテーマカラー
      return this.getDefaultColorScheme();
    }
  
    /**
     * フォントスキームを分析
     */
    public async analyzeFontScheme(context: PowerPoint.RequestContext): Promise<FontScheme> {
      try {
        const detectedFonts = await this.extractFontsFromPresentation(context);
        if (detectedFonts) {
          return detectedFonts;
        }
      } catch (error) {
        console.warn("フォントスキームの分析に失敗しました。デフォルト値を使用します。", error);
      }
  
      // デフォルトのフォントスキーム
      return this.getDefaultFontScheme();
    }
  
    /**
     * スライドからプレースホルダーを検出
     */
    public async detectPlaceholders(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide
    ): Promise<PlaceholderInfo[]> {
      const placeholders: PlaceholderInfo[] = [];
  
      try {
        slide.shapes.load("items");
        await context.sync();
  
        for (const shape of slide.shapes.items) {
          try {
            shape.load(["type", "left", "top", "width", "height", "textFrame"]);
            await context.sync();
            
            const isPlaceholder = await this.isShapePlaceholder(context, shape);
            
            if (isPlaceholder) {
              const placeholderInfo = await this.analyzePlaceholder(context, shape);
              if (placeholderInfo) {
                placeholders.push(placeholderInfo);
              }
            }
          } catch (shapeError) {
            console.warn('個別シェイプの分析中にエラー:', shapeError);
          }
        }
      } catch (error) {
        console.warn('プレースホルダーの検出に失敗しました:', error);
      }
  
      return placeholders;
    }
  
    /**
     * スライドから色情報を抽出
     */
    private async extractColorsFromSlides(
      context: PowerPoint.RequestContext,
      slides: PowerPoint.Slide[]
    ): Promise<ColorScheme | null> {
      const detectedColors = {
        backgrounds: new Set<string>(),
        texts: new Set<string>(),
        accents: new Set<string>()
      };
  
      try {
        for (const slide of slides) {
          slide.shapes.load("items");
          await context.sync();
  
          for (const shape of slide.shapes.items.slice(0, 5)) { // 最初の5個のシェイプのみ分析
            try {
              await this.analyzeShapeColors(context, shape, detectedColors);
            } catch (error) {
              console.warn('シェイプ色分析エラー:', error);
            }
          }
        }
  
        return this.buildColorSchemeFromDetected(detectedColors);
      } catch (error) {
        console.warn('色抽出処理でエラー:', error);
        return null;
      }
    }
  
    /**
     * シェイプの色を分析
     */
    private async analyzeShapeColors(
      context: PowerPoint.RequestContext,
      shape: PowerPoint.Shape,
      detectedColors: any
    ): Promise<void> {
      try {
        // テキストフレームがある場合のフォント色分析
        if (shape.textFrame) {
          shape.textFrame.load("textRange");
          await context.sync();
          
          if (shape.textFrame.textRange.font) {
            shape.textFrame.textRange.font.load("color");
            await context.sync();
            
            const fontColor = shape.textFrame.textRange.font.color;
            if (fontColor && this.isValidColor(fontColor)) {
              detectedColors.texts.add(fontColor);
            }
          }
        }
  
        // 塗りつぶし色の分析は PowerPoint.js API の制限により困難
        // 将来のAPI拡張に備えた構造のみ準備
      } catch (error) {
        console.warn('シェイプ色分析で個別エラー:', error);
      }
    }
  
    /**
     * プレゼンテーションからフォント情報を抽出
     */
    private async extractFontsFromPresentation(
      context: PowerPoint.RequestContext
    ): Promise<FontScheme | null> {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
  
        const detectedFonts = {
          titles: new Set<string>(),
          bodies: new Set<string>()
        };
  
        // 最初の3枚のスライドからフォント情報を収集
        const analyzableSlides = slides.items.slice(0, 3);
        
        for (const slide of analyzableSlides) {
          await this.extractFontsFromSlide(context, slide, detectedFonts);
        }
  
        return this.buildFontSchemeFromDetected(detectedFonts);
      } catch (error) {
        console.warn('フォント抽出処理でエラー:', error);
        return null;
      }
    }
  
    /**
     * 個別スライドからフォント情報を抽出
     */
    private async extractFontsFromSlide(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      detectedFonts: any
    ): Promise<void> {
      try {
        slide.shapes.load("items");
        await context.sync();
  
        for (const shape of slide.shapes.items.slice(0, 5)) {
          await this.analyzeShapeFont(context, shape, detectedFonts);
        }
      } catch (error) {
        console.warn('スライドフォント分析エラー:', error);
      }
    }
  
    /**
     * シェイプのフォント情報を分析
     */
    private async analyzeShapeFont(
      context: PowerPoint.RequestContext,
      shape: PowerPoint.Shape,
      detectedFonts: any
    ): Promise<void> {
      try {
        if (shape.textFrame) {
          shape.textFrame.load("textRange");
          await context.sync();
          
          if (shape.textFrame.textRange.font) {
            shape.textFrame.textRange.font.load(["name", "size"]);
            await context.sync();
            
            const fontName = shape.textFrame.textRange.font.name;
            const fontSize = shape.textFrame.textRange.font.size;
            
            if (fontName) {
              // フォントサイズに基づいてタイトル用か本文用かを判定
              if (fontSize > 20) {
                detectedFonts.titles.add(fontName);
              } else {
                detectedFonts.bodies.add(fontName);
              }
            }
          }
        }
      } catch (error) {
        console.warn('シェイプフォント分析で個別エラー:', error);
      }
    }
  
    /**
     * シェイプがプレースホルダーかどうかを判定
     */
    private async isShapePlaceholder(
      context: PowerPoint.RequestContext,
      shape: PowerPoint.Shape
    ): Promise<boolean> {
      try {
        // PowerPoint.js APIでの直接的なプレースホルダー判定は限定的
        // 位置とサイズの特徴から推測
        const hasTextFrame = shape.textFrame !== null;
        const isInTypicalPosition = this.isInTypicalPlaceholderPosition(shape);
        const hasTypicalSize = this.hasTypicalPlaceholderSize(shape);
        
        return hasTextFrame && (isInTypicalPosition || hasTypicalSize);
      } catch (error) {
        console.warn('プレースホルダー判定エラー:', error);
        console.warn(`context: ${context}`)
        return false;
      }
    }
  
    /**
     * 典型的なプレースホルダー位置かどうかを判定
     */
    private isInTypicalPlaceholderPosition(shape: PowerPoint.Shape): boolean {
      const top = shape.top;
      const left = shape.left;
      
      // タイトル位置（上部中央）
      const isTitlePosition = top < 100 && left > 50 && left < 200;
      // コンテンツ位置（中央部）
      const isContentPosition = top > 100 && top < 400;
      
      return isTitlePosition || isContentPosition;
    }
  
    /**
     * 典型的なプレースホルダーサイズかどうかを判定
     */
    private hasTypicalPlaceholderSize(shape: PowerPoint.Shape): boolean {
      const width = shape.width;
      const height = shape.height;
      
      // 一般的なプレースホルダーサイズの範囲
      const minWidth = 200;
      const maxWidth = 800;
      const minHeight = 40;
      const maxHeight = 500;
      
      return width >= minWidth && width <= maxWidth && 
             height >= minHeight && height <= maxHeight;
    }
  
    /**
     * プレースホルダーを詳細分析
     */
    private async analyzePlaceholder(
      context: PowerPoint.RequestContext,
      shape: PowerPoint.Shape
    ): Promise<PlaceholderInfo | null> {
      try {
        const placeholderType = this.determinePlaceholderType(shape);
        const textFormat = await this.analyzeTextFormat(context, shape);
        
        return {
          type: placeholderType,
          position: {
            x: shape.left,
            y: shape.top,
            width: shape.width,
            height: shape.height
          },
          textFormat
        };
      } catch (error) {
        console.warn('プレースホルダー分析エラー:', error);
        return null;
      }
    }
  
    /**
     * プレースホルダータイプを決定
     */
    private determinePlaceholderType(shape: PowerPoint.Shape): PlaceholderInfo['type'] {
      const top = shape.top;
      const height = shape.height;
      
      if (top < 100 && height > 50) {
        return 'title';
      } else if (top > 300 && height < 150) {
        return 'subtitle';
      } else {
        return 'content';
      }
    }
  
    /**
     * テキストフォーマットを分析
     */
    private async analyzeTextFormat(
      context: PowerPoint.RequestContext,
      shape: PowerPoint.Shape
    ): Promise<TextFormatInfo> {
      const defaultFormat: TextFormatInfo = {
        defaultFontSize: 18,
        defaultFontBold: false,
        defaultAlignment: 'left'
      };
  
      try {
        if (shape.textFrame) {
          shape.textFrame.load("textRange");
          await context.sync();
          
          if (shape.textFrame.textRange.font) {
            shape.textFrame.textRange.font.load(["size", "bold"]);
            await context.sync();
            
            return {
              defaultFontSize: shape.textFrame.textRange.font.size || 18,
              defaultFontBold: shape.textFrame.textRange.font.bold || false,
              defaultAlignment: 'left' // PowerPoint.js APIでの配置取得は限定的
            };
          }
        }
      } catch (error) {
        console.warn('テキストフォーマット分析エラー:', error);
      }
  
      return defaultFormat;
    }
  
    /**
     * 検出された色情報からColorSchemeを構築
     */
    private buildColorSchemeFromDetected(detectedColors: any): ColorScheme | null {
      try {
        const textColors = Array.from(detectedColors.texts) as string[];
        
        if (textColors.length > 0) {
          // 簡易的な色スキーム構築
          return {
            accent1: '#5B9BD5',
            accent2: '#ED7D31',
            accent3: '#A5A5A5',
            accent4: '#FFC000',
            accent5: '#4472C4',
            accent6: '#70AD47',
            background1: '#FFFFFF',
            background2: '#F2F2F2',
            text1: (textColors[0] as string) || '#000000',
            text2: (textColors[1] as string) || '#404040',
            hyperlink: '#0563C1',
            followedHyperlink: '#954F72'
          };
        }
      } catch (error) {
        console.warn('色スキーム構築エラー:', error);
      }
      
      return null;
    }
  
    /**
     * 検出されたフォント情報からFontSchemeを構築
     */
    private buildFontSchemeFromDetected(detectedFonts: any): FontScheme | null {
      try {
        const titleFonts = Array.from(detectedFonts.titles);
        const bodyFonts = Array.from(detectedFonts.bodies);
        
        if (titleFonts.length > 0 || bodyFonts.length > 0) {
          return {
            majorFont: this.createFontSet(titleFonts[0] as string || 'Calibri Light'),
            minorFont: this.createFontSet(bodyFonts[0] as string || 'Calibri')
          };
        }
      } catch (error) {
        console.warn('フォントスキーム構築エラー:', error);
      }
      
      return null;
    }
  
    /**
     * FontSetを作成
     */
    private createFontSet(primaryFont: string): FontSet {
      return {
        latin: primaryFont,
        eastAsian: 'MS Gothic',
        complexScript: 'Arial'
      };
    }
  
    /**
     * デフォルトのカラースキームを取得
     */
    private getDefaultColorScheme(): ColorScheme {
      return {
        accent1: '#5B9BD5',
        accent2: '#ED7D31',
        accent3: '#A5A5A5',
        accent4: '#FFC000',
        accent5: '#4472C4',
        accent6: '#70AD47',
        background1: '#FFFFFF',
        background2: '#F2F2F2',
        text1: '#000000',
        text2: '#404040',
        hyperlink: '#0563C1',
        followedHyperlink: '#954F72'
      };
    }
  
    /**
     * デフォルトのフォントスキームを取得
     */
    private getDefaultFontScheme(): FontScheme {
      return {
        majorFont: {
          latin: 'Calibri Light',
          eastAsian: 'MS Gothic',
          complexScript: 'Arial'
        },
        minorFont: {
          latin: 'Calibri',
          eastAsian: 'MS Gothic',
          complexScript: 'Arial'
        }
      };
    }
  
    /**
     * 有効な色かどうかを判定
     */
    private isValidColor(color: string): boolean {
      if (!color || typeof color !== 'string') {
        return false;
      }
      
      // 16進数カラーコードの基本的な検証
      const hexColorRegex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
      return hexColorRegex.test(color);
    }
  
    /**
     * 色の明度を計算
     */
    public calculateLuminance(color: string): number {
      try {
        const hex = color.replace('#', '');
        const r = parseInt(hex.substr(0, 2), 16) / 255;
        const g = parseInt(hex.substr(2, 2), 16) / 255;
        const b = parseInt(hex.substr(4, 2), 16) / 255;
        
        // sRGB輝度の計算
        const rsRGB = r <= 0.03928 ? r / 12.92 : Math.pow((r + 0.055) / 1.055, 2.4);
        const gsRGB = g <= 0.03928 ? g / 12.92 : Math.pow((g + 0.055) / 1.055, 2.4);
        const bsRGB = b <= 0.03928 ? b / 12.92 : Math.pow((b + 0.055) / 1.055, 2.4);
        
        return 0.2126 * rsRGB + 0.7152 * gsRGB + 0.0722 * bsRGB;
      } catch (error) {
        console.warn('明度計算エラー:', error);
        return 0.5; // デフォルト値
      }
    }
  
    /**
     * 色のコントラスト比を計算
     */
    public calculateContrastRatio(color1: string, color2: string): number {
      const lum1 = this.calculateLuminance(color1);
      const lum2 = this.calculateLuminance(color2);
      
      const lighter = Math.max(lum1, lum2);
      const darker = Math.min(lum1, lum2);
      
      return (lighter + 0.05) / (darker + 0.05);
    }
  
    /**
     * アクセシビリティ準拠の色組み合わせかチェック
     */
    public checkColorAccessibility(foreground: string, background: string): {
      ratio: number;
      isAACompliant: boolean;
      isAAACompliant: boolean;
      recommendation: string;
    } {
      const ratio = this.calculateContrastRatio(foreground, background);
      const isAACompliant = ratio >= 4.5;
      const isAAACompliant = ratio >= 7;
      
      let recommendation = '';
      if (!isAACompliant) {
        recommendation = 'コントラスト比が低すぎます。より濃い色の組み合わせを推奨します。';
      } else if (!isAAACompliant) {
        recommendation = 'AA基準は満たしていますが、より高いコントラストを推奨します。';
      } else {
        recommendation = 'アクセシビリティ基準を満たしています。';
      }
      
      return {
        ratio: Math.round(ratio * 100) / 100,
        isAACompliant,
        isAAACompliant,
        recommendation
      };
    }
  
    /**
     * テーマの一貫性を分析
     */
    public analyzeThemeConsistency(colorScheme: ColorScheme): {
      score: number;
      issues: string[];
      suggestions: string[];
    } {
      const issues: string[] = [];
      const suggestions: string[] = [];
      let score = 100;
  
      // 背景とテキストのコントラストチェック
      const bgTextContrast = this.calculateContrastRatio(colorScheme.background1, colorScheme.text1);
      if (bgTextContrast < 4.5) {
        issues.push('背景とテキストのコントラストが不十分');
        suggestions.push('テキスト色を濃くするか、背景色を薄くしてください');
        score -= 30;
      }
  
      // アクセント色の調和チェック
      const accentColors = [
        colorScheme.accent1,
        colorScheme.accent2,
        colorScheme.accent3
      ];
      
      // 簡易的な色調和チェック（実際の実装では色相環を考慮）
      if (this.areColorsClashing(accentColors)) {
        issues.push('アクセント色の組み合わせが調和していません');
        suggestions.push('類似色相または補色を使用することを推奨');
        score -= 20;
      }
  
      return {
        score: Math.max(0, score),
        issues,
        suggestions
      };
    }
  
    /**
     * 色が衝突しているかを簡易判定
     */
    private areColorsClashing(colors: string[]): boolean {
      // 簡易的な実装：同じような明度の彩度の高い色が複数ある場合を衝突と判定
      const luminances = colors.map(color => this.calculateLuminance(color));
      const luminanceRange = Math.max(...luminances) - Math.min(...luminances);
      
      // 明度の差が小さい場合（0.3未満）は衝突の可能性あり
      return luminanceRange < 0.3 && colors.length > 2;
    }
  
    /**
     * 推奨色パレットを生成
     */
    public generateRecommendedPalette(baseColor: string): {
      primary: string;
      secondary: string;
      accent: string;
      background: string;
      text: string;
    } {
      // 基本色から調和する色パレットを生成
      // 実際の実装では色相環とカラーハーモニー理論を使用
      
      const baseLuminance = this.calculateLuminance(baseColor);
      
      return {
        primary: baseColor,
        secondary: this.adjustColorLuminance(baseColor, 0.2),
        accent: this.generateComplementaryColor(baseColor),
        background: baseLuminance > 0.5 ? '#FFFFFF' : '#F5F5F5',
        text: baseLuminance > 0.5 ? '#333333' : '#FFFFFF'
      };
    }
  
    /**
     * 色の明度を調整
     */
    private adjustColorLuminance(color: string, amount: number): string {
      try {
        const hex = color.replace('#', '');
        const r = Math.min(255, Math.max(0, parseInt(hex.substr(0, 2), 16) + amount * 255));
        const g = Math.min(255, Math.max(0, parseInt(hex.substr(2, 2), 16) + amount * 255));
        const b = Math.min(255, Math.max(0, parseInt(hex.substr(4, 2), 16) + amount * 255));
        
        const rHex = Math.round(r).toString(16).padStart(2, '0');
        const gHex = Math.round(g).toString(16).padStart(2, '0');
        const bHex = Math.round(b).toString(16).padStart(2, '0');
        
        return `#${rHex}${gHex}${bHex}`;
      } catch (error) {
        console.warn('色調整エラー:', error);
        return color;
      }
    }
  
    /**
     * 補色を生成（簡易版）
     */
    private generateComplementaryColor(color: string): string {
      try {
        const hex = color.replace('#', '');
        const r = 255 - parseInt(hex.substr(0, 2), 16);
        const g = 255 - parseInt(hex.substr(2, 2), 16);
        const b = 255 - parseInt(hex.substr(4, 2), 16);
        
        const rHex = r.toString(16).padStart(2, '0');
        const gHex = g.toString(16).padStart(2, '0');
        const bHex = b.toString(16).padStart(2, '0');
        
        return `#${rHex}${gHex}${bHex}`;
      } catch (error) {
        console.warn('補色生成エラー:', error);
        return '#666666';
      }
    }
  
    /**
     * テーマ分析レポートを生成
     */
    public generateThemeAnalysisReport(themeInfo: any): {
      summary: string;
      colorAnalysis: any;
      fontAnalysis: any;
      accessibility: any;
      recommendations: string[];
    } {
      const colorAnalysis = this.analyzeThemeConsistency(themeInfo.colorScheme);
      const accessibility = this.checkColorAccessibility(
        themeInfo.colorScheme.text1,
        themeInfo.colorScheme.background1
      );
  
      const recommendations: string[] = [];
      
      if (colorAnalysis.score < 80) {
        recommendations.push('テーマの色の調和を改善することを推奨');
      }
      
      if (!accessibility.isAACompliant) {
        recommendations.push('アクセシビリティ向上のため、テキストと背景のコントラストを改善');
      }
  
      recommendations.push(...colorAnalysis.suggestions);
  
      return {
        summary: `テーマ品質スコア: ${colorAnalysis.score}/100`,
        colorAnalysis,
        fontAnalysis: {
          majorFont: themeInfo.fontScheme.majorFont.latin,
          minorFont: themeInfo.fontScheme.minorFont.latin,
          readability: 'good' // 実際の実装では詳細なフォント解析
        },
        accessibility,
        recommendations
      };
    }
  }