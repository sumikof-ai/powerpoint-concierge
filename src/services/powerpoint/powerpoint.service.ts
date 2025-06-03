// src/services/powerpoint/powerpoint.service.ts - 修正版（実用的なテーマ対応）
/* global PowerPoint */

import { 
    SlideContent, 
    SlideGenerationOptions, 
    BulkSlideData, 
    SlideInfo,
    PresentationStats 
  } from './types';
  
  /**
   * PowerPoint操作のメインサービスクラス（実用的なテーマ対応版）
   * PowerPoint.js APIの制限を考慮した実装
   */
  export class PowerPointService {
    private defaultOptions: SlideGenerationOptions = {
      includeTransitions: false,
      slideLayout: 'content',
      theme: 'light',
      fontSize: 'medium',
      useThemeAwareGeneration: true
    };
  
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
     * 複数のスライドを一括生成（テーマ対応版）
     */
    public async generateBulkSlides(
      bulkData: BulkSlideData, 
      onProgress?: (current: number, total: number, slideName: string) => void
    ): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const { slides, options = {} } = bulkData;
            const mergedOptions = { ...this.defaultOptions, ...options };
            
            // 現在のプレゼンテーションの基本情報を取得
            const presentationInfo = await this.getCurrentPresentationInfo(context);
            
            // 各スライドを順番に生成
            for (let i = 0; i < slides.length; i++) {
              const slideData = slides[i];
              
              // 進捗コールバック
              if (onProgress) {
                onProgress(i + 1, slides.length, slideData.title);
              }
  
              // テーマ対応でスライドを作成
              await this.createSlideWithThemeSupport(
                context, 
                slideData, 
                mergedOptions, 
                presentationInfo
              );
            }
  
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * 現在のプレゼンテーション情報を取得
     */
    private async getCurrentPresentationInfo(context: PowerPoint.RequestContext) {
      const presentation = context.presentation;
      presentation.load("title");
      
      const slides = presentation.slides;
      slides.load("items");
      await context.sync();
  
      // 最初のスライドがある場合、そこからスタイル情報を推測
      let existingSlideInfo = null;
      if (slides.items.length > 0) {
        const firstSlide = slides.items[0];
        firstSlide.shapes.load("items");
        await context.sync();
  
        // 既存のスライドから色やフォント情報を抽出を試みる
        existingSlideInfo = await this.extractStyleFromSlide(context, firstSlide);
      }
  
      return {
        slideCount: slides.items.length,
        title: presentation.title || "Untitled Presentation",
        existingStyle: existingSlideInfo
      };
    }
  
    /**
     * 既存スライドからスタイル情報を抽出
     */
    private async extractStyleFromSlide(context: PowerPoint.RequestContext, slide: PowerPoint.Slide) {
      const styleInfo = {
        commonFontSize: 18,
        commonFontColor: '#000000',
        backgroundDetected: false
      };
  
      try {
        // シェイプからスタイル情報を抽出
        for (let i = 0; i < Math.min(slide.shapes.items.length, 3); i++) {
          const shape = slide.shapes.items[i];
          
          if (shape.type === PowerPoint.ShapeType.textBox || 
              shape.type === PowerPoint.ShapeType.placeholder) {
            shape.textFrame.load("textRange");
            await context.sync();
  
            // フォントサイズを取得
            if (shape.textFrame.textRange.font) {
              shape.textFrame.textRange.font.load("size");
              await context.sync();
              
              const fontSize = shape.textFrame.textRange.font.size;
              if (fontSize > 0) {
                styleInfo.commonFontSize = fontSize;
              }
            }
          }
        }
      } catch (error) {
        console.log('スタイル抽出中にエラー（無視して続行）:', error);
      }
  
      return styleInfo;
    }
  
    /**
     * テーマ対応でスライドを作成
     */
    private async createSlideWithThemeSupport(
      context: PowerPoint.RequestContext,
      slideData: SlideContent,
      options: SlideGenerationOptions,
      presentationInfo: any
    ): Promise<void> {
      // 新しいスライドを追加
      context.presentation.slides.add();
      await context.sync();
      
      // 最後に追加されたスライドを取得
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      const slide = slides.items[slides.items.length - 1];
  
      // テーマ設定を取得
      const themeColors = this.themePresets[options.theme || 'light'];
      const fontSize = this.getFontSizes(options.fontSize);
  
      // スライドタイプに応じて作成
      switch (slideData.slideType) {
        case 'title':
          await this.createTitleSlideWithTheme(context, slide, slideData, themeColors, fontSize);
          break;
        case 'conclusion':
          await this.createConclusionSlideWithTheme(context, slide, slideData, themeColors, fontSize);
          break;
        default:
          await this.createContentSlideWithTheme(context, slide, slideData, themeColors, fontSize);
      }
  
      // スピーカーノートがある場合はコンソールに出力
      if (slideData.speakerNotes) {
        console.log(`📝 スピーカーノート [${slideData.title}]: ${slideData.speakerNotes}`);
        console.log(`presentationInfo:${presentationInfo}`)
      }
    }
  
    /**
     * テーマ対応タイトルスライドを作成
     */
    private async createTitleSlideWithTheme(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      themeColors: any,
      fontSize: any
    ): Promise<void> {
      // メインタイトル
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 60,
        top: 150,
        width: 620,
        height: 120
      });
      
      await context.sync();
      
      // タイトルのスタイルを設定
      titleBox.textFrame.textRange.font.size = fontSize.title;
      titleBox.textFrame.textRange.font.bold = true;
      titleBox.textFrame.textRange.font.color = themeColors.titleColor;
      
      // 背景とボーダーを設定
      titleBox.fill.setSolidColor(themeColors.backgroundColor);
      if (titleBox.lineFormat) {
        titleBox.lineFormat.color = themeColors.borderColor;
        titleBox.lineFormat.weight = 2;
      }
  
      // サブタイトル（コンテンツがある場合）
      if (slideData.content && slideData.content.length > 0) {
        const subtitleText = slideData.content.join(' • ');
        const subtitleBox = slide.shapes.addTextBox(subtitleText, {
          left: 100,
          top: 300,
          width: 540,
          height: 80
        });
        
        await context.sync();
        
        subtitleBox.textFrame.textRange.font.size = fontSize.subtitle;
        subtitleBox.textFrame.textRange.font.color = themeColors.textColor;
        subtitleBox.fill.setSolidColor(themeColors.backgroundColor);
      }
  
      // アクセント要素（装飾）
      const accentShape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 50,
        top: 280,
        width: 640,
        height: 4
      });
      
      await context.sync();
      accentShape.fill.setSolidColor(themeColors.accentColor);
    }
  
    /**
     * テーマ対応コンテンツスライドを作成
     */
    private async createContentSlideWithTheme(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      themeColors: any,
      fontSize: any
    ): Promise<void> {
      // タイトル
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 50,
        top: 30,
        width: 640,
        height: 70
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
      titleBox.textFrame.textRange.font.color = themeColors.titleColor;
      titleBox.fill.setSolidColor(themeColors.backgroundColor);
      
      // タイトル下のアクセントライン
      const titleUnderline = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 50,
        top: 105,
        width: 100,
        height: 3
      });
      await context.sync();
      titleUnderline.fill.setSolidColor(themeColors.accentColor);
  
      // コンテンツ（箇条書き）
      if (slideData.content && slideData.content.length > 0) {
        // コンテンツが多い場合は2カラムに分割
        if (slideData.content.length > 6) {
          await this.createTwoColumnContent(context, slide, slideData.content, themeColors, fontSize);
        } else {
          await this.createSingleColumnContent(context, slide, slideData.content, themeColors, fontSize);
        }
      }
    }
  
    /**
     * 単一カラムコンテンツを作成
     */
    private async createSingleColumnContent(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      content: string[],
      themeColors: any,
      fontSize: any
    ): Promise<void> {
      const contentText = content.map((item, index) => {
        const bullet = index === 0 ? '●' : '◦';
        return `${bullet} ${item}`;
      }).join('\n\n');
  
      const contentBox = slide.shapes.addTextBox(contentText, {
        left: 70,
        top: 130,
        width: 600,
        height: 380
      });
      
      await context.sync();
      
      contentBox.textFrame.textRange.font.size = fontSize.body;
      contentBox.textFrame.textRange.font.color = themeColors.textColor;
      contentBox.fill.setSolidColor(themeColors.backgroundColor);
      
      // コンテンツエリアの枠線
      if (contentBox.lineFormat) {
        contentBox.lineFormat.color = themeColors.borderColor;
        contentBox.lineFormat.weight = 1;
      }
    }
  
    /**
     * 2カラムコンテンツを作成
     */
    private async createTwoColumnContent(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      content: string[],
      themeColors: any,
      fontSize: any
    ): Promise<void> {
      const midPoint = Math.ceil(content.length / 2);
      const leftContent = content.slice(0, midPoint);
      const rightContent = content.slice(midPoint);
  
      // 左カラム
      if (leftContent.length > 0) {
        const leftText = leftContent.map(item => `• ${item}`).join('\n\n');
        const leftBox = slide.shapes.addTextBox(leftText, {
          left: 50,
          top: 130,
          width: 300,
          height: 380
        });
        
        await context.sync();
        
        leftBox.textFrame.textRange.font.size = fontSize.body;
        leftBox.textFrame.textRange.font.color = themeColors.textColor;
        leftBox.fill.setSolidColor(themeColors.backgroundColor);
      }
  
      // 右カラム
      if (rightContent.length > 0) {
        const rightText = rightContent.map(item => `• ${item}`).join('\n\n');
        const rightBox = slide.shapes.addTextBox(rightText, {
          left: 380,
          top: 130,
          width: 300,
          height: 380
        });
        
        await context.sync();
        
        rightBox.textFrame.textRange.font.size = fontSize.body;
        rightBox.textFrame.textRange.font.color = themeColors.textColor;
        rightBox.fill.setSolidColor(themeColors.backgroundColor);
      }
  
      // 分割線
      const dividerLine = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 360,
        top: 130,
        width: 2,
        height: 300
      });
      await context.sync();
      dividerLine.fill.setSolidColor(themeColors.borderColor);
    }
  
    /**
     * まとめスライドを作成
     */
    private async createConclusionSlideWithTheme(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      themeColors: any,
      fontSize: any
    ): Promise<void> {
      // 「まとめ」ラベル
      const labelBox = slide.shapes.addTextBox("まとめ", {
        left: 50,
        top: 30,
        width: 150,
        height: 50
      });
      
      await context.sync();
      labelBox.textFrame.textRange.font.size = fontSize.accent;
      labelBox.textFrame.textRange.font.bold = true;
      labelBox.textFrame.textRange.font.color = themeColors.backgroundColor;
      labelBox.fill.setSolidColor(themeColors.accentColor);
  
      // メインタイトル
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 220,
        top: 30,
        width: 470,
        height: 70
      });
      
      await context.sync();
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
      titleBox.textFrame.textRange.font.color = themeColors.titleColor;
      titleBox.fill.setSolidColor(themeColors.backgroundColor);
  
      // コンテンツ（重要ポイント）
      if (slideData.content && slideData.content.length > 0) {
        const contentText = slideData.content.map((item, index) => 
          `${index + 1}. ${item}`
        ).join('\n\n');
  
        const contentBox = slide.shapes.addTextBox(contentText, {
          left: 80,
          top: 130,
          width: 580,
          height: 300
        });
        
        await context.sync();
        
        contentBox.textFrame.textRange.font.size = fontSize.body + 2;
        contentBox.textFrame.textRange.font.color = themeColors.textColor;
        contentBox.fill.setSolidColor(themeColors.backgroundColor);
      }
  
      // 装飾フレーム
      const frameShape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 40,
        top: 120,
        width: 620,
        height: 330
      });
      
      await context.sync();
      frameShape.fill.clear();
      if (frameShape.lineFormat) {
        frameShape.lineFormat.color = themeColors.accentColor;
        frameShape.lineFormat.weight = 3;
      }
    }
  
    /**
     * フォントサイズを取得
     */
    private getFontSizes(size?: 'small' | 'medium' | 'large') {
      switch (size) {
        case 'small':
          return { title: 36, heading: 22, subtitle: 18, body: 14, accent: 16 };
        case 'large':
          return { title: 48, heading: 32, subtitle: 26, body: 18, accent: 20 };
        case 'medium':
        default:
          return { title: 42, heading: 26, subtitle: 20, body: 16, accent: 18 };
      }
    }
  
    // ===== 既存のメソッド（従来版との互換性保持） =====
  
    /**
     * 現在のプレゼンテーションの全スライド情報を取得
     */
    public async getAllSlides(): Promise<SlideInfo[]> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
  
            const slideInfos: SlideInfo[] = [];
            
            for (let i = 0; i < slides.items.length; i++) {
              const slide = slides.items[i];
              slide.load("shapes");
              await context.sync();
  
              let title = `スライド ${i + 1}`;
              let content = '';
  
              // シェイプからテキストを抽出
              for (let j = 0; j < slide.shapes.items.length; j++) {
                const shape = slide.shapes.items[j];
                if (shape.type === PowerPoint.ShapeType.textBox || 
                    shape.type === PowerPoint.ShapeType.placeholder) {
                  shape.textFrame.load("textRange");
                  await context.sync();
                  
                  const text = shape.textFrame.textRange.text;
                  if (j === 0 && text) {
                    title = text.substring(0, 50);
                  }
                  content += text + '\n';
                }
              }
  
              slideInfos.push({
                id: slide.id,
                title: title,
                content: content.trim(),
                index: i
              });
            }
  
            resolve(slideInfos);
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * 新しいスライドを追加（従来版との互換性）
     */
    public async addSlide(title: string, content: string): Promise<void> {
      const slideData: SlideContent = {
        title,
        content: content.split('\n• ').filter(item => item.trim() !== ''),
        slideType: 'content'
      };
      
      const bulkData: BulkSlideData = {
        slides: [slideData],
        options: this.defaultOptions
      };
      
      return this.generateBulkSlides(bulkData);
    }
  
    /**
     * プレゼンテーション統計を取得
     */
    public async getPresentationStats(): Promise<PresentationStats> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
  
            let totalWords = 0;
            
            for (let i = 0; i < slides.items.length; i++) {
              const slide = slides.items[i];
              slide.shapes.load("items");
              await context.sync();
  
              for (let j = 0; j < slide.shapes.items.length; j++) {
                const shape = slide.shapes.items[j];
                if (shape.type === PowerPoint.ShapeType.textBox || 
                    shape.type === PowerPoint.ShapeType.placeholder) {
                  shape.textFrame.load("textRange");
                  await context.sync();
                  
                  const text = shape.textFrame.textRange.text;
                  totalWords += text.split(/\s+/).filter(word => word.length > 0).length;
                }
              }
            }
  
            const estimatedDuration = slides.items.length * 2;
  
            resolve({
              slideCount: slides.items.length,
              estimatedDuration,
              wordCount: totalWords
            });
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * テストメソッド - テーマ適用の動作確認
     */
    public async testThemeApplication(): Promise<void> {
      const testSlide: SlideContent = {
        title: "テーマテスト",
        content: ["ライトテーマのテスト", "テキストの色とスタイル", "アクセントカラーの適用"],
        slideType: 'content'
      };
  
      // 各テーマで同じスライドを作成
      for (const themeName of ['light', 'dark', 'colorful'] as const) {
        const bulkData: BulkSlideData = {
          slides: [{
            ...testSlide,
            title: `${testSlide.title} - ${themeName.toUpperCase()}テーマ`
          }],
          options: {
            ...this.defaultOptions,
            theme: themeName
          }
        };
  
        await this.generateBulkSlides(bulkData);
      }
    }
  }