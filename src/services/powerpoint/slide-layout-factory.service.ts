// src/services/powerpoint/slide-layout-factory.service.ts - レイアウト作成ファクトリ
/* global PowerPoint */

import { SlideThemeApplier } from './slide-theme-applier.service';
import { SlideContent, SlideGenerationOptions, SlideLayoutType, ShapeOptions } from './types';

/**
 * 様々なスライドレイアウトを作成するファクトリクラス
 */
export class SlideLayoutFactory {
  constructor(private themeApplier: SlideThemeApplier) {}

  /**
   * 指定されたレイアウトでスライドを作成
   */
  public async createSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    layout: SlideLayoutType,
    options: SlideGenerationOptions
  ): Promise<void> {
    switch (layout) {
      case 'title':
        await this.createTitleSlide(context, slide, slideData, options);
        break;
      case 'content':
        await this.createContentSlide(context, slide, slideData, options);
        break;
      case 'twoContent':
        await this.createTwoContentSlide(context, slide, slideData, options);
        break;
      case 'comparison':
        await this.createComparisonSlide(context, slide, slideData, options);
        break;
      case 'blank':
        await this.createBlankSlide(context, slide, slideData, options);
        break;
      default:
        await this.createContentSlide(context, slide, slideData, options);
    }
  }

  /**
   * タイトルスライドを作成
   */
  private async createTitleSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);
    
    // メインタイトル
    const titleBox = slide.shapes.addTextBox(slideData.title, {
      left: 75,
      top: 150,
      width: 600,
      height: 150
    });
    
    await context.sync();
    
    titleBox.textFrame.textRange.font.size = fontSize.title;
    titleBox.textFrame.textRange.font.bold = true;
    this.themeApplier.applyThemeColors(titleBox, options.theme, 'title');

    // サブタイトル
    if (slideData.content && slideData.content.length > 0) {
      const subtitleText = slideData.content.join(' | ');
      const subtitleBox = slide.shapes.addTextBox(subtitleText, {
        left: 100,
        top: 320,
        width: 550,
        height: 100
      });
      
      await context.sync();
      
      subtitleBox.textFrame.textRange.font.size = fontSize.subtitle;
      this.themeApplier.applyThemeColors(subtitleBox, options.theme, 'subtitle');
    }
  }

  /**
   * 標準コンテンツスライドを作成
   */
  private async createContentSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);
    
    // タイトル
    const titleBox = slide.shapes.addTextBox(slideData.title, {
      left: 50,
      top: 40,
      width: 650,
      height: 80
    });
    
    await context.sync();
    
    titleBox.textFrame.textRange.font.size = fontSize.heading;
    titleBox.textFrame.textRange.font.bold = true;
    this.themeApplier.applyThemeColors(titleBox, options.theme, 'heading');

    // コンテンツ
    if (slideData.content && slideData.content.length > 0) {
      const contentText = slideData.content.map(item => `• ${item}`).join('\n\n');
      const contentBox = slide.shapes.addTextBox(contentText, {
        left: 80,
        top: 140,
        width: 580,
        height: 350
      });
      
      await context.sync();
      
      contentBox.textFrame.textRange.font.size = fontSize.body;
      this.themeApplier.applyThemeColors(contentBox, options.theme, 'body');
    }
  }

  /**
   * 2カラムコンテンツスライドを作成
   */
  private async createTwoContentSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);
    
    // タイトル
    await this.createTitle(context, slide, slideData.title, fontSize, options);

    // コンテンツを2つに分割
    if (slideData.content && slideData.content.length > 0) {
      const midPoint = Math.ceil(slideData.content.length / 2);
      const leftContent = slideData.content.slice(0, midPoint);
      const rightContent = slideData.content.slice(midPoint);

      // 左カラム
      if (leftContent.length > 0) {
        await this.createContentColumn(context, slide, leftContent, {
          left: 50,
          top: 140,
          width: 300,
          height: 350
        }, fontSize, options);
      }

      // 右カラム
      if (rightContent.length > 0) {
        await this.createContentColumn(context, slide, rightContent, {
          left: 380,
          top: 140,
          width: 300,
          height: 350
        }, fontSize, options);
      }
    }
  }

  /**
   * 比較スライドを作成
   */
  private async createComparisonSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);
    
    // タイトル
    await this.createTitle(context, slide, slideData.title, fontSize, options);

    // ヘッダー
    await this.createComparisonHeaders(context, slide, fontSize, options);

    // コンテンツを交互に配置
    if (slideData.content && slideData.content.length > 0) {
      for (let index = 0; index < slideData.content.length && index < 8; index++) {
        const item = slideData.content[index];
        const yPos = 200 + (index * 35);
        const isLeft = index % 2 === 0;
        
        const contentBox = slide.shapes.addTextBox(`• ${item}`, {
          left: isLeft ? 50 : 380,
          top: yPos,
          width: 300,
          height: 30
        });
        
        await context.sync();
        
        contentBox.textFrame.textRange.font.size = fontSize.body;
        this.themeApplier.applyThemeColors(contentBox, options.theme, 'body');
      }
    }
  }

  /**
   * 空白スライドを作成
   */
  private async createBlankSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);
    await this.createTitle(context, slide, slideData.title, fontSize, options);
  }

  /**
   * タイトルを作成（共通処理）
   */
  private async createTitle(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    title: string,
    fontSize: any,
    options: SlideGenerationOptions
  ): Promise<void> {
    const titleBox = slide.shapes.addTextBox(title, {
      left: 50,
      top: 40,
      width: 650,
      height: 80
    });
    
    await context.sync();
    
    titleBox.textFrame.textRange.font.size = fontSize.heading;
    titleBox.textFrame.textRange.font.bold = true;
    this.themeApplier.applyThemeColors(titleBox, options.theme, 'heading');
  }

  /**
   * コンテンツカラムを作成（共通処理）
   */
  private async createContentColumn(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    content: string[],
    shapeOptions: ShapeOptions,
    fontSize: any,
    options: SlideGenerationOptions
  ): Promise<void> {
    const contentText = content.map(item => `• ${item}`).join('\n\n');
    const contentBox = slide.shapes.addTextBox(contentText, shapeOptions);
    
    await context.sync();
    
    contentBox.textFrame.textRange.font.size = fontSize.body;
    this.themeApplier.applyThemeColors(contentBox, options.theme, 'body');
  }

  /**
   * 比較スライドのヘッダーを作成
   */
  private async createComparisonHeaders(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    fontSize: any,
    options: SlideGenerationOptions
  ): Promise<void> {
    // 左側ヘッダー
    const leftHeaderBox = slide.shapes.addTextBox("項目", {
      left: 50,
      top: 140,
      width: 300,
      height: 40
    });
    
    await context.sync();
    
    leftHeaderBox.textFrame.textRange.font.bold = true;
    leftHeaderBox.textFrame.textRange.font.size = fontSize.accent;
    this.themeApplier.applyThemeColors(leftHeaderBox, options.theme, 'accent');

    // 右側ヘッダー
    const rightHeaderBox = slide.shapes.addTextBox("詳細", {
      left: 380,
      top: 140,
      width: 300,
      height: 40
    });
    
    await context.sync();
    
    rightHeaderBox.textFrame.textRange.font.bold = true;
    rightHeaderBox.textFrame.textRange.font.size = fontSize.accent;
    this.themeApplier.applyThemeColors(rightHeaderBox, options.theme, 'accent');
  }
}