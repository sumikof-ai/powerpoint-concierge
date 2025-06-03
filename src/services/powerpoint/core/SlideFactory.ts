// src/services/powerpoint/core/SlideFactory.ts - スライド作成ファクトリー
/* global PowerPoint */

import { SlideContent, SlideGenerationOptions, SlideLayoutType } from '../types';
import { ContentRenderer } from './ContentRenderer';
import { ThemeApplier } from './ThemeApplier';

/**
 * スライド作成とレイアウト決定を担当するファクトリークラス
 */
export class SlideFactory {
  private contentRenderer: ContentRenderer;
  private themeApplier: ThemeApplier;

  constructor() {
    this.contentRenderer = new ContentRenderer();
    this.themeApplier = new ThemeApplier();
  }

  /**
   * スライドタイプに基づいてレイアウトを決定
   */
  public determineSlideLayout(
    slideType: 'title' | 'content' | 'conclusion',
    contentAmount: number,
    defaultLayout?: string
  ): SlideLayoutType {
    switch (slideType) {
      case 'title':
        return 'title';
      case 'conclusion':
        return 'content';
      case 'content':
        // コンテンツ量に基づいて最適なレイアウトを選択
        if (contentAmount > 300) {
          return 'twoContent'; // 大量のコンテンツは2カラムに
        }
        return (defaultLayout as SlideLayoutType) || 'content';
      default:
        return 'content';
    }
  }

  /**
   * スライドを作成してコンテンツを配置
   */
  public async createSlideWithContent(
    context: PowerPoint.RequestContext,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    // 新しいスライドを追加
    context.presentation.slides.add();
    await context.sync();
    
    // 最後に追加されたスライドを取得
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    const slide = slides.items[slides.items.length - 1];

    // コンテンツ量を計算
    const contentAmount = this.calculateContentAmount(slideData);
    
    // 最適なレイアウトを決定
    const layout = this.determineSlideLayout(
      slideData.slideType,
      contentAmount,
      options.slideLayout
    );

    // レイアウトに応じてコンテンツを配置
    await this.renderSlideContent(context, slide, slideData, layout, options);

    // スピーカーノートのログ出力
    if (slideData.speakerNotes) {
      console.log(`📝 スピーカーノート [${slideData.title}]: ${slideData.speakerNotes}`);
    }
  }

  /**
   * 指定されたレイアウトでスライドコンテンツをレンダリング
   */
  private async renderSlideContent(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    layout: SlideLayoutType,
    options: SlideGenerationOptions
  ): Promise<void> {
    switch (layout) {
      case 'title':
        await this.contentRenderer.renderTitleSlide(context, slide, slideData, options);
        break;
      case 'content':
        await this.contentRenderer.renderContentSlide(context, slide, slideData, options);
        break;
      case 'twoContent':
        await this.contentRenderer.renderTwoContentSlide(context, slide, slideData, options);
        break;
      case 'comparison':
        await this.contentRenderer.renderComparisonSlide(context, slide, slideData, options);
        break;
      case 'blank':
        await this.contentRenderer.renderBlankSlide(context, slide, slideData, options);
        break;
      default:
        await this.contentRenderer.renderContentSlide(context, slide, slideData, options);
    }
  }

  /**
   * コンテンツ量を計算
   */
  private calculateContentAmount(slideData: SlideContent): number {
    const titleLength = slideData.title.length;
    const contentLength = slideData.content.reduce((sum, item) => sum + item.length, 0);
    return titleLength + contentLength;
  }

  /**
   * 複数のスライドを一括作成
   */
  public async createBulkSlides(
    context: PowerPoint.RequestContext,
    slides: SlideContent[],
    options: SlideGenerationOptions,
    onProgress?: (current: number, total: number, slideName: string) => void
  ): Promise<void> {
    for (let i = 0; i < slides.length; i++) {
      const slideData = slides[i];
      
      // 進捗報告
      if (onProgress) {
        onProgress(i + 1, slides.length, slideData.title);
      }

      // スライドを作成
      await this.createSlideWithContent(context, slideData, options);
      
      // PowerPoint APIの制限を避けるため、少し待機
      if (i < slides.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }
  }

  /**
   * レイアウトテンプレートを取得
   */
  public getLayoutTemplates(): { [key: string]: any } {
    return {
      title: {
        titlePosition: { left: 75, top: 150, width: 600, height: 150 },
        subtitlePosition: { left: 100, top: 320, width: 550, height: 100 }
      },
      content: {
        titlePosition: { left: 50, top: 40, width: 650, height: 80 },
        contentPosition: { left: 80, top: 140, width: 580, height: 350 }
      },
      twoContent: {
        titlePosition: { left: 50, top: 40, width: 650, height: 80 },
        leftContentPosition: { left: 50, top: 140, width: 300, height: 350 },
        rightContentPosition: { left: 380, top: 140, width: 300, height: 350 }
      },
      comparison: {
        titlePosition: { left: 50, top: 40, width: 650, height: 80 },
        leftHeaderPosition: { left: 50, top: 140, width: 300, height: 40 },
        rightHeaderPosition: { left: 380, top: 140, width: 300, height: 40 }
      }
    };
  }

  /**
   * スライドの品質チェック
   */
  public validateSlideContent(slideData: SlideContent): {
    isValid: boolean;
    warnings: string[];
    suggestions: string[];
  } {
    const warnings: string[] = [];
    const suggestions: string[] = [];

    // タイトルの長さチェック
    if (slideData.title.length > 100) {
      warnings.push('タイトルが長すぎます（100文字以内を推奨）');
    }

    // コンテンツ数のチェック
    if (slideData.content.length > 7) {
      warnings.push('コンテンツ項目が多すぎます（7項目以内を推奨）');
      suggestions.push('複数のスライドに分割することを検討してください');
    }

    // 各コンテンツ項目の長さチェック
    slideData.content.forEach((item, index) => {
      if (item.length > 150) {
        warnings.push(`項目 ${index + 1} が長すぎます（150文字以内を推奨）`);
      }
    });

    // 空のコンテンツチェック
    if (slideData.content.length === 0 && slideData.slideType !== 'title') {
      warnings.push('コンテンツが空です');
    }

    return {
      isValid: warnings.length === 0,
      warnings,
      suggestions
    };
  }

  /**
   * スライドの最適化提案
   */
  public suggestOptimizations(slideData: SlideContent): string[] {
    const suggestions: string[] = [];
    const contentAmount = this.calculateContentAmount(slideData);

    // コンテンツ量に基づく提案
    if (contentAmount > 500) {
      suggestions.push('コンテンツ量が多いため、2カラムレイアウトまたは複数スライドへの分割を推奨');
    }

    // スライドタイプに基づく提案
    if (slideData.slideType === 'title' && slideData.content.length > 2) {
      suggestions.push('タイトルスライドはシンプルに保つことを推奨（2項目以内）');
    }

    if (slideData.slideType === 'conclusion' && slideData.content.length > 5) {
      suggestions.push('まとめスライドは要点を絞ることを推奨（5項目以内）');
    }

    return suggestions;
  }
}