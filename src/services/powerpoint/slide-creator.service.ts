// src/services/powerpoint/slide-creator.service.ts - スライド作成専用サービス
/* global PowerPoint */

import { SlideThemeApplier } from './slide-theme-applier.service';
import { SlideLayoutFactory } from './slide-layout-factory.service';
import { 
  SlideContent, 
  SlideGenerationOptions, 
  BulkSlideData, 
  SlideLayoutType 
} from './types';

/**
 * スライド作成を担当するサービス
 */
export class SlideCreator {
  private themeApplier: SlideThemeApplier;
  private layoutFactory: SlideLayoutFactory;
  private defaultOptions: SlideGenerationOptions = {
    includeTransitions: false,
    slideLayout: 'content',
    theme: 'light',
    fontSize: 'medium',
  };

  constructor() {
    this.themeApplier = new SlideThemeApplier();
    this.layoutFactory = new SlideLayoutFactory(this.themeApplier);
  }

  /**
   * 複数のスライドを一括生成
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
          
          // 各スライドを順番に生成
          for (let i = 0; i < slides.length; i++) {
            const slideData = slides[i];
            
            // 進捗コールバック
            if (onProgress) {
              onProgress(i + 1, slides.length, slideData.title);
            }

            // スライドレイアウトの決定
            const layout = this.determineSlideLayout(slideData.slideType, mergedOptions.slideLayout);
            
            // スライドを作成
            await this.createSlideWithLayout(context, slideData, layout, mergedOptions);
          }

          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * 単一のスライドを作成
   */
  public async createSingleSlide(slideData: SlideContent, options?: SlideGenerationOptions): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const mergedOptions = { ...this.defaultOptions, ...options };
          const layout = this.determineSlideLayout(slideData.slideType, mergedOptions.slideLayout);
          await this.createSlideWithLayout(context, slideData, layout, mergedOptions);
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * テキストボックスを追加（テスト用）
   */
  public async addTextBox(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          let slide;
          try {
            slide = context.presentation.getSelectedSlides().getItemAt(0);
          } catch {
            slide = context.presentation.slides.getItemAt(0);
          }

          const textBox = slide.shapes.addTextBox(text, {
            left: 100,
            top: 200,
            width: 500,
            height: 200
          });
          
          await context.sync();
          
          textBox.textFrame.textRange.font.size = 14;
          textBox.fill.setSolidColor("white");
          if (textBox.lineFormat) {
            textBox.lineFormat.color = "blue";
            textBox.lineFormat.weight = 1;
            textBox.lineFormat.dashStyle = "solid";
          }

          await context.sync();
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * スライドタイプに基づいてレイアウトを決定
   */
  private determineSlideLayout(
    slideType: 'title' | 'content' | 'conclusion', 
    defaultLayout?: string
  ): SlideLayoutType {
    switch (slideType) {
      case 'title':
        return 'title';
      case 'conclusion':
        return 'content';
      case 'content':
      default:
        return (defaultLayout as SlideLayoutType) || 'content';
    }
  }

  /**
   * 指定されたレイアウトでスライドを作成
   */
  private async createSlideWithLayout(
    context: PowerPoint.RequestContext,
    slideData: SlideContent,
    layout: SlideLayoutType,
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
    
    // レイアウトファクトリーを使用してスライドを作成
    await this.layoutFactory.createSlide(context, slide, slideData, layout, options);

    // スピーカーノートのログ出力
    if (slideData.speakerNotes) {
      console.log(`スピーカーノート (${slideData.title}): ${slideData.speakerNotes}`);
    }
  }
}