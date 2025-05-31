// src/services/powerpoint/slide-manager.service.ts - スライド管理サービス
/* global PowerPoint */

import { SlideContent, SlideGenerationOptions } from './types';
import { SlideLayoutFactory } from './slide-layout-factory.service';
import { SlideThemeApplier } from './slide-theme-applier.service';

/**
 * スライドの更新、削除、管理を担当するサービス
 */
export class SlideManager {
  private layoutFactory: SlideLayoutFactory;
  private themeApplier: SlideThemeApplier;
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
   * 指定したスライドのコンテンツを更新
   */
  public async updateSlide(slideIndex: number, title: string, content: string): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          if (slideIndex >= slides.items.length) {
            throw new Error(`スライド ${slideIndex + 1} が見つかりません`);
          }

          const slide = slides.items[slideIndex];
          
          // 既存のテキストボックスをクリア
          await this.clearSlideTextBoxes(context, slide);

          // 新しいコンテンツで再作成
          const slideData: SlideContent = {
            title,
            content: this.parseContentString(content),
            slideType: 'content'
          };

          await this.layoutFactory.createSlide(context, slide, slideData, 'content', this.defaultOptions);
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * 指定したスライドを削除
   */
  public async deleteSlide(slideIndex: number): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          if (slideIndex >= slides.items.length) {
            throw new Error(`スライド ${slideIndex + 1} が見つかりません`);
          }

          slides.items[slideIndex].delete();
          await context.sync();
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * スライド間にトランジションを追加
   * 注意: PowerPoint.js では現在トランジション機能のサポートが限定的
   */
  public async addTransitions(transitionType: 'fade' | 'slide' | 'none' = 'fade'): Promise<void> {
    return new Promise((resolve) => {
      console.log(`トランジション設定をリクエストしました: ${transitionType}`);
      console.log('注意: PowerPoint.js では現在トランジション機能のサポートが限定的です');
      resolve();
    });
  }

  /**
   * スライドのテキストボックスをクリア
   */
  private async clearSlideTextBoxes(context: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<void> {
    slide.shapes.load("items");
    await context.sync();

    // 既存のテキストボックスをクリア
    for (let i = slide.shapes.items.length - 1; i >= 0; i--) {
      const shape = slide.shapes.items[i];
      if (shape.type === PowerPoint.ShapeType.textBox) {
        shape.delete();
      }
    }

    await context.sync();
  }

  /**
   * コンテンツ文字列をパース
   */
  private parseContentString(content: string): string[] {
    return content
      .split('\n• ')
      .map(item => item.replace(/^• /, '').trim())
      .filter(item => item !== '');
  }
}