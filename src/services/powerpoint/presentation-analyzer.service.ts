// src/services/powerpoint/presentation-analyzer.service.ts - プレゼンテーション分析サービス
/* global PowerPoint */

import { SlideInfo, PresentationStats } from './types';

/**
 * プレゼンテーションの分析と情報取得を担当するサービス
 */
export class PresentationAnalyzer {

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
            const slideInfo = await this.extractSlideInfo(context, slide, i);
            slideInfos.push(slideInfo);
          }

          resolve(slideInfos);
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * プレゼンテーション全体の統計を取得
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
            const wordCount = await this.countWordsInSlide(context, slide);
            totalWords += wordCount;
          }

          // 1スライドあたり平均2分として概算
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
   * 個別スライドの情報を抽出
   */
  private async extractSlideInfo(
    context: PowerPoint.RequestContext, 
    slide: PowerPoint.Slide, 
    index: number
  ): Promise<SlideInfo> {
    slide.load("shapes");
    await context.sync();

    let title = `スライド ${index + 1}`;
    let content = '';

    // シェイプからテキストを抽出
    for (let j = 0; j < slide.shapes.items.length; j++) {
      const shape = slide.shapes.items[j];
      if (this.isTextShape(shape)) {
        shape.textFrame.load("textRange");
        await context.sync();
        
        const text = shape.textFrame.textRange.text;
        if (j === 0 && text) {
          title = text.substring(0, 50);
        }
        content += text + '\n';
      }
    }

    return {
      id: slide.id,
      title: title,
      content: content.trim(),
      index: index
    };
  }

  /**
   * スライド内の単語数をカウント
   */
  private async countWordsInSlide(context: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<number> {
    slide.shapes.load("items");
    await context.sync();

    let wordCount = 0;

    for (let j = 0; j < slide.shapes.items.length; j++) {
      const shape = slide.shapes.items[j];
      if (this.isTextShape(shape)) {
        shape.textFrame.load("textRange");
        await context.sync();
        
        const text = shape.textFrame.textRange.text;
        wordCount += this.countWords(text);
      }
    }

    return wordCount;
  }

  /**
   * シェイプがテキストシェイプかどうかを判定
   */
  private isTextShape(shape: PowerPoint.Shape): boolean {
    return shape.type === PowerPoint.ShapeType.textBox || 
           shape.type === PowerPoint.ShapeType.placeholder;
  }

  /**
   * テキスト内の単語数をカウント
   */
  private countWords(text: string): number {
    return text.split(/\s+/).filter(word => word.length > 0).length;
  }

  /**
   * スライドのサムネイル情報を取得（将来の拡張用）
   */
  public async getSlideThumbnails(): Promise<string[]> {
    // PowerPoint.js では現在サムネイル取得のサポートが限定的
    console.log('サムネイル取得機能は将来の実装予定です');
    return [];
  }

  /**
   * プレゼンテーションのメタデータを取得
   */
  public async getPresentationMetadata(): Promise<{
    title?: string;
    author?: string;
    subject?: string;
    lastModified?: Date;
  }> {
    return new Promise((resolve) => {
      PowerPoint.run(async (_) => {
        try {
          // PowerPoint.js では現在メタデータアクセスが限定的
          console.log('プレゼンテーションメタデータの取得機能は将来の実装予定です');
          resolve({});
        } catch (error) {
          console.warn('メタデータの取得に失敗しました:', error);
          resolve({});
        }
      });
    });
  }
}