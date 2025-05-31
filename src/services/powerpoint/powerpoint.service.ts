// src/services/powerpoint/powerpoint.service.ts - メインサービス
/* global PowerPoint */

import { SlideCreator } from './slide-creator.service';
import { SlideManager } from './slide-manager.service';
import { PresentationAnalyzer } from './presentation-analyzer.service';
import { SlideInfo, SlideGenerationOptions, SlideContent, BulkSlideData } from './types';

/**
 * PowerPoint操作のメインサービスクラス
 * 各専門サービスを統合し、統一されたAPIを提供
 */
export class PowerPointService {
  private slideCreator: SlideCreator;
  private slideManager: SlideManager;
  private presentationAnalyzer: PresentationAnalyzer;

  constructor() {
    this.slideCreator = new SlideCreator();
    this.slideManager = new SlideManager();
    this.presentationAnalyzer = new PresentationAnalyzer();
  }

  /**
   * 現在のプレゼンテーションの全スライド情報を取得
   */
  public async getAllSlides(): Promise<SlideInfo[]> {
    return this.presentationAnalyzer.getAllSlides();
  }

  /**
   * 複数のスライドを一括生成
   */
  public async generateBulkSlides(
    bulkData: BulkSlideData, 
    onProgress?: (current: number, total: number, slideName: string) => void
  ): Promise<void> {
    return this.slideCreator.generateBulkSlides(bulkData, onProgress);
  }

  /**
   * 新しいスライドを追加（従来版、互換性のため保持）
   */
  public async addSlide(title: string, content: string): Promise<void> {
    const slideData: SlideContent = {
      title,
      content: content.split('\n• ').filter(item => item.trim() !== ''),
      slideType: 'content'
    };
    return this.slideCreator.createSingleSlide(slideData);
  }

  /**
   * 指定したスライドのコンテンツを更新
   */
  public async updateSlide(slideIndex: number, title: string, content: string): Promise<void> {
    return this.slideManager.updateSlide(slideIndex, title, content);
  }

  /**
   * 指定したスライドを削除
   */
  public async deleteSlide(slideIndex: number): Promise<void> {
    return this.slideManager.deleteSlide(slideIndex);
  }

  /**
   * スライド間にトランジションを追加
   */
  public async addTransitions(transitionType: 'fade' | 'slide' | 'none' = 'fade'): Promise<void> {
    return this.slideManager.addTransitions(transitionType);
  }

  /**
   * プレゼンテーション全体の統計を取得
   */
  public async getPresentationStats(): Promise<{
    slideCount: number;
    estimatedDuration: number;
    wordCount: number;
  }> {
    return this.presentationAnalyzer.getPresentationStats();
  }

  /**
   * テキストボックスを追加（テスト用）
   */
  public async addTextBox(text: string): Promise<void> {
    return this.slideCreator.addTextBox(text);
  }
}