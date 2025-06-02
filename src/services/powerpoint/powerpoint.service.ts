// src/services/powerpoint/powerpoint.service.ts - メインサービス（テーマ対応版）
/* global PowerPoint */

import { SlideCreator } from './slide-creator.service';
import { SlideManager } from './slide-manager.service';
import { PresentationAnalyzer } from './presentation-analyzer.service';
import { ThemeService } from './theme.service';
import { SmartContentPlacerService } from './smart-content-placer.service';
import { SlideInfo, SlideGenerationOptions, SlideContent, BulkSlideData } from './types';
import { ThemeInfo } from './theme.service';

/**
 * PowerPoint操作のメインサービスクラス（テーマ対応版）
 * 各専門サービスを統合し、統一されたAPIを提供
 */
export class PowerPointService {
  private slideCreator: SlideCreator;
  private slideManager: SlideManager;
  private presentationAnalyzer: PresentationAnalyzer;
  private themeService: ThemeService;
  private smartContentPlacer: SmartContentPlacerService;

  constructor() {
    this.themeService = new ThemeService();
    this.smartContentPlacer = new SmartContentPlacerService(this.themeService);
    this.slideCreator = new SlideCreator();
    this.slideManager = new SlideManager();
    this.presentationAnalyzer = new PresentationAnalyzer();
  }

  /**
   * 現在のプレゼンテーションのテーマ情報を取得
   */
  public async getThemeInfo(): Promise<ThemeInfo> {
    return this.themeService.getCurrentThemeInfo();
  }

  /**
   * 現在のプレゼンテーションの全スライド情報を取得
   */
  public async getAllSlides(): Promise<SlideInfo[]> {
    return this.presentationAnalyzer.getAllSlides();
  }

  /**
   * 複数のスライドを一括生成（テーマ対応版）
   */
  public async generateBulkSlides(
    bulkData: BulkSlideData, 
    onProgress?: (current: number, total: number, slideName: string) => void
  ): Promise<void> {
    // テーマ情報を事前に取得してキャッシュ
    await this.themeService.getCurrentThemeInfo();
    
    // テーマ対応はSlideCreator内で自動的に行われるため、
    // 追加のフラグは不要
    return this.slideCreator.generateBulkSlides(bulkData, onProgress);
  }

  /**
   * 新しいスライドを追加（テーマ対応版、従来版との互換性保持）
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

  /**
   * スマートコンテンツ配置のテスト
   */
  public async testSmartContentPlacement(slideData: SlideContent): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          // 新しいスライドを追加
          context.presentation.slides.add();
          await context.sync();
          
          // 最後のスライドを取得
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const slide = slides.items[slides.items.length - 1];
          
          // スマートコンテンツ配置を使用
          await this.smartContentPlacer.placeContent(context, slide, slideData, {
            theme: 'light',
            fontSize: 'medium'
          });
          
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * プレゼンテーションのテーマを適用
   */
  public async applyTheme(themeName: string): Promise<void> {
    console.log(`テーマ「${themeName}」の適用を試みています...`);
    // PowerPoint.js APIの制限により、現在は限定的な実装
    // 将来的にAPIが拡張された場合、実際のテーマ適用を実装
  }

  /**
   * カスタムレイアウトでスライドを作成
   */
  public async createSlideWithCustomLayout(
    slideData: SlideContent,
    layoutName: string
  ): Promise<void> {
    const options: SlideGenerationOptions = {
      slideLayout: layoutName as 'title' | 'content' | 'twoContent' | 'comparison' | 'blank'
    };
    
    return this.slideCreator.createSingleSlide(slideData, options);
  }
}