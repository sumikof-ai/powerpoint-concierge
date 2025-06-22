// src/services/powerpoint/powerpoint.service.ts - SlideContentGenerator統合版
/* global PowerPoint, console */

import {
  SlideContent,
  SlideGenerationOptions,
  BulkSlideData,
  SlideInfo,
  PresentationStats,
} from "./types";
import { SlideFactory } from "./core/SlideFactory";
import { ContentRenderer } from "./core/ContentRenderer";
import { ThemeApplier } from "./core/ThemeApplier";
import { PresentationAnalyzer } from "./presentation-analyzer.service";
import { SlideContentGenerator } from "./core/SlideContentGenerator";
import { PresentationOutline } from "../../taskpane/components/types";
import { OpenAIService } from "../openai.service";
import { TemplateBasedGenerationService } from "./template/TemplateBasedGenerationService";
import { TemplateAdaptationService } from "./template/TemplateAdaptationService";
import {
  TemplateInfo,
  TemplateRecommendation,
  TemplateRegistrationRequest,
} from "./template-types";

/**
 * PowerPoint操作のメインサービスクラス（テンプレート統合版）
 * 各専門サービスを組み合わせて高レベルな操作を提供
 */
export class PowerPointService {
  private slideFactory: SlideFactory;
  private contentRenderer: ContentRenderer;
  private themeApplier: ThemeApplier;
  private presentationAnalyzer: PresentationAnalyzer;
  private templateBasedService: TemplateBasedGenerationService;
  private templateAdaptationService: TemplateAdaptationService;

  private defaultOptions: SlideGenerationOptions = {
    includeTransitions: false,
    slideLayout: "content",
    theme: "light",
    fontSize: "medium",
    useThemeAwareGeneration: true,
  };

  constructor() {
    this.slideFactory = new SlideFactory();
    this.contentRenderer = new ContentRenderer();
    this.themeApplier = new ThemeApplier();
    this.presentationAnalyzer = new PresentationAnalyzer();
    this.templateBasedService = new TemplateBasedGenerationService();
    this.templateAdaptationService = new TemplateAdaptationService();
  }

  /**
   * アウトラインからの詳細化スライド生成（新機能）
   */
  public async generateSlidesFromOutline(
    outline: PresentationOutline,
    openAIService: OpenAIService,
    options: SlideGenerationOptions = {},
    onProgress?: (current: number, total: number, slideName: string) => void
  ): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const mergedOptions = { ...this.defaultOptions, ...options };

          // SlideContentGeneratorを使用してコンテンツを詳細化
          const slideContentGenerator = new SlideContentGenerator(openAIService);

          // 詳細化進捗のコールバック
          const detailProgress = (current: number, total: number, slideName: string) => {
            if (onProgress) {
              onProgress(current, total, `📝 ${slideName} の詳細化中...`);
            }
          };

          // エラーハンドリング付きで詳細化を実行
          const detailedSlides = await slideContentGenerator.generateWithErrorHandling(
            outline,
            mergedOptions,
            detailProgress,
            (slideIndex, error) => {
              console.warn(`スライド ${slideIndex + 1} の詳細化でエラー:`, error.message);
              if (onProgress) {
                onProgress(
                  slideIndex + 1,
                  outline.slides.length,
                  `⚠️ スライド ${slideIndex + 1} フォールバック処理`
                );
              }
            }
          );

          // PowerPointスライド生成進捗のコールバック
          const slideProgress = (current: number, total: number, slideName: string) => {
            if (onProgress) {
              onProgress(current, total, `🎨 ${slideName} のスライド作成中...`);
            }
          };

          // 詳細化されたコンテンツでスライドを作成
          await this.slideFactory.createBulkSlides(
            context,
            detailedSlides,
            mergedOptions,
            slideProgress
          );

          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * 複数のスライドを一括生成（従来機能）
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

          // スライド品質チェック
          const validationResults = this.validateSlidesBeforeGeneration(slides);
          if (validationResults.hasErrors) {
            console.warn("スライド品質チェックで警告が検出されました:", validationResults.warnings);
          }

          // スライドファクトリーを使用して一括生成
          await this.slideFactory.createBulkSlides(context, slides, mergedOptions, onProgress);

          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * 単一のスライドを作成（文字列パラメータ版）
   */
  public async addSlide(title: string, content: string): Promise<void> {
    const slideData: SlideContent = {
      title,
      content: this.parseContentString(content),
      slideType: "content",
    };

    const bulkData: BulkSlideData = {
      slides: [slideData],
      options: this.defaultOptions,
    };

    return this.generateBulkSlides(bulkData);
  }

  /**
   * 単一のスライドを作成（SlideContentオブジェクト版）
   */
  public async addSlideFromContent(
    slideContent: SlideContent,
    options: SlideGenerationOptions = {}
  ): Promise<void> {
    const bulkData: BulkSlideData = {
      slides: [slideContent],
      options: { ...this.defaultOptions, ...options },
    };

    return this.generateBulkSlides(bulkData);
  }

  /**
   * 現在のプレゼンテーションの全スライド情報を取得
   */
  public async getAllSlides(): Promise<SlideInfo[]> {
    return this.presentationAnalyzer.getAllSlides();
  }

  /**
   * プレゼンテーション統計を取得
   */
  public async getPresentationStats(): Promise<PresentationStats> {
    return this.presentationAnalyzer.getPresentationStats();
  }

  /**
   * 指定したスライドのコンテンツを更新（SlideManager機能を統合）
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
            slideType: "content",
          };

          // ContentRendererを使用してスライドをレンダリング
          await this.contentRenderer.renderContentSlide(
            context,
            slide,
            slideData,
            this.defaultOptions
          );

          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * 指定したスライドを削除（SlideManager機能を統合）
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
   * スライド間にトランジションを追加（SlideManager機能を統合）
   * 注意: PowerPoint.js では現在トランジション機能のサポートが限定的
   */
  public async addTransitions(transitionType: "fade" | "slide" | "none" = "fade"): Promise<void> {
    return new Promise((resolve) => {
      console.log(`トランジション設定をリクエストしました: ${transitionType}`);
      console.log("注意: PowerPoint.js では現在トランジション機能のサポートが限定的です");
      resolve();
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
            height: 200,
          });

          await context.sync();

          // デフォルトスタイルを適用
          textBox.textFrame.textRange.font.size = 14;
          this.themeApplier.applyThemeColors(textBox, "light", "body");

          await context.sync();
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * テーマテスト機能
   */
  public async testThemeApplication(): Promise<void> {
    const testSlide: SlideContent = {
      title: "テーマテスト",
      content: ["ライトテーマのテスト", "テキストの色とスタイル", "アクセントカラーの適用"],
      slideType: "content",
    };

    // 各テーマで同じスライドを作成
    for (const themeName of ["light", "dark", "colorful"] as const) {
      const bulkData: BulkSlideData = {
        slides: [
          {
            ...testSlide,
            title: `${testSlide.title} - ${themeName.toUpperCase()}テーマ`,
          },
        ],
        options: {
          ...this.defaultOptions,
          theme: themeName,
        },
      };

      await this.generateBulkSlides(bulkData);
    }
  }

  /**
   * アウトライン詳細化のテスト機能
   */
  public async testDetailedGeneration(openAIService: OpenAIService): Promise<void> {
    const testOutline: PresentationOutline = {
      title: "詳細化テスト用プレゼンテーション",
      estimatedDuration: 15,
      slides: [
        {
          slideNumber: 1,
          title: "テスト概要",
          content: ["目的", "範囲", "期待効果"],
          slideType: "title",
        },
        {
          slideNumber: 2,
          title: "現状分析",
          content: ["課題", "機会", "制約条件"],
          slideType: "content",
        },
        {
          slideNumber: 3,
          title: "まとめ",
          content: ["要点", "次ステップ", "アクション"],
          slideType: "conclusion",
        },
      ],
    };

    await this.generateSlidesFromOutline(
      testOutline,
      openAIService,
      { theme: "light", fontSize: "medium" },
      (current, total, status) => {
        console.log(`詳細化テスト進捗: ${current}/${total} - ${status}`);
      }
    );
  }

  /**
   * スライドのテキストボックスをクリア（SlideManager機能を統合）
   */
  private async clearSlideTextBoxes(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide
  ): Promise<void> {
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
   * スライド品質を検証
   */
  private validateSlidesBeforeGeneration(slides: SlideContent[]): {
    hasErrors: boolean;
    warnings: string[];
    suggestions: string[];
  } {
    const warnings: string[] = [];
    const suggestions: string[] = [];

    slides.forEach((slide, index) => {
      const validation = this.slideFactory.validateSlideContent(slide);

      if (!validation.isValid) {
        warnings.push(`スライド ${index + 1}: ${validation.warnings.join(", ")}`);
      }

      suggestions.push(...validation.suggestions);
    });

    return {
      hasErrors: warnings.length > 0,
      warnings,
      suggestions,
    };
  }

  /**
   * コンテンツ文字列をパース
   */
  private parseContentString(content: string): string[] {
    return content
      .split("\n• ")
      .map((item) => item.replace(/^• /, "").trim())
      .filter((item) => item !== "");
  }

  /**
   * スライド生成オプションを最適化
   */
  public optimizeGenerationOptions(
    slides: SlideContent[],
    userPreferences?: Partial<SlideGenerationOptions>
  ): SlideGenerationOptions {
    const optimized = { ...this.defaultOptions, ...userPreferences };

    // スライド数に基づく最適化
    if (slides.length > 10) {
      optimized.fontSize = "medium"; // 大量のスライドは中サイズで統一
    }

    // コンテンツ量に基づく最適化
    const totalContentLength = slides.reduce(
      (sum, slide) => sum + slide.title.length + slide.content.join("").length,
      0
    );

    if (totalContentLength > 2000) {
      optimized.slideLayout = "twoContent"; // 大量のコンテンツは2カラムを推奨
    }

    // スライドタイプの分布に基づく最適化
    const hasMultipleTitles = slides.filter((s) => s.slideType === "title").length > 1;
    if (hasMultipleTitles) {
      optimized.includeTransitions = true; // セクション区切りがある場合はトランジション追加
    }

    return optimized;
  }

  /**
   * プレゼンテーション品質レポートを生成
   */
  public async generateQualityReport(): Promise<{
    slideCount: number;
    estimatedDuration: number;
    contentAnalysis: string[];
    recommendations: string[];
  }> {
    const stats = await this.getPresentationStats();

    const contentAnalysis: string[] = [];
    const recommendations: string[] = [];

    // スライド数の分析
    if (stats.slideCount > 20) {
      contentAnalysis.push("スライド数が多め（20枚超）");
      recommendations.push("内容を整理して15枚以内に収めることを推奨");
    }

    // 文字数の分析
    if (stats.wordCount > 1000) {
      contentAnalysis.push("テキスト量が多め");
      recommendations.push("図表やビジュアル要素の活用を検討");
    }

    // 予想時間の分析
    if (stats.estimatedDuration > 30) {
      contentAnalysis.push("発表時間が長め（30分超）");
      recommendations.push("聴衆の集中力を維持するため、適度な休憩を検討");
    }

    return {
      slideCount: stats.slideCount,
      estimatedDuration: stats.estimatedDuration,
      contentAnalysis,
      recommendations,
    };
  }

  /**
   * エクスポート形式の提案
   */
  public suggestExportFormats(presentationType: "presentation" | "handout" | "notes"): string[] {
    const suggestions: string[] = [];

    switch (presentationType) {
      case "presentation":
        suggestions.push("フルスクリーン表示用のPPTX形式");
        suggestions.push("PDF形式（配布用）");
        break;
      case "handout":
        suggestions.push("6スライド/ページのPDF形式");
        suggestions.push("ノート付きPDF形式");
        break;
      case "notes":
        suggestions.push("スピーカーノート付きPDF形式");
        suggestions.push("DOCX形式（編集用）");
        break;
    }

    return suggestions;
  }

  // ==================== テンプレート統合機能 ====================

  /**
   * テンプレートベースのスライド生成
   */
  public async generateSlidesWithTemplate(
    userInput: string,
    outline: PresentationOutline,
    openAIService: OpenAIService,
    options: SlideGenerationOptions = {},
    onProgress?: (phase: string, current: number, total: number, message: string) => void
  ): Promise<void> {
    try {
      // 1. テンプレート選択
      if (onProgress) {
        onProgress("template-selection", 1, 4, "テンプレートを選択中...");
      }

      const recommendations = await this.templateBasedService.selectOptimalTemplate(userInput);

      if (recommendations.length === 0) {
        // テンプレートが見つからない場合は従来の方法で生成
        await this.generateSlidesFromOutline(
          outline,
          openAIService,
          options,
          onProgress
            ? (current, total, slideName) => {
                onProgress("creating", current, total, `スライド作成中: ${slideName}`);
              }
            : undefined
        );
        return;
      }

      const selectedTemplate = recommendations[0].template;

      // 2. アウトライン適応
      if (onProgress) {
        onProgress("outline-adaptation", 2, 4, "アウトラインをテンプレートに適応中...");
      }

      const adaptedOutline = await this.templateBasedService.adaptOutlineToTemplate(
        outline,
        selectedTemplate.id
      );

      // 3. テンプレート最適化コンテンツ生成
      if (onProgress) {
        onProgress("content-generation", 3, 4, "テンプレート最適化コンテンツを生成中...");
      }

      const optimizedSlides = await this.templateBasedService.generateTemplateOptimizedContent(
        adaptedOutline,
        openAIService,
        (current, total, message) => {
          if (onProgress) {
            onProgress("content-generation", current, total, message);
          }
        }
      );

      // 4. PowerPointスライド作成
      if (onProgress) {
        onProgress("slide-creation", 4, 4, "PowerPointスライドを作成中...");
      }

      await this.createSlidesFromOptimizedContent(optimizedSlides, selectedTemplate, options);
    } catch (error) {
      console.error("Template-based slide generation failed:", error);
      // フォールバック: 通常の生成方法
      await this.generateSlidesFromOutline(
        outline,
        openAIService,
        options,
        onProgress
          ? (current, total, slideName) => {
              onProgress("creating", current, total, `スライド作成中: ${slideName}`);
            }
          : undefined
      );
    }
  }

  /**
   * 現在のプレゼンテーションをテンプレートとして登録
   */
  public async registerCurrentPresentationAsTemplate(
    templateName: string,
    description: string,
    metadata: Partial<TemplateRegistrationRequest["metadata"]>
  ): Promise<TemplateInfo> {
    try {
      const registrationRequest: TemplateRegistrationRequest = {
        metadata: {
          presentationStyle: metadata.presentationStyle || "formal",
          targetAudience: metadata.targetAudience || "general",
          purpose: metadata.purpose || "report",
          tags: [templateName.toLowerCase(), "user-created", ...(metadata.tags || [])],
          ...metadata,
        },
        autoAnalyze: true,
      };

      const template = await this.templateBasedService.registerTemplate(registrationRequest);

      // 登録されたテンプレートの情報を更新
      template.name = templateName;
      template.description = description;

      return template;
    } catch (error) {
      console.error("Template registration failed:", error);
      throw new Error(`テンプレートの登録に失敗しました: ${error.message}`);
    }
  }

  /**
   * テンプレート推奨を取得
   */
  public async getTemplateRecommendations(
    userInput: string,
    preferences?: {
      categories?: string[];
      maxResults?: number;
    }
  ): Promise<TemplateRecommendation[]> {
    try {
      return await this.templateBasedService.selectOptimalTemplate(userInput, {
        preferences: {
          maxResults: preferences?.maxResults || 5,
          minimumScore: 0.3,
        },
      });
    } catch (error) {
      console.error("Template recommendation failed:", error);
      return [];
    }
  }

  /**
   * テンプレートライブラリの管理
   */
  public getTemplateLibrary() {
    return this.templateBasedService.getTemplateLibrary();
  }

  public getPopularTemplates(limit: number = 10) {
    return this.templateBasedService.getPopularTemplates(limit);
  }

  public getRecentTemplates(limit: number = 5) {
    return this.templateBasedService.getRecentTemplates(limit);
  }

  public searchTemplates(query: string, filters?: any) {
    return this.templateBasedService.searchTemplates(query, filters);
  }

  /**
   * テンプレートフィードバック
   */
  public recordTemplateFeedback(
    templateId: string,
    rating: number,
    feedback?: string,
    success?: boolean
  ): void {
    this.templateBasedService.recordTemplateFeedback(templateId, rating, feedback, success);
  }

  /**
   * 現在のプレゼンテーションテンプレート検出
   */
  public async detectCurrentTemplate(): Promise<TemplateInfo | null> {
    try {
      return await this.templateAdaptationService.detectTemplate();
    } catch (error) {
      console.error("Template detection failed:", error);
      return null;
    }
  }

  private async createSlidesFromOptimizedContent(
    optimizedSlides: any[],
    template: TemplateInfo,
    options: SlideGenerationOptions
  ): Promise<void> {
    return PowerPoint.run(async (context) => {
      for (let i = 0; i < optimizedSlides.length; i++) {
        const slideContent = optimizedSlides[i];

        try {
          // テンプレート情報を考慮したスライド作成
          const mergedOptions = {
            ...options,
            slideLayout: slideContent.layoutSuggestion || options.slideLayout,
            templateStyle: template.metadata.presentationStyle,
            templateDensity: template.metadata.contentDensity,
          };

          const slideData: SlideContent = {
            title: slideContent.title,
            content: Array.isArray(slideContent.content)
              ? slideContent.content
              : [slideContent.content],
            slideType: slideContent.slideType || "content",
            speakerNotes: slideContent.speakerNotes || "",
          };

          await this.addSlideFromContent(slideData, mergedOptions);
        } catch (error) {
          console.error(`Failed to create slide ${i + 1} from optimized content:`, error);

          // フォールバック: 基本的なスライド作成
          const fallbackSlideData: SlideContent = {
            title: slideContent.title || `スライド ${i + 1}`,
            content: ["テンプレート最適化でエラーが発生しました"],
            slideType: "content",
          };

          await this.addSlideFromContent(fallbackSlideData, options);
        }
      }

      await context.sync();
    });
  }

  /**
   * アウトライン詳細化の進捗管理
   */
  public async generateSlidesWithDetailedProgress(
    outline: PresentationOutline,
    openAIService: OpenAIService,
    options: SlideGenerationOptions = {},
    onDetailProgress?: (
      phase: "analyzing" | "detailing" | "creating",
      current: number,
      total: number,
      message: string
    ) => void
  ): Promise<void> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const mergedOptions = { ...this.defaultOptions, ...options };

          // Phase 1: アウトライン分析
          if (onDetailProgress) {
            onDetailProgress("analyzing", 1, 3, "アウトライン構造を分析中...");
          }

          const slideContentGenerator = new SlideContentGenerator(openAIService);

          // Phase 2: 詳細化
          if (onDetailProgress) {
            onDetailProgress("detailing", 2, 3, "スライドコンテンツを詳細化中...");
          }

          const detailedSlides = await slideContentGenerator.generateWithErrorHandling(
            outline,
            mergedOptions,
            (current, total, slideName) => {
              if (onDetailProgress) {
                onDetailProgress("detailing", current, total, `📝 ${slideName} を詳細化中...`);
              }
            }
          );

          // Phase 3: PowerPoint作成
          if (onDetailProgress) {
            onDetailProgress("creating", 3, 3, "PowerPointスライドを作成中...");
          }

          await this.slideFactory.createBulkSlides(
            context,
            detailedSlides,
            mergedOptions,
            (current, total, slideName) => {
              if (onDetailProgress) {
                onDetailProgress("creating", current, total, `🎨 ${slideName} を作成中...`);
              }
            }
          );

          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }
}
