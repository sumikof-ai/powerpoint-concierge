// TemplateIntegrationTest.ts
// テンプレート統合システムのテスト・デモンストレーション
/* global console, performance */

import { TemplateBasedGenerationService } from "./TemplateBasedGenerationService";
import { TemplateAdaptationService } from "./TemplateAdaptationService";
import { PowerPointService } from "../powerpoint.service";
import { OpenAIService } from "../../openai.service";

export class TemplateIntegrationTest {
  private templateService: TemplateBasedGenerationService;
  private adaptationService: TemplateAdaptationService;
  private powerPointService: PowerPointService;

  constructor() {
    this.templateService = new TemplateBasedGenerationService();
    this.adaptationService = new TemplateAdaptationService();
    this.powerPointService = new PowerPointService();
  }

  /**
   * テンプレート統合の完全テスト
   */
  async runFullIntegrationTest(openAIService: OpenAIService): Promise<{
    success: boolean;
    results: any[];
    errors: string[];
  }> {
    console.log("🧪 テンプレート統合テストを開始...");

    const results: any[] = [];
    const errors: string[] = [];

    try {
      // 1. テンプレートライブラリのテスト
      console.log("📚 テンプレートライブラリをテスト中...");
      const libraryTest = await this.testTemplateLibrary();
      results.push(libraryTest);

      // 2. テンプレート推奨システムのテスト
      console.log("🎯 テンプレート推奨システムをテスト中...");
      const recommendationTest = await this.testTemplateRecommendation();
      results.push(recommendationTest);

      // 3. テンプレート適応のテスト
      console.log("🔄 テンプレート適応をテスト中...");
      const adaptationTest = await this.testTemplateAdaptation();
      results.push(adaptationTest);

      // 4. PowerPoint統合のテスト
      console.log("🎨 PowerPoint統合をテスト中...");
      const powerPointTest = await this.testPowerPointIntegration();
      results.push(powerPointTest);

      // 5. エンドツーエンドテスト
      console.log("🚀 エンドツーエンドテストを実行中...");
      const e2eTest = await this.testEndToEndWorkflow(openAIService);
      results.push(e2eTest);

      console.log("✅ テンプレート統合テストが完了しました");

      return {
        success: true,
        results,
        errors,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "不明なエラー";
      console.error("❌ テンプレート統合テストでエラー:", errorMessage);
      errors.push(errorMessage);

      return {
        success: false,
        results,
        errors,
      };
    }
  }

  /**
   * テンプレートライブラリの基本機能テスト
   */
  private async testTemplateLibrary(): Promise<any> {
    try {
      const library = this.templateService.getTemplateLibrary();

      const test = {
        name: "Template Library Test",
        totalTemplates: library.templates.length,
        categories: Object.keys(library.categories),
        popularTemplates: this.templateService.getPopularTemplates(3),
        recentTemplates: this.templateService.getRecentTemplates(2),
        success: true,
      };

      console.log(`  📊 ${test.totalTemplates}個のテンプレートが利用可能`);
      console.log(`  📁 ${test.categories.length}個のカテゴリ: ${test.categories.join(", ")}`);

      return test;
    } catch (error) {
      return {
        name: "Template Library Test",
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * テンプレート推奨システムのテスト
   */
  private async testTemplateRecommendation(): Promise<any> {
    try {
      const testInputs = [
        "営業戦略についてのプレゼンテーションを作成してください",
        "学術研究の成果発表資料を作りたいです",
        "シンプルな企業紹介の資料を作成したい",
        "技術的な製品説明のプレゼンテーションが必要です",
      ];

      const recommendationResults = [];

      for (const input of testInputs) {
        const recommendations = await this.templateService.selectOptimalTemplate(input);
        recommendationResults.push({
          input: input.substring(0, 20) + "...",
          recommendationCount: recommendations.length,
          topScore: recommendations.length > 0 ? recommendations[0].score : 0,
          topTemplate: recommendations.length > 0 ? recommendations[0].template.name : "なし",
        });
      }

      console.log(`  🎯 ${testInputs.length}個の入力パターンをテスト`);
      recommendationResults.forEach((result) => {
        console.log(
          `    "${result.input}": ${result.recommendationCount}個の推奨, トップスコア: ${(result.topScore * 100).toFixed(1)}%`
        );
      });

      return {
        name: "Template Recommendation Test",
        testInputs: testInputs.length,
        results: recommendationResults,
        success: true,
      };
    } catch (error) {
      return {
        name: "Template Recommendation Test",
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * テンプレート適応機能のテスト
   */
  private async testTemplateAdaptation(): Promise<any> {
    try {
      // テスト用のアウトライン
      const testOutline = {
        title: "テスト用プレゼンテーション",
        audience: "一般",
        objective: "テスト",
        slides: [
          {
            slideNumber: 1,
            title: "タイトルスライド",
            contentType: "title",
            keyPoints: ["テスト項目1", "テスト項目2"],
            detailLevel: "basic",
          },
          {
            slideNumber: 2,
            title: "コンテンツスライド",
            contentType: "content",
            keyPoints: ["内容1", "内容2", "内容3"],
            detailLevel: "detailed",
          },
        ],
        estimatedDuration: 10,
      };

      // デフォルトテンプレートを取得
      const library = this.templateService.getTemplateLibrary();
      const testTemplate = library.templates[0];

      if (!testTemplate) {
        throw new Error("テスト用テンプレートが見つかりません");
      }

      // アウトライン適応をテスト
      const adaptedOutline = await this.templateService.adaptOutlineToTemplate(
        testOutline,
        testTemplate.id
      );

      console.log(`  🔄 アウトライン適応完了: ${adaptedOutline.adaptedSlides.length}スライド`);
      console.log(`  📊 適応信頼度: ${(adaptedOutline.confidence * 100).toFixed(1)}%`);
      console.log(`  🔧 適応項目: ${adaptedOutline.adaptations.length}個`);

      return {
        name: "Template Adaptation Test",
        originalSlides: testOutline.slides.length,
        adaptedSlides: adaptedOutline.adaptedSlides.length,
        adaptations: adaptedOutline.adaptations.length,
        confidence: adaptedOutline.confidence,
        templateUsed: testTemplate.name,
        success: true,
      };
    } catch (error) {
      return {
        name: "Template Adaptation Test",
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * PowerPoint統合のテスト
   */
  private async testPowerPointIntegration(): Promise<any> {
    try {
      // テンプレート検出のテスト
      const detectedTemplate = await this.powerPointService.detectCurrentTemplate();

      // テンプレート推奨のテスト
      const recommendations =
        await this.powerPointService.getTemplateRecommendations("ビジネス提案書を作成したいです");

      console.log(`  🔍 テンプレート検出: ${detectedTemplate ? "成功" : "失敗"}`);
      console.log(`  💡 推奨テンプレート: ${recommendations.length}個`);

      return {
        name: "PowerPoint Integration Test",
        templateDetected: !!detectedTemplate,
        detectedTemplateName: detectedTemplate?.name || "なし",
        recommendationsCount: recommendations.length,
        topRecommendation: recommendations.length > 0 ? recommendations[0].template.name : "なし",
        success: true,
      };
    } catch (error) {
      return {
        name: "PowerPoint Integration Test",
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * エンドツーエンドワークフローのテスト
   */
  private async testEndToEndWorkflow(
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    _openAIService: OpenAIService
  ): Promise<any> {
    try {
      console.log("  🚀 完全なテンプレートワークフローをテスト中...");

      const userInput = "マーケティング戦略についてのプレゼンテーションを作成してください";

      // 1. テンプレート推奨
      const recommendations = await this.templateService.selectOptimalTemplate(userInput);
      if (recommendations.length === 0) {
        throw new Error("テンプレート推奨が見つかりませんでした");
      }

      const selectedTemplate = recommendations[0].template;
      console.log(`    ✓ テンプレート選択: ${selectedTemplate.name}`);

      // 2. テスト用アウトライン
      const testOutline = {
        title: "マーケティング戦略プレゼンテーション",
        audience: "経営陣",
        objective: "新しいマーケティング戦略の提案",
        slides: [
          {
            slideNumber: 1,
            title: "マーケティング戦略概要",
            contentType: "title",
            keyPoints: ["戦略の概要", "目標設定"],
            detailLevel: "basic",
          },
          {
            slideNumber: 2,
            title: "市場分析",
            contentType: "content",
            keyPoints: ["市場トレンド", "競合分析", "機会と脅威"],
            detailLevel: "detailed",
          },
        ],
        estimatedDuration: 15,
      };

      // 3. アウトライン適応
      const adaptedOutline = await this.templateService.adaptOutlineToTemplate(
        testOutline,
        selectedTemplate.id
      );
      console.log(
        `    ✓ アウトライン適応完了: 信頼度 ${(adaptedOutline.confidence * 100).toFixed(1)}%`
      );

      // 4. コンテンツ最適化（モックバージョン）
      const optimizedSlides = await this.mockTemplateOptimizedGeneration(adaptedOutline);
      console.log(`    ✓ コンテンツ最適化完了: ${optimizedSlides.length}スライド`);

      return {
        name: "End-to-End Workflow Test",
        userInput,
        recommendationsFound: recommendations.length,
        selectedTemplate: selectedTemplate.name,
        adaptationConfidence: adaptedOutline.confidence,
        optimizedSlidesCount: optimizedSlides.length,
        workflowSteps: [
          "テンプレート推奨",
          "テンプレート選択",
          "アウトライン適応",
          "コンテンツ最適化",
        ],
        success: true,
      };
    } catch (error) {
      return {
        name: "End-to-End Workflow Test",
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * テンプレート最適化コンテンツ生成のモック
   */
  private async mockTemplateOptimizedGeneration(adaptedOutline: any): Promise<any[]> {
    // 実際のAI生成をモック
    return adaptedOutline.adaptedSlides.map((slide: any) => ({
      title: slide.adaptedContent.title,
      content: slide.adaptedContent.content,
      slideType: slide.adaptedContent.slideType,
      templateOptimizations: [
        "テンプレートスタイルに調整",
        "対象聴衆に最適化",
        "レイアウト最適化適用",
      ],
      appliedTemplate: adaptedOutline.selectedTemplate.name,
    }));
  }

  /**
   * テンプレート機能のパフォーマンステスト
   */
  async runPerformanceTest(): Promise<any> {
    console.log("⚡ テンプレート機能のパフォーマンステストを開始...");

    const performanceResults = {
      templateLibraryLoad: 0,
      templateRecommendation: 0,
      templateAdaptation: 0,
      totalOperations: 0,
    };

    try {
      // テンプレートライブラリ読み込み時間
      const libraryStart = performance.now();
      this.templateService.getTemplateLibrary();
      performanceResults.templateLibraryLoad = performance.now() - libraryStart;

      // テンプレート推奨時間
      const recommendationStart = performance.now();
      await this.templateService.selectOptimalTemplate("テスト用プレゼンテーション");
      performanceResults.templateRecommendation = performance.now() - recommendationStart;

      // テンプレート適応時間
      const adaptationStart = performance.now();
      const library = this.templateService.getTemplateLibrary();
      if (library.templates.length > 0) {
        await this.templateService.adaptOutlineToTemplate(
          { title: "テスト", slides: [] },
          library.templates[0].id
        );
      }
      performanceResults.templateAdaptation = performance.now() - adaptationStart;

      performanceResults.totalOperations =
        performanceResults.templateLibraryLoad +
        performanceResults.templateRecommendation +
        performanceResults.templateAdaptation;

      console.log("⚡ パフォーマンステスト結果:");
      console.log(
        `  📚 ライブラリ読み込み: ${performanceResults.templateLibraryLoad.toFixed(2)}ms`
      );
      console.log(
        `  🎯 テンプレート推奨: ${performanceResults.templateRecommendation.toFixed(2)}ms`
      );
      console.log(`  🔄 テンプレート適応: ${performanceResults.templateAdaptation.toFixed(2)}ms`);
      console.log(`  🏁 総実行時間: ${performanceResults.totalOperations.toFixed(2)}ms`);

      return {
        name: "Performance Test",
        results: performanceResults,
        success: true,
        benchmark: {
          acceptable: performanceResults.totalOperations < 5000, // 5秒以内
          fast: performanceResults.totalOperations < 1000, // 1秒以内
        },
      };
    } catch (error) {
      return {
        name: "Performance Test",
        success: false,
        error: error.message,
        partialResults: performanceResults,
      };
    }
  }

  /**
   * デバッグ情報の出力
   */
  printDebugInfo(): void {
    console.log("🔍 テンプレート統合システム デバッグ情報:");

    const library = this.templateService.getTemplateLibrary();
    console.log(`📊 統計情報:`);
    console.log(`  - 総テンプレート数: ${library.statistics.totalTemplates}`);
    console.log(`  - カテゴリ別分布:`, library.statistics.byCategory);
    console.log(`  - 最近追加: ${library.statistics.recentlyAdded.length}個`);

    console.log(`📁 利用可能カテゴリ:`);
    Object.entries(library.categories).forEach(([category, templates]) => {
      console.log(`  - ${category}: ${templates.length}個`);
    });

    console.log(`🔧 システム情報:`);
    console.log(`  - TemplateBasedGenerationService: 初期化済み`);
    console.log(`  - TemplateAdaptationService: 初期化済み`);
    console.log(`  - PowerPointService: 初期化済み`);
  }
}
