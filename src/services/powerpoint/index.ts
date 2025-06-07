// src/services/powerpoint/index.ts - SlideContentGenerator統合版

import { SlideFactory } from './core/SlideFactory';
import { ThemeApplier } from './core/ThemeApplier';
import { SlideContentGenerator } from './core/SlideContentGenerator';
import { PowerPointService } from './powerpoint.service';
import { PresentationAnalyzer } from './presentation-analyzer.service';
import { ThemeAnalyzer } from './theme/ThemeAnalyzer';
import { ThemeService } from './theme/ThemeService';

// メインサービス（PowerPointService.ts から）
export { PowerPointService } from './powerpoint.service';

// コアサービス（分割後）
export { SlideFactory } from './core/SlideFactory';
export { ContentRenderer } from './core/ContentRenderer';
export { ThemeApplier } from './core/ThemeApplier';
export { SlideContentGenerator } from './core/SlideContentGenerator';

// 専門サービス（既存）
export { PresentationAnalyzer } from './presentation-analyzer.service';

// 配置とテーマサービス
export { SmartContentPlacerService } from './smart-content-placer.service';

// テーマサービス（新規作成）
export { ThemeService } from './theme/ThemeService';
export { ThemeAnalyzer } from './theme/ThemeAnalyzer';

// 型定義
export * from './types';

// テーマ関連型定義
export * from './theme-types';

// テンプレート関連型定義とサービス
export * from './template-types';
export { TemplateAdaptationService } from './template/TemplateAdaptationService';
export { TemplateBasedGenerationService } from './template/TemplateBasedGenerationService';
export { TemplatePatternExtractor } from './template/TemplatePatternExtractor';

/**
 * PowerPoint操作のファクトリー関数
 * 使用例: const pptService = createPowerPointService();
 */
export function createPowerPointService(): PowerPointService {
  return new PowerPointService();
}

/**
 * 軽量版PowerPointサービス（基本機能のみ）
 * メモリ使用量を抑えたい場合に使用
 */
export function createLightweightPowerPointService(): {
  slideFactory: SlideFactory;
  themeApplier: ThemeApplier;
} {
  return {
    slideFactory: new SlideFactory(),
    themeApplier: new ThemeApplier()
  };
}

/**
 * 詳細化機能付きPowerPointサービス（新機能）
 * SlideContentGeneratorを含む完全版サービス
 */
export function createEnhancedPowerPointService(openAIService?: any): {
  powerPointService: PowerPointService;
  slideContentGenerator: SlideContentGenerator | null;
} {
  const powerPointService = new PowerPointService();
  const slideContentGenerator = openAIService 
    ? new SlideContentGenerator(openAIService) 
    : null;

  return {
    powerPointService,
    slideContentGenerator
  };
}

/**
 * テーマ専用サービス（新版）
 * テーマ機能のみを使用する場合
 */
export function createThemeService(): {
  themeService: ThemeService;
  themeAnalyzer: ThemeAnalyzer;
} {
  return {
    themeService: new ThemeService(),
    themeAnalyzer: new ThemeAnalyzer()
  };
}

/**
 * 分析専用サービス
 * プレゼンテーション分析のみを行う場合
 */
export function createAnalysisService(): PresentationAnalyzer {
  return new PresentationAnalyzer();
}

/**
 * サービスの互換性チェック
 */
export function checkServiceCompatibility(): {
  isCompatible: boolean;
  warnings: string[];
} {
  const warnings: string[] = [];
  let isCompatible = true;

  // PowerPoint.js APIの存在チェック
  if (typeof PowerPoint === 'undefined') {
    warnings.push('PowerPoint.js APIが利用できません');
    isCompatible = false;
  }

  // 必要な機能の存在チェック
  try {
    if (PowerPoint && !PowerPoint.run) {
      warnings.push('PowerPoint.run メソッドが利用できません');
      isCompatible = false;
    }
  } catch (error) {
    warnings.push('PowerPoint API アクセスエラー');
    isCompatible = false;
  }

  return { isCompatible, warnings };
}

/**
 * 高度なPowerPointサービス
 * 全機能を含む包括的なサービス
 */
export function createAdvancedPowerPointService(): {
  powerPointService: PowerPointService;
  themeService: ThemeService;
  themeAnalyzer: ThemeAnalyzer;
  presentationAnalyzer: PresentationAnalyzer;
} {
  return {
    powerPointService: new PowerPointService(),
    themeService: new ThemeService(),
    themeAnalyzer: new ThemeAnalyzer(),
    presentationAnalyzer: new PresentationAnalyzer()
  };
}

/**
 * 詳細化機能のテスト用ファクトリー
 * 開発・テスト環境での使用を想定
 */
export function createTestSlideContentGenerator(
  openAIService: any,
  testMode: boolean = false
): SlideContentGenerator {
  const generator = new SlideContentGenerator(openAIService);
  
  if (testMode) {
    // テストモード用の設定があれば追加
    console.log('SlideContentGenerator をテストモードで初期化しました');
  }
  
  return generator;
}

/**
 * パフォーマンス監視付きサービス
 * 本番環境での使用を想定
 */
export function createMonitoredPowerPointService(): {
  service: PowerPointService;
  getPerformanceMetrics: () => any;
} {
  const service = new PowerPointService();
  const startTime = Date.now();
  let operationCount = 0;

  // 元のメソッドをラップして監視機能を追加
  const originalGenerateSlidesFromOutline = service.generateSlidesFromOutline.bind(service);
  
  service.generateSlidesFromOutline = async (...args) => {
    operationCount++;
    const opStartTime = Date.now();
    
    try {
      const result = await originalGenerateSlidesFromOutline(...args);
      const duration = Date.now() - opStartTime;
      console.log(`詳細化スライド生成完了: ${duration}ms`);
      return result;
    } catch (error) {
      console.error('詳細化スライド生成エラー:', error);
      throw error;
    }
  };

  return {
    service,
    getPerformanceMetrics: () => ({
      uptime: Date.now() - startTime,
      operationCount,
      averageOperationTime: operationCount > 0 ? (Date.now() - startTime) / operationCount : 0
    })
  };
}

/**
 * エラーハンドリング強化版サービス
 * エラー処理とログ機能を強化
 */
export function createRobustPowerPointService(
  errorHandler?: (error: Error, context: string) => void
): PowerPointService {
  const service = new PowerPointService();
  
  // エラーハンドリングの強化
  const originalMethods = [
    'generateSlidesFromOutline',
    'generateBulkSlides',
    'addSlide',
    'updateSlide',
    'deleteSlide'
  ];

  originalMethods.forEach(methodName => {
    const originalMethod = (service as any)[methodName];
    if (typeof originalMethod === 'function') {
      (service as any)[methodName] = async (...args: any[]) => {
        try {
          return await originalMethod.apply(service, args);
        } catch (error) {
          const errorContext = `PowerPointService.${methodName}`;
          console.error(`${errorContext} でエラー:`, error);
          
          if (errorHandler) {
            errorHandler(error instanceof Error ? error : new Error('不明なエラー'), errorContext);
          }
          
          throw error;
        }
      };
    }
  });

  return service;
}

/**
 * リファクタリング完了を記録（更新版）
 */
export const REFACTORING_INFO = {
  version: '3.0.0',
  completedAt: new Date().toISOString(),
  changes: [
    'SlideContentGenerator を新規作成（スライド毎の詳細化機能）',
    'PowerPointService に詳細化機能を統合',
    'ChatInput に詳細な進捗表示を追加',
    '3段階API呼び出し戦略の完全実装',
    'エラーハンドリングとフォールバック機能の強化',
    'スライドタイプ別最適化プロンプトの実装',
    'リアルタイム進捗管理とユーザー体験の向上'
  ],
  newFeatures: [
    '🔥 スライド毎の詳細化機能',
    '📊 段階的進捗表示（分析→詳細化→作成）',
    '🎯 スライドタイプ別最適化',
    '🔧 エラー時フォールバック',
    '📈 パフォーマンス監視',
    '🛡️ ロバストエラーハンドリング'
  ],
  benefits: [
    '説明資料として使える詳細なコンテンツ',
    '一貫性のある高品質なプレゼンテーション',
    '大幅な作業時間短縮',
    '聴衆の自立理解を促進'
  ]
} as const;

/**
 * 機能テスト用のヘルパー関数
 */
export async function testSlideContentGeneration(
  openAIService: any,
  testOutline?: any
): Promise<{
  success: boolean;
  results: any[];
  errors: string[];
}> {
  const results: any[] = [];
  const errors: string[] = [];

  try {
    const generator = new SlideContentGenerator(openAIService);
    const service = new PowerPointService();

    // デフォルトテストアウトライン
    const outline = testOutline || {
      title: "テスト用プレゼンテーション",
      estimatedDuration: 10,
      slides: [
        {
          slideNumber: 1,
          title: "テスト概要",
          content: ["目的", "範囲"],
          slideType: 'title'
        },
        {
          slideNumber: 2,
          title: "テスト詳細",
          content: ["内容1", "内容2"],
          slideType: 'content'
        }
      ]
    };

    // 詳細化テスト
    console.log('詳細化テストを開始...');
    const detailedSlides = await generator.generateDetailedSlides(
      outline,
      { theme: 'light', fontSize: 'medium' },
      (current, total, name) => {
        console.log(`進捗: ${current}/${total} - ${name}`);
      }
    );

    results.push({
      type: 'detailed_slides',
      count: detailedSlides.length,
      success: true
    });

    // PowerPoint生成テスト
    console.log('PowerPoint生成テストを開始...');
    await service.generateSlidesFromOutline(
      outline,
      openAIService,
      { theme: 'light', fontSize: 'medium' }
    );

    results.push({
      type: 'powerpoint_generation',
      success: true
    });

    return {
      success: true,
      results,
      errors
    };

  } catch (error) {
    errors.push(error instanceof Error ? error.message : '不明なエラー');
    return {
      success: false,
      results,
      errors
    };
  }
}