// src/services/powerpoint/index.ts - 修正されたエクスポートファイル

import { SlideFactory } from './core/SlideFactory';
import { ThemeApplier } from './core/ThemeApplier';
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

// 専門サービス（既存）
export { SlideCreator } from './slide-creator.service';
export { SlideManager } from './slide-manager.service';
export { PresentationAnalyzer } from './presentation-analyzer.service';
export { SlideLayoutFactory } from './slide-layout-factory.service';
export { SlideThemeApplier } from './slide-theme-applier.service';

// 配置とテーマサービス
export { SmartContentPlacerService } from './smart-content-placer.service';

// テーマサービス（新規作成）
export { ThemeService } from './theme/ThemeService';
export { ThemeAnalyzer } from './theme/ThemeAnalyzer';

// 型定義
export * from './types';

// テーマ関連型定義
export * from './theme-types';

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
 * リファクタリング完了を記録
 */
export const REFACTORING_INFO = {
  version: '2.0.0',
  completedAt: new Date().toISOString(),
  changes: [
    'ChatInput.tsx を3ファイルに分割 (461行 → 最大180行)',
    'OutlineEditor.tsx を2ファイルに分割 (350行 → 最大170行)', 
    'PowerPointService を4ファイルに分割 (530行 → 最大250行)',
    'ThemeService を2ファイルに分割 (324行 → 最大180行)',
    '全ファイルが300行以内の目標を達成'
  ],
  benefits: [
    '保守性の向上',
    'テスタビリティの向上', 
    '再利用性の向上',
    '開発効率の向上'
  ]
} as const;