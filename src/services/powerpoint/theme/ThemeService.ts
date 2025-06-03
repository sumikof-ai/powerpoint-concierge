// src/services/powerpoint/theme/ThemeService.ts - リファクタリング版テーマサービス
/* global PowerPoint */

import { ThemeAnalyzer } from './ThemeAnalyzer';
import { 
  ThemeInfo, 
  ColorScheme, 
  FontScheme, 
  LayoutInfo, 
  PlaceholderInfo 
} from '../theme-types';

/**
 * PowerPointのテーマ情報を取得・管理するメインサービス
 */
export class ThemeService {
  private themeAnalyzer: ThemeAnalyzer;
  private cachedThemeInfo: ThemeInfo | null = null;

  constructor() {
    this.themeAnalyzer = new ThemeAnalyzer();
  }

  /**
   * 現在のプレゼンテーションのテーマ情報を取得
   */
  public async getCurrentThemeInfo(): Promise<ThemeInfo> {
    if (this.cachedThemeInfo) {
      return this.cachedThemeInfo;
    }

    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const presentation = context.presentation;
          
          // プレゼンテーションの基本情報を読み込み
          presentation.load("title");
          await context.sync();

          // 各種情報を並行して取得
          const [layouts, colorScheme, fontScheme] = await Promise.all([
            this.getAvailableLayouts(context),
            this.themeAnalyzer.analyzeColorScheme(context),
            this.themeAnalyzer.analyzeFontScheme(context)
          ]);

          const themeInfo: ThemeInfo = {
            name: await this.detectThemeName(context),
            colorScheme,
            fontScheme,
            availableLayouts: layouts
          };

          this.cachedThemeInfo = themeInfo;
          resolve(themeInfo);
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  /**
   * 利用可能なスライドレイアウトを取得
   */
  private async getAvailableLayouts(context: PowerPoint.RequestContext): Promise<LayoutInfo[]> {
    const commonLayouts: LayoutInfo[] = [
      {
        name: "Title Slide",
        type: 'title',
        placeholders: [
          {
            type: 'title',
            position: { x: 75, y: 150, width: 600, height: 150 },
            textFormat: {
              defaultFontSize: 44,
              defaultFontBold: true,
              defaultAlignment: 'center'
            }
          },
          {
            type: 'subtitle',
            position: { x: 100, y: 320, width: 550, height: 100 },
            textFormat: {
              defaultFontSize: 24,
              defaultFontBold: false,
              defaultAlignment: 'center'
            }
          }
        ]
      },
      {
        name: "Title and Content",
        type: 'content',
        placeholders: [
          {
            type: 'title',
            position: { x: 50, y: 40, width: 650, height: 80 },
            textFormat: {
              defaultFontSize: 32,
              defaultFontBold: true,
              defaultAlignment: 'left'
            }
          },
          {
            type: 'content',
            position: { x: 50, y: 140, width: 650, height: 350 },
            textFormat: {
              defaultFontSize: 18,
              defaultFontBold: false,
              defaultAlignment: 'left'
            }
          }
        ]
      },
      this.createTwoContentLayout(),
      this.createComparisonLayout(),
      this.createBlankLayout()
    ];

    // 実際のスライドから情報を取得して調整
    try {
      const enhancedLayouts = await this.enhanceLayoutsFromSlides(context, commonLayouts);
      return enhancedLayouts;
    } catch (error) {
      console.warn("実際のレイアウト情報の取得に失敗しました。デフォルト値を使用します。", error);
      return commonLayouts;
    }
  }

  /**
   * 2カラムレイアウトを作成
   */
  private createTwoContentLayout(): LayoutInfo {
    return {
      name: "Two Content",
      type: 'comparison',
      placeholders: [
        {
          type: 'title',
          position: { x: 50, y: 40, width: 650, height: 80 },
          textFormat: {
            defaultFontSize: 32,
            defaultFontBold: true,
            defaultAlignment: 'left'
          }
        },
        {
          type: 'content',
          position: { x: 50, y: 140, width: 300, height: 350 },
          textFormat: {
            defaultFontSize: 16,
            defaultFontBold: false,
            defaultAlignment: 'left'
          }
        },
        {
          type: 'content',
          position: { x: 380, y: 140, width: 300, height: 350 },
          textFormat: {
            defaultFontSize: 16,
            defaultFontBold: false,
            defaultAlignment: 'left'
          }
        }
      ]
    };
  }

  /**
   * 比較レイアウトを作成
   */
  private createComparisonLayout(): LayoutInfo {
    return {
      name: "Comparison",
      type: 'comparison',
      placeholders: [
        {
          type: 'title',
          position: { x: 50, y: 40, width: 650, height: 80 },
          textFormat: {
            defaultFontSize: 32,
            defaultFontBold: true,
            defaultAlignment: 'left'
          }
        }
      ]
    };
  }

  /**
   * 空白レイアウトを作成
   */
  private createBlankLayout(): LayoutInfo {
    return {
      name: "Blank",
      type: 'blank',
      placeholders: []
    };
  }

  /**
   * 実際のスライドからレイアウト情報を強化
   */
  private async enhanceLayoutsFromSlides(
    context: PowerPoint.RequestContext, 
    baseLayouts: LayoutInfo[]
  ): Promise<LayoutInfo[]> {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) {
      return baseLayouts;
    }

    // 最初のスライドからレイアウト情報を分析
    const firstSlide = slides.items[0];
    const detectedPlaceholders = await this.themeAnalyzer.detectPlaceholders(context, firstSlide);
    
    // ベースレイアウトを実際の情報で強化
    return baseLayouts.map(layout => ({
      ...layout,
      placeholders: this.mergePlaceholderInfo(layout.placeholders, detectedPlaceholders)
    }));
  }

  /**
   * プレースホルダー情報をマージ
   */
  private mergePlaceholderInfo(
    baseplaceholders: PlaceholderInfo[], 
    detectedPlaceholders: PlaceholderInfo[]
  ): PlaceholderInfo[] {
    return baseplaceholders.map(basePh => {
      const detected = detectedPlaceholders.find(dp => dp.type === basePh.type);
      if (detected) {
        return {
          ...basePh,
          position: detected.position,
          textFormat: {
            ...basePh.textFormat,
            ...detected.textFormat
          }
        };
      }
      return basePh;
    });
  }

  /**
   * テーマ名を検出
   */
  private async detectThemeName(context: PowerPoint.RequestContext): Promise<string> {
    try {
      // PowerPoint.js APIの制限により、テーマ名の直接取得は現在サポートされていない
      // カラースキームやフォントから推測
      const colorScheme = await this.themeAnalyzer.analyzeColorScheme(context);
      
      // 標準的なOfficeテーマを判定
      if (colorScheme.accent1 === '#5B9BD5') {
        return "Office Theme";
      } else if (colorScheme.accent1 === '#0078D4') {
        return "Modern Theme";
      } else {
        return "Custom Theme";
      }
    } catch (error) {
      console.warn("テーマ名の検出に失敗しました:", error);
      return "Unknown Theme";
    }
  }

  /**
   * 指定されたコンテンツタイプに最適なレイアウトを選択
   */
  public selectOptimalLayout(
    contentType: string, 
    contentAmount: number,
    availableLayouts?: LayoutInfo[]
  ): LayoutInfo | null {
    const layouts = availableLayouts || this.cachedThemeInfo?.availableLayouts || [];
    
    let preferredLayoutType: string;
    
    switch (contentType) {
      case 'title':
        preferredLayoutType = 'title';
        break;
      case 'bullets':
        preferredLayoutType = contentAmount > 300 ? 'comparison' : 'content';
        break;
      case 'comparison':
        preferredLayoutType = 'comparison';
        break;
      case 'conclusion':
        preferredLayoutType = 'content';
        break;
      case 'chart':
      case 'image_with_text':
        preferredLayoutType = 'content';
        break;
      default:
        preferredLayoutType = 'content';
    }

    const optimalLayout = layouts.find(layout => layout.type === preferredLayoutType);
    return optimalLayout || layouts.find(layout => layout.type === 'content') || layouts[0] || null;
  }

  /**
   * プレースホルダー情報からスタイルを取得
   */
  public getPlaceholderStyle(
    placeholderType: 'title' | 'subtitle' | 'content' | 'footer',
    layoutInfo?: LayoutInfo
  ): { position: any; textFormat: any } | null {
    if (!layoutInfo) return null;

    const placeholder = layoutInfo.placeholders.find(p => p.type === placeholderType);
    if (!placeholder) return null;

    return {
      position: placeholder.position,
      textFormat: placeholder.textFormat
    };
  }

  /**
   * キャッシュされたテーマ情報を取得
   */
  public getCachedThemeInfo(): ThemeInfo | null {
    return this.cachedThemeInfo;
  }

  /**
   * テーマ情報のキャッシュをクリア
   */
  public clearCache(): void {
    this.cachedThemeInfo = null;
  }

  /**
   * テーマの変更を検出
   */
  public async detectThemeChanges(): Promise<boolean> {
    if (!this.cachedThemeInfo) {
      return false;
    }

    try {
      const currentTheme = await this.getCurrentThemeInfo();
      // 簡易的な変更検出（より詳細な比較が必要な場合は拡張）
      return currentTheme.name !== this.cachedThemeInfo.name ||
             JSON.stringify(currentTheme.colorScheme) !== JSON.stringify(this.cachedThemeInfo.colorScheme);
    } catch (error) {
      console.warn("テーマ変更の検出に失敗しました:", error);
      return false;
    }
  }

  /**
   * テーマ適用のプレビューを生成
   */
  public generateThemePreview(theme: ThemeInfo): {
    title: string;
    description: string;
    colorPalette: string[];
    fontInfo: string;
  } {
    return {
      title: theme.name,
      description: `${theme.availableLayouts.length}種類のレイアウトが利用可能`,
      colorPalette: [
        theme.colorScheme.accent1,
        theme.colorScheme.accent2,
        theme.colorScheme.accent3,
        theme.colorScheme.background1,
        theme.colorScheme.text1
      ],
      fontInfo: `見出し: ${theme.fontScheme.majorFont.latin}, 本文: ${theme.fontScheme.minorFont.latin}`
    };
  }
}