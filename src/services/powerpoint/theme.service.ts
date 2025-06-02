// src/services/powerpoint/theme.service.ts - PowerPointテーマ情報取得サービス
/* global PowerPoint */

/**
 * PowerPointのテーマ情報を取得・管理するサービス
 */
export class ThemeService {
  private cachedThemeInfo: ThemeInfo | null = null;

  /**
   * 現在のプレゼンテーションのテーマ情報を取得
   */
  public async getCurrentThemeInfo(): Promise<ThemeInfo> {
    return new Promise((resolve, reject) => {
      PowerPoint.run(async (context) => {
        try {
          const presentation = context.presentation;
          
          // プレゼンテーションの基本情報を読み込み
          presentation.load("title");
          await context.sync();

          // スライドマスターからレイアウト情報を取得
          const layouts = await this.getAvailableLayouts(context);
          
          // カラースキームを取得（現時点ではデフォルト値を使用）
          const colorScheme = await this.getColorScheme(context);
          
          // フォントスキームを取得（現時点ではデフォルト値を使用）
          const fontScheme = await this.getFontScheme(context);

          const themeInfo: ThemeInfo = {
            name: "Default Theme", // PowerPoint.js APIの制限により、テーマ名の直接取得は現在サポートされていない
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
    // PowerPoint.js APIの現在の制限により、レイアウトの詳細な情報取得は限定的
    // 一般的なレイアウトタイプを事前定義
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
      {
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
      },
      {
        name: "Blank",
        type: 'blank',
        placeholders: []
      }
    ];

    // 実際のスライドマスターから情報を取得する試み
    try {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 最初のスライドからレイアウト情報を推測
      if (slides.items.length > 0) {
        const firstSlide = slides.items[0];
        firstSlide.load("layout");
        await context.sync();
        
        // レイアウト名が取得できた場合は、それに基づいてレイアウト情報を調整
        // （現在のAPIでは限定的）
      }
    } catch (error) {
      console.warn("レイアウト情報の取得に失敗しました。デフォルト値を使用します。", error);
    }

    return commonLayouts;
  }

  /**
   * カラースキームを取得
   */
  private async getColorScheme(context: PowerPoint.RequestContext): Promise<ColorScheme> {
    // PowerPoint.js APIの制限により、現在はデフォルト値を返す
    // 将来的にAPIが拡張された場合、実際のテーマカラーを取得する実装に変更
    
    try {
      // 最初のスライドからカラー情報を推測する試み
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      if (slides.items.length > 0) {
        const firstSlide = slides.items[0];
        firstSlide.shapes.load("items");
        await context.sync();

        // シェイプから使用されている色を分析
        // （現在のAPIでは限定的）
      }
    } catch (error) {
      console.warn("カラースキームの取得に失敗しました。デフォルト値を使用します。", error);
    }

    // デフォルトのOfficeテーマカラー
    return {
      accent1: '#5B9BD5',
      accent2: '#ED7D31',
      accent3: '#A5A5A5',
      accent4: '#FFC000',
      accent5: '#4472C4',
      accent6: '#70AD47',
      background1: '#FFFFFF',
      background2: '#F2F2F2',
      text1: '#000000',
      text2: '#404040',
      hyperlink: '#0563C1',
      followedHyperlink: '#954F72'
    };
  }

  /**
   * フォントスキームを取得
   */
  private async getFontScheme(context: PowerPoint.RequestContext): Promise<FontScheme> {
    // PowerPoint.js APIの制限により、現在はデフォルト値を返す
    
    try {
      // プレゼンテーション全体のフォント設定を取得する試み
      const presentation = context.presentation;
      presentation.load("properties");
      await context.sync();

      // 実際のフォント情報を取得
      // （現在のAPIでは限定的）
    } catch (error) {
      console.warn("フォントスキームの取得に失敗しました。デフォルト値を使用します。", error);
    }

    // デフォルトのフォントスキーム
    return {
      majorFont: {
        latin: 'Calibri Light',
        eastAsian: 'MS Gothic',
        complexScript: 'Arial'
      },
      minorFont: {
        latin: 'Calibri',
        eastAsian: 'MS Gothic',
        complexScript: 'Arial'
      }
    };
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
    
    // コンテンツタイプに基づいてレイアウトを選択
    let preferredLayoutType: string;
    
    switch (contentType) {
      case 'title':
        preferredLayoutType = 'title';
        break;
      case 'bullets':
        preferredLayoutType = contentAmount > 100 ? 'comparison' : 'content';
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

    // 優先レイアウトを検索
    const optimalLayout = layouts.find(layout => layout.type === preferredLayoutType);
    
    // 見つからない場合は最初のコンテンツレイアウトを返す
    return optimalLayout || layouts.find(layout => layout.type === 'content') || layouts[0] || null;
  }

  /**
   * プレースホルダー情報からテキストボックスの位置とスタイルを決定
   */
  public getPlaceholderStyle(
    placeholderType: 'title' | 'subtitle' | 'content' | 'footer',
    layoutInfo?: LayoutInfo
  ): { position: any; textFormat: TextFormatInfo } | null {
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
}

// 型定義
export interface ThemeInfo {
  name: string;
  colorScheme: ColorScheme;
  fontScheme: FontScheme;
  availableLayouts: LayoutInfo[];
}

export interface ColorScheme {
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  background1: string;
  background2: string;
  text1: string;
  text2: string;
  hyperlink: string;
  followedHyperlink: string;
}

export interface FontScheme {
  majorFont: FontSet;
  minorFont: FontSet;
}

export interface FontSet {
  latin: string;
  eastAsian: string;
  complexScript: string;
}

export interface LayoutInfo {
  name: string;
  type: 'title' | 'content' | 'comparison' | 'blank' | 'section' | 'custom';
  placeholders: PlaceholderInfo[];
  isCustom?: boolean;
}

export interface PlaceholderInfo {
  type: 'title' | 'subtitle' | 'content' | 'footer' | 'slideNumber' | 'date' | 'picture' | 'chart';
  position: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  textFormat: TextFormatInfo;
  isRequired?: boolean;
}

export interface TextFormatInfo {
  defaultFontSize: number;
  defaultFontBold: boolean;
  defaultFontItalic?: boolean;
  defaultAlignment: 'left' | 'center' | 'right' | 'justify';
  defaultColor?: string;
  defaultLineSpacing?: number;
  bulletStyle?: BulletStyle;
}

export interface BulletStyle {
  type: 'none' | 'bullet' | 'number' | 'custom';
  symbol?: string;
  indentLevel?: number;
  color?: string;
}