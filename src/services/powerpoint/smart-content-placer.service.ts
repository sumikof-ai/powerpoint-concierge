// src/services/powerpoint/smart-content-placer.service.ts - スマートコンテンツ配置サービス
/* global PowerPoint */

import { ThemeService, ThemeInfo, LayoutInfo, PlaceholderInfo } from './theme.service';
import { SlideContent, SlideGenerationOptions } from './types';

/**
 * テーマに基づいてコンテンツを適切に配置するサービス
 */
export class SmartContentPlacerService {
  private themeService: ThemeService;

  constructor(themeService: ThemeService) {
    this.themeService = themeService;
  }

  /**
   * スライドにコンテンツをスマートに配置
   */
  public async placeContent(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    // テーマ情報を取得
    const themeInfo = await this.themeService.getCurrentThemeInfo();
    
    // 最適なレイアウトを選択
    const optimalLayout = this.themeService.selectOptimalLayout(
      slideData.slideType,
      this.calculateContentAmount(slideData),
      themeInfo.availableLayouts
    );

    if (!optimalLayout) {
      // フォールバック: 通常のテキストボックス配置
      await this.placeContentWithoutLayout(context, slide, slideData, options);
      return;
    }

    // レイアウトに基づいてコンテンツを配置
    await this.placeContentWithLayout(context, slide, slideData, optimalLayout, themeInfo, options);
  }

  /**
   * レイアウト情報を使用してコンテンツを配置
   */
  private async placeContentWithLayout(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    layout: LayoutInfo,
    themeInfo: ThemeInfo,
    options: SlideGenerationOptions
  ): Promise<void> {
    console.log("options:",options);
    // タイトルプレースホルダーの処理
    const titlePlaceholder = layout.placeholders.find(p => p.type === 'title');
    if (titlePlaceholder && slideData.title) {
      await this.addContentToPlaceholder(
        context,
        slide,
        slideData.title,
        titlePlaceholder,
        themeInfo,
        'title'
      );
    }

    // サブタイトルプレースホルダーの処理（タイトルスライドの場合）
    if (slideData.slideType === 'title') {
      const subtitlePlaceholder = layout.placeholders.find(p => p.type === 'subtitle');
      if (subtitlePlaceholder && slideData.content.length > 0) {
        const subtitleText = slideData.content.join(' | ');
        await this.addContentToPlaceholder(
          context,
          slide,
          subtitleText,
          subtitlePlaceholder,
          themeInfo,
          'subtitle'
        );
      }
    }

    // コンテンツプレースホルダーの処理
    const contentPlaceholders = layout.placeholders.filter(p => p.type === 'content');
    if (contentPlaceholders.length > 0 && slideData.content.length > 0) {
      if (contentPlaceholders.length === 1) {
        // 単一のコンテンツエリア
        const contentText = this.formatBulletPoints(slideData.content);
        await this.addContentToPlaceholder(
          context,
          slide,
          contentText,
          contentPlaceholders[0],
          themeInfo,
          'content'
        );
      } else if (contentPlaceholders.length >= 2) {
        // 複数のコンテンツエリア（2カラムレイアウトなど）
        const midPoint = Math.ceil(slideData.content.length / 2);
        const leftContent = slideData.content.slice(0, midPoint);
        const rightContent = slideData.content.slice(midPoint);

        if (leftContent.length > 0) {
          await this.addContentToPlaceholder(
            context,
            slide,
            this.formatBulletPoints(leftContent),
            contentPlaceholders[0],
            themeInfo,
            'content'
          );
        }

        if (rightContent.length > 0 && contentPlaceholders[1]) {
          await this.addContentToPlaceholder(
            context,
            slide,
            this.formatBulletPoints(rightContent),
            contentPlaceholders[1],
            themeInfo,
            'content'
          );
        }
      }
    }

    await context.sync();
  }

  /**
   * プレースホルダーにコンテンツを追加
   */
  private async addContentToPlaceholder(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    content: string,
    placeholder: PlaceholderInfo,
    themeInfo: ThemeInfo,
    contentType: 'title' | 'subtitle' | 'content'
  ): Promise<void> {
    // テキストボックスを作成
    const textBox = slide.shapes.addTextBox(content, placeholder.position);
    await context.sync();

    // テキストフォーマットを適用
    const textFrame = textBox.textFrame;
    const textRange = textFrame.textRange;

    // フォントサイズ
    textRange.font.size = placeholder.textFormat.defaultFontSize;

    // フォントスタイル
    if (placeholder.textFormat.defaultFontBold) {
      textRange.font.bold = true;
    }
    if (placeholder.textFormat.defaultFontItalic) {
      textRange.font.italic = true;
    }

    // フォントファミリー（テーマから取得）
    if (contentType === 'title' || contentType === 'subtitle') {
      textRange.font.name = themeInfo.fontScheme.majorFont.latin;
    } else {
      textRange.font.name = themeInfo.fontScheme.minorFont.latin;
    }

    // テキスト配置
    // PowerPoint.js APIの現在のバージョンでは、直接的なテキスト配置設定が限定的
    // 利用可能なAPIメソッドを確認して設定
    if (placeholder.textFormat.defaultAlignment && placeholder.textFormat.defaultAlignment !== 'left') {
      try {
        // PowerPoint.js APIで利用可能な場合のみ設定
        const alignment = placeholder.textFormat.defaultAlignment;
        console.log(`テキスト配置を設定: ${alignment}`);
        
        // 将来のAPI拡張に備えたプレースホルダー
        // 現在のAPIバージョンではこの機能が制限されている可能性があります
        
        // 代替案: テキストの先頭にスペースを追加して擬似的に配置を調整
        if (alignment === 'center' && content.trim()) {
          // 簡易的な中央寄せ（完全ではない）
          const lines = content.split('\n');
          const maxLength = Math.max(...lines.map(line => line.length));
          const paddedContent = lines.map(line => {
            const padding = Math.floor((maxLength - line.length) / 2);
            return ' '.repeat(Math.max(0, padding)) + line;
          }).join('\n');
          textRange.text = paddedContent;
        }
      } catch (error) {
        console.warn('テキスト配置の設定はこのバージョンのPowerPoint.js APIではサポートされていません');
      }
    }

    // 色の適用（テーマカラーを使用）
    if (contentType === 'title') {
      textRange.font.color = themeInfo.colorScheme.text1;
    } else if (contentType === 'subtitle') {
      textRange.font.color = themeInfo.colorScheme.text2;
    } else {
      textRange.font.color = placeholder.textFormat.defaultColor || themeInfo.colorScheme.text1;
    }

    // 背景色の設定
    textBox.fill.setSolidColor(themeInfo.colorScheme.background1);

    // 行間隔の設定
    if (placeholder.textFormat.defaultLineSpacing) {
      // PowerPoint.js APIでサポートされる場合に設定
      // 現在は制限があるため、将来の実装用
    }

    await context.sync();
  }

  /**
   * レイアウト情報なしでコンテンツを配置（フォールバック）
   */
  private async placeContentWithoutLayout(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    // 既存の実装を使用（基本的なテキストボックス配置）
    const fontSize = this.getDefaultFontSize(options.fontSize);

    // タイトル
    if (slideData.title) {
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 50,
        top: 40,
        width: 650,
        height: 80
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
    }

    // コンテンツ
    if (slideData.content && slideData.content.length > 0) {
      const contentText = this.formatBulletPoints(slideData.content);
      const contentBox = slide.shapes.addTextBox(contentText, {
        left: 80,
        top: 140,
        width: 580,
        height: 350
      });
      
      await context.sync();
      
      contentBox.textFrame.textRange.font.size = fontSize.body;
    }

    await context.sync();
  }

  /**
   * コンテンツ量を計算
   */
  private calculateContentAmount(slideData: SlideContent): number {
    const titleLength = slideData.title.length;
    const contentLength = slideData.content.reduce((sum, item) => sum + item.length, 0);
    return titleLength + contentLength;
  }

  /**
   * 箇条書きをフォーマット
   */
  private formatBulletPoints(items: string[]): string {
    return items.map(item => `• ${item}`).join('\n\n');
  }

  /**
   * デフォルトのフォントサイズを取得
   */
  private getDefaultFontSize(size?: 'small' | 'medium' | 'large') {
    switch (size) {
      case 'small':
        return { title: 32, heading: 20, subtitle: 16, body: 12 };
      case 'large':
        return { title: 44, heading: 28, subtitle: 22, body: 16 };
      case 'medium':
      default:
        return { title: 38, heading: 24, subtitle: 18, body: 14 };
    }
  }

  /**
   * スライドレイアウトを適用
   */
  public async applySlideLayout(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    layoutName: string
  ): Promise<void> {
    try {
      // PowerPoint.js APIでレイアウトを適用
      // 現在のAPIでは直接的なレイアウト適用は限定的
      console.log(`レイアウト「${layoutName}」を適用します`);
      
      // 将来のAPI拡張に備えた実装
      // slide.applyLayout(layoutName);
      // await context.sync();
    } catch (error) {
      console.warn('レイアウトの適用に失敗しました:', error);
      console.warn('適用レイアウト:',[context,slide]);
    }
  }

  /**
   * スライドのプレースホルダーを検出
   */
  public async detectPlaceholders(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide
  ): Promise<PlaceholderInfo[]> {
    const placeholders: PlaceholderInfo[] = [];

    try {
      slide.shapes.load("items");
      await context.sync();

      for (const shape of slide.shapes.items) {
        try {
          shape.load(["type", "left", "top", "width", "height", "textFrame"]);
          await context.sync();
          
          // PowerPoint.js の ShapeType は列挙型
          // 実際の値を確認してプレースホルダーかどうかを判定
          const shapeType = shape.type;
          
          // デバッグ用：実際のtypeの値を確認
          console.log(`Shape type detected: ${shapeType}`);
          
          // プレースホルダーの判定（複数の方法で試行）
          let isPlaceholder = false;
          
          // 方法1: 文字列での比較
          if (typeof shapeType === 'string') {
            isPlaceholder = shapeType === 'Placeholder' || shapeType.toLowerCase() === 'placeholder';
          }
          
          // 方法2: PowerPoint.ShapeType enumとの比較（存在する場合）
          if (!isPlaceholder && PowerPoint.ShapeType) {
            // PowerPoint.ShapeType の各値と比較
            for (const key in PowerPoint.ShapeType) {
              if (PowerPoint.ShapeType[key] === shapeType && key.toLowerCase() === 'placeholder') {
                isPlaceholder = true;
                break;
              }
            }
          }
          
          // 方法3: プレースホルダーの特徴から推測（フォールバック）
          if (!isPlaceholder && shape.textFrame) {
            // テキストフレームがあり、特定の位置にある場合
            const hasTextFrame = true; // textFrameの存在確認
            const isInTitlePosition = shape.top < 100 && shape.height > 40;
            const isInContentPosition = shape.top > 100 && shape.height > 100;
            
            if (hasTextFrame && (isInTitlePosition || isInContentPosition)) {
              console.log('プレースホルダーと推測されるシェイプを検出');
              isPlaceholder = true;
            }
          }
          
          if (isPlaceholder) {
            // プレースホルダー情報を構築
            const placeholderInfo: PlaceholderInfo = {
              type: this.detectPlaceholderType(shape),
              position: {
                x: shape.left,
                y: shape.top,
                width: shape.width,
                height: shape.height
              },
              textFormat: {
                defaultFontSize: 18,
                defaultFontBold: false,
                defaultAlignment: 'left'
              }
            };

            placeholders.push(placeholderInfo);
          }
        } catch (shapeError) {
          console.warn('シェイプの処理中にエラーが発生しました:', shapeError);
        }
      }
    } catch (error) {
      console.warn('プレースホルダーの検出に失敗しました:', error);
    }

    return placeholders;
  }

  /**
   * シェイプからプレースホルダータイプを推測
   */
  private detectPlaceholderType(shape: PowerPoint.Shape): PlaceholderInfo['type'] {
    // PowerPoint.js APIの制限により、現在は位置とサイズから推測
    // また、textFrameの内容も参考にする
    try {
      // シェイプの位置とサイズから推測
      if (shape.top < 100 && shape.height > 50) {
        return 'title';
      } else if (shape.top > 300 && shape.height < 150) {
        return 'subtitle';
      } else if (shape.width > 400 && shape.height > 200) {
        return 'content';
      } else {
        return 'content';
      }
    } catch (error) {
      console.warn('プレースホルダータイプの検出中にエラー:', error);
      return 'content';
    }
  }

  /**
   * テーマに基づいた色の自動選択
   */
  public selectThemeColor(
    themeInfo: ThemeInfo,
    colorType: 'primary' | 'secondary' | 'accent' | 'background' | 'text'
  ): string {
    switch (colorType) {
      case 'primary':
        return themeInfo.colorScheme.accent1;
      case 'secondary':
        return themeInfo.colorScheme.accent2;
      case 'accent':
        return themeInfo.colorScheme.accent3;
      case 'background':
        return themeInfo.colorScheme.background1;
      case 'text':
        return themeInfo.colorScheme.text1;
      default:
        return themeInfo.colorScheme.text1;
    }
  }

  /**
   * 階層構造を持つ箇条書きを作成
   */
  public createHierarchicalBullets(
    items: string[],
    indentLevels?: number[]
  ): string {
    return items.map((item, index) => {
      const indentLevel = indentLevels?.[index] || 0;
      const indent = '  '.repeat(indentLevel);
      const bullet = indentLevel === 0 ? '•' : '◦';
      return `${indent}${bullet} ${item}`;
    }).join('\n');
  }

  /**
   * コンテンツの自動分割
   */
  public splitContentForMultipleSlides(
    content: string[],
    maxItemsPerSlide: number = 5
  ): string[][] {
    const slides: string[][] = [];
    
    for (let i = 0; i < content.length; i += maxItemsPerSlide) {
      slides.push(content.slice(i, i + maxItemsPerSlide));
    }
    
    return slides;
  }
}