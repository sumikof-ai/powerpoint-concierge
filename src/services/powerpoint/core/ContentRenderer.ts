// src/services/powerpoint/core/ContentRenderer.ts - コンテンツレンダリングサービス
/* global PowerPoint, console */

import { SlideContent, SlideGenerationOptions } from "../types";
import { ThemeApplier } from "./ThemeApplier";

/**
 * スライドコンテンツのレンダリングを担当するサービス
 */
export class ContentRenderer {
  private themeApplier: ThemeApplier;

  constructor() {
    this.themeApplier = new ThemeApplier();
  }

  /**
   * タイトルスライドをレンダリング
   */
  public async renderTitleSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);

    // メインタイトル
    const titleBox = slide.shapes.addTextBox(slideData.title, {
      left: 75,
      top: 150,
      width: 600,
      height: 150,
    });

    await context.sync();

    titleBox.textFrame.textRange.font.size = fontSize.title;
    titleBox.textFrame.textRange.font.bold = true;
    this.themeApplier.applyThemeColors(titleBox, options.theme, "title");

    // サブタイトル（コンテンツがある場合）
    if (slideData.content && slideData.content.length > 0) {
      const subtitleText = slideData.content.join(" • ");
      const subtitleBox = slide.shapes.addTextBox(subtitleText, {
        left: 100,
        top: 320,
        width: 550,
        height: 100,
      });

      await context.sync();

      subtitleBox.textFrame.textRange.font.size = fontSize.subtitle;
      this.themeApplier.applyThemeColors(subtitleBox, options.theme, "subtitle");
    }

    // 装飾要素を追加
    await this.addTitleDecorations(context, slide, options);
  }

  /**
   * 標準コンテンツスライドをレンダリング
   */
  public async renderContentSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);

    // タイトル
    await this.addSlideTitle(context, slide, slideData.title, fontSize, options);

    // コンテンツ（箇条書き）
    if (slideData.content && slideData.content.length > 0) {
      const contentText = this.formatBulletPoints(slideData.content);
      const contentBox = slide.shapes.addTextBox(contentText, {
        left: 80,
        top: 140,
        width: 580,
        height: 350,
      });

      await context.sync();

      contentBox.textFrame.textRange.font.size = fontSize.body;
      this.themeApplier.applyThemeColors(contentBox, options.theme, "body");
    }
  }

  /**
   * 2カラムコンテンツスライドをレンダリング
   */
  public async renderTwoContentSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);

    // タイトル
    await this.addSlideTitle(context, slide, slideData.title, fontSize, options);

    // コンテンツを2つに分割
    if (slideData.content && slideData.content.length > 0) {
      const midPoint = Math.ceil(slideData.content.length / 2);
      const leftContent = slideData.content.slice(0, midPoint);
      const rightContent = slideData.content.slice(midPoint);

      // 左カラム
      if (leftContent.length > 0) {
        await this.renderContentColumn(
          context,
          slide,
          leftContent,
          { left: 50, top: 140, width: 300, height: 350 },
          fontSize,
          options
        );
      }

      // 右カラム
      if (rightContent.length > 0) {
        await this.renderContentColumn(
          context,
          slide,
          rightContent,
          { left: 380, top: 140, width: 300, height: 350 },
          fontSize,
          options
        );
      }

      // 分割線を追加
      await this.addDividerLine(context, slide, options);
    }
  }

  /**
   * 比較スライドをレンダリング
   */
  public async renderComparisonSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);

    // タイトル
    await this.addSlideTitle(context, slide, slideData.title, fontSize, options);

    // 比較ヘッダー
    await this.addComparisonHeaders(context, slide, fontSize, options);

    // コンテンツを交互に配置
    if (slideData.content && slideData.content.length > 0) {
      const contentBoxes = [];
      for (let index = 0; index < slideData.content.length && index < 8; index++) {
        const item = slideData.content[index];
        const yPos = 200 + index * 35;
        const isLeft = index % 2 === 0;

        const contentBox = slide.shapes.addTextBox(`• ${item}`, {
          left: isLeft ? 50 : 380,
          top: yPos,
          width: 300,
          height: 30,
        });
        contentBoxes.push(contentBox);
      }

      // ループ後にまとめてsync
      await context.sync();

      // スタイル適用
      contentBoxes.forEach((contentBox) => {
        contentBox.textFrame.textRange.font.size = fontSize.body;
        this.themeApplier.applyThemeColors(contentBox, options.theme, "body");
      });
    }
  }

  /**
   * 空白スライドをレンダリング
   */
  public async renderBlankSlide(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    slideData: SlideContent,
    options: SlideGenerationOptions
  ): Promise<void> {
    const fontSize = this.themeApplier.getFontSize(options.fontSize);

    // タイトルのみ
    await this.addSlideTitle(context, slide, slideData.title, fontSize, options);
  }

  /**
   * スライドタイトルを追加（共通処理）
   */
  private async addSlideTitle(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    title: string,
    fontSize: any,
    options: SlideGenerationOptions
  ): Promise<void> {
    const titleBox = slide.shapes.addTextBox(title, {
      left: 50,
      top: 40,
      width: 650,
      height: 80,
    });

    await context.sync();

    titleBox.textFrame.textRange.font.size = fontSize.heading;
    titleBox.textFrame.textRange.font.bold = true;
    this.themeApplier.applyThemeColors(titleBox, options.theme, "heading");
  }

  /**
   * コンテンツカラムをレンダリング
   */
  private async renderContentColumn(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    content: string[],
    position: { left: number; top: number; width: number; height: number },
    fontSize: any,
    options: SlideGenerationOptions
  ): Promise<void> {
    const contentText = this.formatBulletPoints(content);
    const contentBox = slide.shapes.addTextBox(contentText, position);

    await context.sync();

    contentBox.textFrame.textRange.font.size = fontSize.body;
    this.themeApplier.applyThemeColors(contentBox, options.theme, "body");
  }

  /**
   * 比較ヘッダーを追加
   */
  private async addComparisonHeaders(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    fontSize: any,
    options: SlideGenerationOptions
  ): Promise<void> {
    // 左側ヘッダー
    const leftHeaderBox = slide.shapes.addTextBox("項目", {
      left: 50,
      top: 140,
      width: 300,
      height: 40,
    });

    await context.sync();

    leftHeaderBox.textFrame.textRange.font.bold = true;
    leftHeaderBox.textFrame.textRange.font.size = fontSize.accent;
    this.themeApplier.applyThemeColors(leftHeaderBox, options.theme, "accent");

    // 右側ヘッダー
    const rightHeaderBox = slide.shapes.addTextBox("詳細", {
      left: 380,
      top: 140,
      width: 300,
      height: 40,
    });

    await context.sync();

    rightHeaderBox.textFrame.textRange.font.bold = true;
    rightHeaderBox.textFrame.textRange.font.size = fontSize.accent;
    this.themeApplier.applyThemeColors(rightHeaderBox, options.theme, "accent");
  }

  /**
   * タイトルスライドの装飾要素を追加
   */
  private async addTitleDecorations(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    options: SlideGenerationOptions
  ): Promise<void> {
    try {
      // アクセント要素（装飾線）
      const accentShape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 50,
        top: 280,
        width: 640,
        height: 4,
      });

      await context.sync();

      const accentColor = this.themeApplier.getAccentColor(options.theme);
      accentShape.fill.setSolidColor(accentColor);
    } catch (error) {
      console.warn("装飾要素の追加に失敗しました:", error);
    }
  }

  /**
   * 分割線を追加
   */
  private async addDividerLine(
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    options: SlideGenerationOptions
  ): Promise<void> {
    try {
      const dividerLine = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 360,
        top: 130,
        width: 2,
        height: 300,
      });

      await context.sync();

      const borderColor = this.themeApplier.getBorderColor(options.theme);
      dividerLine.fill.setSolidColor(borderColor);
    } catch (error) {
      console.warn("分割線の追加に失敗しました:", error);
    }
  }

  /**
   * 箇条書きをフォーマット
   */
  private formatBulletPoints(items: string[]): string {
    return items
      .map((item, index) => {
        const bullet = index === 0 ? "●" : "◦";
        return `${bullet} ${item}`;
      })
      .join("\n\n");
  }

  /**
   * テキストの自動調整
   */
  public adjustTextForSpace(text: string, maxLength: number): string {
    if (text.length <= maxLength) {
      return text;
    }

    // 文字数制限に合わせて調整
    const truncated = text.substring(0, maxLength - 3) + "...";
    return truncated;
  }

  /**
   * 改行位置の最適化
   */
  public optimizeLineBreaks(text: string, maxLineLength: number = 50): string {
    const words = text.split(" ");
    const lines: string[] = [];
    let currentLine = "";

    for (const word of words) {
      if ((currentLine + " " + word).length <= maxLineLength) {
        currentLine = currentLine ? currentLine + " " + word : word;
      } else {
        if (currentLine) {
          lines.push(currentLine);
        }
        currentLine = word;
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }

    return lines.join("\n");
  }
}
