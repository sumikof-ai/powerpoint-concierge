// src/services/powerpoint/core/SlideContentGenerator.ts - スライドコンテンツ詳細化サービス
/* global console, setTimeout */

import { OpenAIService } from "../../openai.service";
import { SlideContent, SlideGenerationOptions } from "../types";
import { PresentationOutline, SlideOutline } from "../../../taskpane/components/types";

/**
 * スライド毎のコンテンツ詳細化を担当するサービス
 */
export class SlideContentGenerator {
  private openAIService: OpenAIService;

  constructor(openAIService: OpenAIService) {
    this.openAIService = openAIService;
  }

  /**
   * アウトラインの各スライドを詳細化してPowerPointコンテンツに変換
   */
  public async generateDetailedSlides(
    outline: PresentationOutline,
    options: SlideGenerationOptions,
    onProgress?: (current: number, total: number, slideName: string) => void
  ): Promise<SlideContent[]> {
    const detailedSlides: SlideContent[] = [];

    for (let i = 0; i < outline.slides.length; i++) {
      const slide = outline.slides[i];

      // 進捗報告
      if (onProgress) {
        onProgress(i + 1, outline.slides.length, slide.title);
      }

      // スライドコンテンツを詳細化
      const detailedContent = await this.generateDetailedSlideContent(slide, outline, i, options);

      detailedSlides.push(detailedContent);

      // API呼び出し間隔を調整（レート制限対策）
      if (i < outline.slides.length - 1) {
        await this.delay(500); // 500ms待機
      }
    }

    return detailedSlides;
  }

  /**
   * 個別スライドのコンテンツを詳細化
   */
  private async generateDetailedSlideContent(
    slide: SlideOutline,
    fullOutline: PresentationOutline,
    slideIndex: number,
    options: SlideGenerationOptions
  ): Promise<SlideContent> {
    // コンテキスト情報を構築
    const context = this.buildSlideContext(slide, fullOutline, slideIndex);

    // OpenAI APIを呼び出してコンテンツを詳細化
    const detailedContent = await this.callDetailedContentAPI(slide, context, options);

    return {
      title: detailedContent.title || slide.title,
      content: detailedContent.content,
      slideType: slide.slideType,
      speakerNotes: detailedContent.speakerNotes,
    };
  }

  /**
   * スライドのコンテキスト情報を構築
   */
  private buildSlideContext(
    slide: SlideOutline,
    fullOutline: PresentationOutline,
    slideIndex: number
  ): SlideContext {
    const previousSlide = slideIndex > 0 ? fullOutline.slides[slideIndex - 1] : null;
    const nextSlide =
      slideIndex < fullOutline.slides.length - 1 ? fullOutline.slides[slideIndex + 1] : null;

    return {
      presentationTitle: fullOutline.title,
      slideTitle: slide.title,
      slideNumber: slideIndex + 1,
      totalSlides: fullOutline.slides.length,
      slideType: slide.slideType,
      currentContent: slide.content,
      previousSlideTitle: previousSlide?.title || null,
      previousSlideContent: previousSlide?.content || null,
      nextSlideTitle: nextSlide?.title || null,
      nextSlideContent: nextSlide?.content || null,
      estimatedDuration: fullOutline.estimatedDuration,
    };
  }

  /**
   * OpenAI APIを呼び出してコンテンツを詳細化
   */
  private async callDetailedContentAPI(
    slide: SlideOutline,
    context: SlideContext,
    options: SlideGenerationOptions
  ): Promise<DetailedSlideContent> {
    const systemPrompt = this.buildSystemPrompt(slide.slideType, options);
    const userPrompt = this.buildUserPrompt(slide, context);

    try {
      const response = await this.openAIService.sendRequest([
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ]);

      // レスポンスをパース
      return this.parseDetailedContentResponse(response, slide);
    } catch (error) {
      console.error(`スライド ${slide.slideNumber} の詳細化でエラー:`, error);

      // フォールバック: 元のコンテンツを返す
      return {
        title: slide.title,
        content: slide.content.map((item) => `• ${item}`),
        speakerNotes: `スライド ${slide.slideNumber}: ${slide.title}`,
      };
    }
  }

  /**
   * スライドタイプに応じたシステムプロンプトを構築
   */
  private buildSystemPrompt(slideType: string, options: SlideGenerationOptions): string {
    console.log(`not implemented ${options}`);
    const basePrompt = `
あなたは優秀なプレゼンテーション作成のプロフェッショナルです。
与えられたアウトライン情報から、説明資料として使える詳細で実用的なスライドコンテンツを作成してください。

【重要な要件】
1. 単なる箇条書きではなく、説明資料として理解しやすい詳細な内容を作成
2. 聴衆が自立して理解できるレベルの情報を含める
3. 具体例、数値、手順などを適切に含める
4. ビジネス文書として適切なトーンと専門性を保つ
5. 各ポイントは独立して理解できるように詳細化する

【出力形式】
必ず以下のJSON形式で回答してください：
{
  "title": "詳細化されたスライドタイトル",
  "content": [
    "詳細化されたポイント1（説明や具体例を含む）",
    "詳細化されたポイント2（データや手順を含む）",
    "詳細化されたポイント3（背景や理由を含む）"
  ],
  "speakerNotes": "発表者向けの補足説明（発表のコツ、強調ポイント、質疑応答対策など）"
}

【スライドタイプ別の指針】`;

    switch (slideType) {
      case "title":
        return (
          basePrompt +
          `
- タイトルスライド：印象的で興味を引くタイトルに調整
- サブタイトルまたは概要として、プレゼンテーションの価値と期待できる成果を明記
- 聴衆にとっての意義や重要性を伝える`
        );

      case "content":
        return (
          basePrompt +
          `
- メインコンテンツ：各ポイントを詳細に展開
- 背景、現状、課題、解決策、効果などを具体的に説明
- データ、事例、比較情報を積極的に含める
- 実装可能な具体的なアクションアイテムを提示`
        );

      case "conclusion":
        return (
          basePrompt +
          `
- まとめスライド：プレゼンテーション全体の要点を再整理
- 次のステップや具体的なアクションプランを明示
- 聴衆への明確なメッセージと行動喚起を含める
- プレゼンテーションの価値と成果を再確認`
        );

      default:
        return (
          basePrompt +
          `
- 一般的なコンテンツスライドとして詳細化
- 論理的な構成で情報を整理
- 具体性と実用性を重視`
        );
    }
  }

  /**
   * ユーザープロンプトを構築
   */
  private buildUserPrompt(_slide: SlideOutline, context: SlideContext): string {
    // TODO: implement user prompt building logic
    return `
【プレゼンテーション全体の情報】
タイトル: ${context.presentationTitle}
総スライド数: ${context.totalSlides}
予想時間: ${context.estimatedDuration}分

【現在のスライド情報】
スライド番号: ${context.slideNumber}/${context.totalSlides}
スライドタイプ: ${context.slideType}
現在のタイトル: ${context.slideTitle}
現在のコンテンツ: ${context.currentContent.map((item, idx) => `${idx + 1}. ${item}`).join("\n")}

【前後のスライドコンテキスト】
${context.previousSlideTitle ? `前のスライド: ${context.previousSlideTitle}` : "（最初のスライド）"}
${context.nextSlideTitle ? `次のスライド: ${context.nextSlideTitle}` : "（最後のスライド）"}

【詳細化の指示】
上記のアウトライン情報を基に、説明資料として使える詳細なスライドコンテンツを作成してください。

- 各ポイントを3-5倍の詳細さに拡張
- 具体例、データ、手順、理由などを含める
- 前後のスライドとの流れを考慮した内容にする
- 聴衆が自立して理解できるレベルの説明を含める
- ビジネス現場で実際に使える実用的な内容にする

必ずJSON形式で回答してください。`;
  }

  /**
   * APIレスポンスをパースして詳細コンテンツを抽出
   */
  private parseDetailedContentResponse(
    response: string,
    originalSlide: SlideOutline
  ): DetailedSlideContent {
    try {
      // JSON部分を抽出
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error("JSON形式の応答が見つかりません");
      }

      const parsed = JSON.parse(jsonMatch[0]) as DetailedSlideContent;

      // 基本的な検証
      if (!parsed.title || !Array.isArray(parsed.content)) {
        throw new Error("無効なレスポンス形式");
      }

      // コンテンツが空の場合はフォールバック
      if (parsed.content.length === 0) {
        parsed.content = originalSlide.content.map(
          (item) => `• ${item}（詳細化処理でエラーが発生）`
        );
      }

      // スピーカーノートが無い場合はデフォルトを設定
      if (!parsed.speakerNotes) {
        parsed.speakerNotes = `スライド ${originalSlide.slideNumber}: ${parsed.title}`;
      }

      return parsed;
    } catch (error) {
      console.error("詳細コンテンツのパースでエラー:", error);

      // フォールバック
      return {
        title: originalSlide.title,
        content: originalSlide.content.map((item) => `• ${item}`),
        speakerNotes: `スライド ${originalSlide.slideNumber}: ${originalSlide.title}`,
      };
    }
  }

  /**
   * 遅延処理（API制限対策）
   */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * コンテンツの長さを調整
   */
  public adjustContentLength(content: string[], maxLength: number = 200): string[] {
    return content.map((item) => {
      if (item.length <= maxLength) {
        return item;
      }

      // 長すぎる場合は適切な位置で改行
      const sentences = item.split("。");
      let result = "";

      for (const sentence of sentences) {
        if ((result + sentence + "。").length <= maxLength) {
          result += sentence + "。";
        } else {
          break;
        }
      }

      return result || item.substring(0, maxLength - 3) + "...";
    });
  }

  /**
   * スライドコンテンツの品質チェック
   */
  public validateSlideContent(content: SlideContent): {
    isValid: boolean;
    warnings: string[];
    suggestions: string[];
  } {
    const warnings: string[] = [];
    const suggestions: string[] = [];

    // タイトルの長さチェック
    if (content.title.length > 120) {
      warnings.push("タイトルが長すぎます");
      suggestions.push("タイトルを簡潔にまとめることを推奨");
    }

    // コンテンツ項目数のチェック
    if (content.content.length > 7) {
      warnings.push("コンテンツ項目が多すぎます");
      suggestions.push("重要なポイントに絞り込むことを推奨");
    }

    // 各項目の長さチェック
    content.content.forEach((item, index) => {
      if (item.length > 300) {
        warnings.push(`項目 ${index + 1} が長すぎます`);
        suggestions.push(`項目 ${index + 1} を複数の項目に分割することを推奨`);
      }
    });

    // 読みやすさのチェック
    const totalLength = content.content.join("").length;
    if (totalLength > 1000) {
      warnings.push("総文字数が多すぎます");
      suggestions.push("内容を簡潔にまとめるか、複数スライドに分割することを推奨");
    }

    return {
      isValid: warnings.length === 0,
      warnings,
      suggestions,
    };
  }

  /**
   * エラー時のフォールバック用コンテンツ生成
   */
  public createFallbackContent(slide: SlideOutline): SlideContent {
    return {
      title: slide.title,
      content: slide.content.map((item) => `• ${item}（標準コンテンツ）`),
      slideType: slide.slideType,
      speakerNotes: `スライド ${slide.slideNumber}: ${slide.title}の標準版コンテンツです。`,
    };
  }

  /**
   * バッチ処理のエラーハンドリング
   */
  public async generateWithErrorHandling(
    outline: PresentationOutline,
    options: SlideGenerationOptions,
    onProgress?: (current: number, total: number, slideName: string) => void,
    onError?: (slideIndex: number, error: Error) => void
  ): Promise<SlideContent[]> {
    const results: SlideContent[] = [];
    const errors: Array<{ slideIndex: number; error: Error }> = [];

    for (let i = 0; i < outline.slides.length; i++) {
      try {
        if (onProgress) {
          onProgress(i + 1, outline.slides.length, outline.slides[i].title);
        }

        const detailedContent = await this.generateDetailedSlideContent(
          outline.slides[i],
          outline,
          i,
          options
        );

        results.push(detailedContent);
      } catch (error) {
        console.error(`スライド ${i + 1} の生成でエラー:`, error);

        // エラーコールバックを呼び出し
        if (onError) {
          onError(i, error instanceof Error ? error : new Error("不明なエラー"));
        }

        // フォールバックコンテンツを作成
        const fallbackContent = this.createFallbackContent(outline.slides[i]);
        results.push(fallbackContent);

        errors.push({
          slideIndex: i,
          error: error instanceof Error ? error : new Error("不明なエラー"),
        });
      }

      // API制限対策の待機
      if (i < outline.slides.length - 1) {
        await this.delay(500);
      }
    }

    // エラーサマリーをログ出力
    if (errors.length > 0) {
      console.warn(`${errors.length}個のスライドで詳細化エラーが発生しました:`, errors);
    }

    return results;
  }
}

// 型定義
interface SlideContext {
  presentationTitle: string;
  slideTitle: string;
  slideNumber: number;
  totalSlides: number;
  slideType: string;
  currentContent: string[];
  previousSlideTitle: string | null;
  previousSlideContent: string[] | null;
  nextSlideTitle: string | null;
  nextSlideContent: string[] | null;
  estimatedDuration: number;
}

interface DetailedSlideContent {
  title: string;
  content: string[];
  speakerNotes: string;
}
