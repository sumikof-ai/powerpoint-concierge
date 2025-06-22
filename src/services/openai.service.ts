// src/services/openai.service.ts
/* global console, fetch */
import { OpenAISettings, OpenAIRequest, OpenAIResponse } from "../taskpane/components/types";
import { PresentationOutline } from "../taskpane/components/types";

export class OpenAIService {
  private settings: OpenAISettings;

  constructor(settings: OpenAISettings) {
    this.settings = settings;
  }

  public updateSettings(settings: OpenAISettings) {
    this.settings = settings;
  }

  /**
   * OpenAI APIにリクエストを送信
   */
  public async sendRequest(
    messages: { role: "system" | "user" | "assistant"; content: string }[]
  ): Promise<string> {
    if (!this.settings.apiKey) {
      throw new Error("APIキーが設定されていません");
    }

    const request: OpenAIRequest = {
      model: this.settings.model,
      messages,
      temperature: this.settings.temperature,
      max_tokens: this.settings.maxTokens,
    };

    try {
      const response = await fetch(`${this.settings.baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${this.settings.apiKey}`,
        },
        body: JSON.stringify(request),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(
          errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`
        );
      }

      const data: OpenAIResponse = await response.json();

      if (!data.choices || data.choices.length === 0) {
        throw new Error("APIレスポンスが無効です");
      }

      return data.choices[0].message.content;
    } catch (error) {
      console.error("OpenAI API Error:", error);
      if (error instanceof Error) {
        throw error;
      }
      throw new Error("APIリクエストでエラーが発生しました");
    }
  }

  /**
   * プレゼンテーションのアウトライン生成（構造化データとして返す）
   */
  public async generateStructuredOutline(topic: string): Promise<PresentationOutline> {
    const systemPrompt = `
あなたは優秀なプレゼンテーション作成アシスタントです。
与えられたトピックについて、効果的で構造化されたプレゼンテーションのアウトラインを作成してください。

必ず以下のJSON形式で回答してください：
{
  "title": "プレゼンテーションのタイトル",
  "estimatedDuration": 数値（分単位）,
  "slides": [
    {
      "slideNumber": 1,
      "title": "スライドタイトル",
      "content": ["要点1", "要点2", "要点3"],
      "slideType": "title|content|conclusion",
      "speakerNotes": "オプション：スピーカーノート"
    }
  ]
}

slideTypeの説明：
- "title": タイトルスライド（表紙）
- "content": メインコンテンツスライド
- "conclusion": まとめ・結論スライド

一般的な構成：
1. タイトルスライド
2. 目次/アジェンダ（オプション）
3. メインコンテンツ（3-7個のセクション）
4. まとめ/結論
5. 質疑応答（オプション）

各スライドのcontentは3-5個の要点を含めてください。
日本語で回答し、JSONのみを返してください。
    `;

    const userPrompt = `以下のトピックについてプレゼンテーションのアウトラインを作成してください：\n\n${topic}`;

    try {
      const response = await this.sendRequest([
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ]);

      // レスポンスからJSONを抽出
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error("JSONフォーマットの応答が得られませんでした");
      }

      const outline: PresentationOutline = JSON.parse(jsonMatch[0]);

      // データの検証
      if (!outline.title || !outline.slides || !Array.isArray(outline.slides)) {
        throw new Error("無効なアウトライン形式です");
      }

      // スライド番号の正規化
      outline.slides.forEach((slide, index) => {
        slide.slideNumber = index + 1;
      });

      return outline;
    } catch (error) {
      console.error("Structured outline generation error:", error);
      throw new Error(
        `アウトライン生成エラー: ${error instanceof Error ? error.message : "不明なエラー"}`
      );
    }
  }

  /**
   * 既存のアウトラインを修正指示に基づいて再生成
   */
  public async regenerateOutline(
    currentOutline: PresentationOutline,
    instruction: string
  ): Promise<PresentationOutline> {
    const systemPrompt = `
あなたは優秀なプレゼンテーション作成アシスタントです。
現在のプレゼンテーションアウトラインを、ユーザーの指示に従って修正してください。

必ず以下のJSON形式で回答してください：
{
  "title": "プレゼンテーションのタイトル",
  "estimatedDuration": 数値（分単位）,
  "slides": [
    {
      "slideNumber": 1,
      "title": "スライドタイトル",
      "content": ["要点1", "要点2", "要点3"],
      "slideType": "title|content|conclusion",
      "speakerNotes": "オプション：スピーカーノート"
    }
  ]
}

既存の構造を活かしつつ、指示に従って適切に修正してください。
日本語で回答し、JSONのみを返してください。
    `;

    const userPrompt = `
現在のアウトライン:
${JSON.stringify(currentOutline, null, 2)}

修正指示:
${instruction}

上記の修正指示に従って、アウトラインを更新してください。
    `;

    try {
      const response = await this.sendRequest([
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ]);

      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error("JSONフォーマットの応答が得られませんでした");
      }

      const outline: PresentationOutline = JSON.parse(jsonMatch[0]);

      if (!outline.title || !outline.slides || !Array.isArray(outline.slides)) {
        throw new Error("無効なアウトライン形式です");
      }

      outline.slides.forEach((slide, index) => {
        slide.slideNumber = index + 1;
      });

      return outline;
    } catch (error) {
      console.error("Outline regeneration error:", error);
      throw new Error(
        `アウトライン再生成エラー: ${error instanceof Error ? error.message : "不明なエラー"}`
      );
    }
  }

  /**
   * プレゼンテーションのアウトライン生成（従来版、テキストで返す）
   */
  public async generateOutline(topic: string): Promise<string> {
    const systemPrompt = `
あなたは優秀なプレゼンテーション作成アシスタントです。
与えられたトピックについて、効果的で構造化されたプレゼンテーションのアウトラインを作成してください。

アウトラインの形式:
1. タイトルスライド
2. 目次/アジェンダ
3. メインコンテンツ（3-5個のセクション）
4. まとめ/結論
5. 質疑応答

各スライドについて以下の情報を含めてください：
- スライドタイトル
- 主要なポイント（3-5個）
- スピーカーノート（オプション）

日本語で回答してください。
    `;

    const userPrompt = `以下のトピックについてプレゼンテーションのアウトラインを作成してください：\n\n${topic}`;

    return await this.sendRequest([
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ]);
  }

  /**
   * 特定のスライドコンテンツ生成
   */
  public async generateSlideContent(slideTitle: string, context: string): Promise<string> {
    const systemPrompt = `
あなたは優秀なプレゼンテーション作成アシスタントです。
指定されたスライドタイトルについて、詳細なコンテンツを作成してください。

コンテンツの形式:
- スライドタイトル
- メインコンテンツ（箇条書きまたは段落形式）
- 重要なポイントやキーワード
- スピーカーノート

PowerPointで使用できるテキスト形式で出力してください。
日本語で回答してください。
    `;

    const userPrompt = `
コンテキスト: ${context}
スライドタイトル: ${slideTitle}

上記について詳細なスライドコンテンツを作成してください。
    `;

    return await this.sendRequest([
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ]);
  }

  /**
   * スライド編集用のコンテンツ生成
   */
  public async editSlideContent(currentContent: string, editInstruction: string): Promise<string> {
    const systemPrompt = `
あなたは優秀なプレゼンテーション編集アシスタントです。
現在のスライドコンテンツを、指定された編集指示に従って修正してください。

PowerPointで使用できるテキスト形式で出力してください。
日本語で回答してください。
    `;

    const userPrompt = `
現在のコンテンツ:
${currentContent}

編集指示:
${editInstruction}

上記の編集指示に従って、コンテンツを修正してください。
    `;

    return await this.sendRequest([
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ]);
  }
}
