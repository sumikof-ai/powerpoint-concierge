// src/taskpane/services/openai.ts
import { OpenAISettings, OpenAIRequest, OpenAIResponse, APIError } from '../taskpane/components/types';

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
  public async sendRequest(messages: { role: 'system' | 'user' | 'assistant'; content: string }[]): Promise<string> {
    if (!this.settings.apiKey) {
      throw new Error('APIキーが設定されていません');
    }

    const request: OpenAIRequest = {
      model: this.settings.model,
      messages,
      temperature: this.settings.temperature,
      max_tokens: this.settings.maxTokens,
    };

    try {
      const response = await fetch(`${this.settings.baseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.settings.apiKey}`,
        },
        body: JSON.stringify(request),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`);
      }

      const data: OpenAIResponse = await response.json();
      
      if (!data.choices || data.choices.length === 0) {
        throw new Error('APIレスポンスが無効です');
      }

      return data.choices[0].message.content;
    } catch (error) {
      console.error('OpenAI API Error:', error);
      if (error instanceof Error) {
        throw error;
      }
      throw new Error('APIリクエストでエラーが発生しました');
    }
  }

  /**
   * プレゼンテーションのアウトライン生成
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
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
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
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
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
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ]);
  }
}