// TemplateBasedGenerationService.ts
// テンプレートベースのスマート生成サービス

import {
  TemplateInfo,
  TemplateRecommendation,
  TemplateSelectionCriteria,
  TemplateLibrary,
  AdaptedOutline,
  TemplateCategory,
  TemplateRegistrationRequest,
  TemplateUsageStats
} from '../template-types';
import { TemplateAdaptationService } from './TemplateAdaptationService';

export class TemplateBasedGenerationService {
  private adaptationService: TemplateAdaptationService;
  private templateLibrary: TemplateLibrary;
  private usageStats: Map<string, TemplateUsageStats>;

  constructor() {
    this.adaptationService = new TemplateAdaptationService();
    this.templateLibrary = this.initializeTemplateLibrary();
    this.usageStats = new Map();
    this.loadUsageStats();
  }

  /**
   * テンプレートをライブラリに登録
   */
  async registerTemplate(request: TemplateRegistrationRequest): Promise<TemplateInfo> {
    try {
      let templateInfo: TemplateInfo;

      if (request.file) {
        // ファイルベースの登録
        templateInfo = await this.registerTemplateFromFile(request);
      } else {
        // 現在のプレゼンテーションから登録
        templateInfo = await this.registerCurrentPresentationAsTemplate(request);
      }

      // ライブラリに追加
      this.addTemplateToLibrary(templateInfo);
      
      // 使用統計の初期化
      this.usageStats.set(templateInfo.id, {
        templateId: templateInfo.id,
        usageCount: 0,
        lastUsed: new Date(),
        averageRating: 0,
        userFeedback: [],
        successRate: 0
      });

      await this.saveTemplateLibrary();
      
      return templateInfo;
    } catch (error) {
      console.error('Template registration failed:', error);
      throw new Error(`テンプレートの登録に失敗しました: ${error.message}`);
    }
  }

  /**
   * ユーザー入力から最適なテンプレートを選択
   */
  async selectOptimalTemplate(
    userInput: string,
    criteria?: Partial<TemplateSelectionCriteria>
  ): Promise<TemplateRecommendation[]> {
    try {
      const selectionCriteria: TemplateSelectionCriteria = {
        userInput,
        presentationContext: criteria?.presentationContext || {},
        preferences: {
          maxResults: 5,
          minimumScore: 0.3,
          ...criteria?.preferences
        }
      };

      // AI分析によるコンテキスト抽出
      const analyzedContext = await this.analyzeUserInput(userInput);
      
      // テンプレートスコアリング
      const scoredTemplates = await this.scoreTemplates(analyzedContext, selectionCriteria);
      
      // 上位テンプレートの推奨理由生成
      const recommendations = await this.generateRecommendations(scoredTemplates, analyzedContext);
      
      return recommendations
        .filter(r => r.score >= selectionCriteria.preferences.minimumScore!)
        .slice(0, selectionCriteria.preferences.maxResults);

    } catch (error) {
      console.error('Template selection failed:', error);
      
      // フォールバック: デフォルトテンプレートを返す
      return this.getDefaultRecommendations();
    }
  }

  /**
   * アウトラインをテンプレートに適応
   */
  async adaptOutlineToTemplate(
    outline: any,
    templateId: string
  ): Promise<AdaptedOutline> {
    const template = this.getTemplateById(templateId);
    if (!template) {
      throw new Error(`Template not found: ${templateId}`);
    }

    try {
      const adaptedOutline = await this.adaptationService.adaptOutlineToTemplate(outline, template);
      
      // 使用統計の更新
      this.updateTemplateUsage(templateId);
      
      return adaptedOutline;
    } catch (error) {
      console.error('Outline adaptation failed:', error);
      throw error;
    }
  }

  /**
   * テンプレート最適化されたコンテンツを生成
   */
  async generateTemplateOptimizedContent(
    adaptedOutline: AdaptedOutline,
    openAIService: any,
    progressCallback?: (current: number, total: number, message: string) => void
  ): Promise<any[]> {
    try {
      const totalSlides = adaptedOutline.adaptedSlides.length;
      const results = [];

      for (let i = 0; i < totalSlides; i++) {
        const slide = adaptedOutline.adaptedSlides[i];
        
        if (progressCallback) {
          progressCallback(i + 1, totalSlides, `テンプレート最適化中: ${slide.adaptedContent.title}`);
        }

        try {
          // テンプレート特化のコンテンツ生成
          const optimizedContent = await this.generateTemplateSpecificContent(
            slide,
            adaptedOutline.selectedTemplate,
            openAIService
          );

          results.push(optimizedContent);
        } catch (error) {
          console.error(`Failed to optimize slide ${i + 1}:`, error);
          
          // フォールバック: 基本的なコンテンツを使用
          results.push(this.createFallbackContent(slide));
        }
      }

      return results;
    } catch (error) {
      console.error('Template-optimized content generation failed:', error);
      throw error;
    }
  }

  /**
   * テンプレートライブラリの取得
   */
  getTemplateLibrary(): TemplateLibrary {
    return { ...this.templateLibrary };
  }

  /**
   * カテゴリ別テンプレート取得
   */
  getTemplatesByCategory(category: TemplateCategory): TemplateInfo[] {
    return this.templateLibrary.categories[category] || [];
  }

  /**
   * テンプレート検索
   */
  searchTemplates(query: string, filters?: {
    categories?: TemplateCategory[];
    industry?: string[];
    purpose?: string[];
  }): TemplateInfo[] {
    const queryLower = query.toLowerCase();
    
    return this.templateLibrary.templates.filter(template => {
      // テキスト検索
      const matchesQuery = 
        template.name.toLowerCase().includes(queryLower) ||
        template.description?.toLowerCase().includes(queryLower) ||
        template.metadata.tags.some(tag => tag.toLowerCase().includes(queryLower));

      if (!matchesQuery) return false;

      // フィルター適用
      if (filters?.categories && !filters.categories.includes(template.category)) {
        return false;
      }

      if (filters?.industry && template.metadata.industry) {
        const hasMatchingIndustry = filters.industry.some(industry =>
          template.metadata.industry!.includes(industry)
        );
        if (!hasMatchingIndustry) return false;
      }

      if (filters?.purpose && template.metadata.purpose !== filters.purpose[0]) {
        return false;
      }

      return true;
    });
  }

  /**
   * 人気テンプレートの取得
   */
  getPopularTemplates(limit: number = 10): TemplateInfo[] {
    const templatesWithUsage = this.templateLibrary.templates.map(template => ({
      template,
      usage: this.usageStats.get(template.id)
    }));

    return templatesWithUsage
      .sort((a, b) => (b.usage?.usageCount || 0) - (a.usage?.usageCount || 0))
      .slice(0, limit)
      .map(item => item.template);
  }

  /**
   * 最近使用したテンプレートの取得
   */
  getRecentTemplates(limit: number = 5): TemplateInfo[] {
    const templatesWithUsage = this.templateLibrary.templates.map(template => ({
      template,
      usage: this.usageStats.get(template.id)
    }));

    return templatesWithUsage
      .filter(item => item.usage?.lastUsed)
      .sort((a, b) => b.usage!.lastUsed.getTime() - a.usage!.lastUsed.getTime())
      .slice(0, limit)
      .map(item => item.template);
  }

  /**
   * テンプレート使用フィードバック
   */
  recordTemplateFeedback(
    templateId: string,
    rating: number,
    feedback?: string,
    success?: boolean
  ): void {
    const stats = this.usageStats.get(templateId);
    if (!stats) return;

    // 評価の更新
    if (stats.averageRating) {
      stats.averageRating = (stats.averageRating + rating) / 2;
    } else {
      stats.averageRating = rating;
    }

    // フィードバックの追加
    if (feedback) {
      stats.userFeedback = stats.userFeedback || [];
      stats.userFeedback.push(feedback);
    }

    // 成功率の更新
    if (success !== undefined) {
      // 簡単な成功率計算（実際の実装ではより複雑な計算を行う）
      stats.successRate = success ? 
        Math.min(1.0, (stats.successRate || 0) + 0.1) : 
        Math.max(0.0, (stats.successRate || 0) - 0.1);
    }

    this.saveUsageStats();
  }

  private async analyzeUserInput(userInput: string): Promise<any> {
    // ユーザー入力の分析（実際の実装ではより高度な分析を行う）
    const analysis = {
      keywords: this.extractKeywords(userInput),
      suggestedCategory: this.suggestCategory(userInput),
      estimatedComplexity: this.estimateComplexity(userInput),
      detectedPurpose: this.detectPurpose(userInput),
      audienceLevel: this.detectAudienceLevel(userInput)
    };

    return analysis;
  }

  private async scoreTemplates(
    analyzedContext: any,
    criteria: TemplateSelectionCriteria
  ): Promise<Array<{ template: TemplateInfo; score: number }>> {
    return this.templateLibrary.templates.map(template => ({
      template,
      score: this.calculateTemplateScore(template, analyzedContext, criteria)
    })).sort((a, b) => b.score - a.score);
  }

  private calculateTemplateScore(
    template: TemplateInfo,
    context: any,
    _criteria: TemplateSelectionCriteria
  ): number {
    let score = 0;

    // カテゴリマッチング
    if (context.suggestedCategory === template.category) {
      score += 0.3;
    }

    // 目的マッチング
    if (context.detectedPurpose === template.metadata.purpose) {
      score += 0.25;
    }

    // 聴衆レベルマッチング
    if (context.audienceLevel === template.metadata.targetAudience) {
      score += 0.2;
    }

    // 人気度
    const usage = this.usageStats.get(template.id);
    if (usage && usage.usageCount > 0) {
      score += Math.min(0.15, usage.usageCount * 0.01);
    }

    // 成功率
    if (usage && usage.successRate) {
      score += usage.successRate * 0.1;
    }

    return Math.min(1.0, score);
  }

  private async generateRecommendations(
    scoredTemplates: Array<{ template: TemplateInfo; score: number }>,
    context: any
  ): Promise<TemplateRecommendation[]> {
    return scoredTemplates.map(({ template, score }) => ({
      template,
      score,
      reasoning: this.generateReasoningForTemplate(template, context, score),
      adaptations: this.suggestAdaptations(template, context)
    }));
  }

  private generateReasoningForTemplate(
    template: TemplateInfo,
    context: any,
    score: number
  ): string[] {
    const reasons: string[] = [];

    if (context.suggestedCategory === template.category) {
      reasons.push(`${template.category}カテゴリにマッチしています`);
    }

    if (context.detectedPurpose === template.metadata.purpose) {
      reasons.push(`${template.metadata.purpose}の目的に適しています`);
    }

    if (score > 0.8) {
      reasons.push('高い適合度を示しています');
    } else if (score > 0.6) {
      reasons.push('良好な適合度があります');
    }

    const usage = this.usageStats.get(template.id);
    if (usage && usage.usageCount > 10) {
      reasons.push('多くのユーザーに使用されています');
    }

    return reasons;
  }

  private suggestAdaptations(template: TemplateInfo, _context: any): any[] {
    // 基本的な適応提案
    return [
      {
        type: 'style',
        description: `${template.metadata.presentationStyle}スタイルに調整`,
        confidence: 0.8,
        changes: []
      }
    ];
  }

  private extractKeywords(userInput: string): string[] {
    // 簡単なキーワード抽出
    return userInput.toLowerCase().split(/\s+/).filter(word => word.length > 2);
  }

  private suggestCategory(userInput: string): TemplateCategory {
    const input = userInput.toLowerCase();
    
    if (input.includes('ビジネス') || input.includes('提案') || input.includes('営業')) {
      return 'business';
    }
    if (input.includes('学術') || input.includes('研究') || input.includes('論文')) {
      return 'academic';
    }
    if (input.includes('マーケティング') || input.includes('広告') || input.includes('宣伝')) {
      return 'marketing';
    }
    if (input.includes('技術') || input.includes('エンジニア') || input.includes('開発')) {
      return 'technical';
    }
    
    return 'business'; // デフォルト
  }

  private estimateComplexity(userInput: string): 'simple' | 'moderate' | 'complex' {
    const wordCount = userInput.split(/\s+/).length;
    
    if (wordCount < 50) return 'simple';
    if (wordCount < 200) return 'moderate';
    return 'complex';
  }

  private detectPurpose(userInput: string): string {
    const input = userInput.toLowerCase();
    
    if (input.includes('提案') || input.includes('ピッチ') || input.includes('営業')) {
      return 'pitch';
    }
    if (input.includes('報告') || input.includes('レポート') || input.includes('結果')) {
      return 'report';
    }
    if (input.includes('研修') || input.includes('トレーニング') || input.includes('教育')) {
      return 'training';
    }
    if (input.includes('分析') || input.includes('データ') || input.includes('統計')) {
      return 'analysis';
    }
    
    return 'report'; // デフォルト
  }

  private detectAudienceLevel(userInput: string): string {
    const input = userInput.toLowerCase();
    
    if (input.includes('経営') || input.includes('役員') || input.includes('マネージャー')) {
      return 'executive';
    }
    if (input.includes('技術') || input.includes('エンジニア') || input.includes('専門')) {
      return 'technical';
    }
    if (input.includes('学術') || input.includes('研究') || input.includes('大学')) {
      return 'academic';
    }
    
    return 'general'; // デフォルト
  }

  private async generateTemplateSpecificContent(
    slide: any,
    template: TemplateInfo,
    openAIService: any
  ): Promise<any> {
    // テンプレート特化のプロンプト生成
    const prompt = this.buildTemplateSpecificPrompt(slide, template);
    
    try {
      const response = await openAIService.generateContent(prompt);
      return this.parseTemplateOptimizedResponse(response, slide, template);
    } catch (error) {
      console.error('Template-specific content generation failed:', error);
      return this.createFallbackContent(slide);
    }
  }

  private buildTemplateSpecificPrompt(slide: any, template: TemplateInfo): string {
    return `
${template.name}テンプレート用にスライドコンテンツを最適化してください。

テンプレート仕様:
- カテゴリ: ${template.category}
- スタイル: ${template.metadata.presentationStyle}
- 対象聴衆: ${template.metadata.targetAudience}
- 目的: ${template.metadata.purpose}
- コンテンツ密度: ${template.metadata.contentDensity}

現在のスライド:
${JSON.stringify(slide.adaptedContent, null, 2)}

最適化要求:
1. テンプレートの${template.metadata.presentationStyle}スタイルに準拠
2. ${template.metadata.targetAudience}向けの適切な表現レベル
3. ${template.metadata.contentDensity}密度でのコンテンツ調整
4. ${template.metadata.purpose}目的に最適化

JSON形式で回答:
{
  "title": "最適化されたタイトル",
  "content": ["最適化されたコンテンツ項目"],
  "speakerNotes": "発表者ノート",
  "templateOptimizations": ["適用された最適化の説明"]
}
`;
  }

  private parseTemplateOptimizedResponse(response: string, slide: any, template: TemplateInfo): any {
    try {
      const parsed = JSON.parse(response);
      return {
        ...slide.adaptedContent,
        title: parsed.title || slide.adaptedContent.title,
        content: parsed.content || slide.adaptedContent.content,
        speakerNotes: parsed.speakerNotes || '',
        templateOptimizations: parsed.templateOptimizations || [],
        appliedTemplate: template.name
      };
    } catch (error) {
      console.error('Failed to parse template-optimized response:', error);
      return this.createFallbackContent(slide);
    }
  }

  private createFallbackContent(slide: any): any {
    return {
      ...slide.adaptedContent,
      speakerNotes: 'テンプレート最適化でエラーが発生しました。基本コンテンツを使用しています。',
      templateOptimizations: [],
      appliedTemplate: 'fallback'
    };
  }

  private getDefaultRecommendations(): TemplateRecommendation[] {
    const defaultTemplate = this.templateLibrary.templates[0];
    if (!defaultTemplate) {
      return [];
    }

    return [{
      template: defaultTemplate,
      score: 0.5,
      reasoning: ['デフォルトテンプレートです'],
      adaptations: []
    }];
  }

  private getTemplateById(id: string): TemplateInfo | undefined {
    return this.templateLibrary.templates.find(t => t.id === id);
  }

  private updateTemplateUsage(templateId: string): void {
    const stats = this.usageStats.get(templateId);
    if (stats) {
      stats.usageCount++;
      stats.lastUsed = new Date();
      this.saveUsageStats();
    }
  }

  private initializeTemplateLibrary(): TemplateLibrary {
    // デフォルトテンプレートライブラリの初期化
    const defaultTemplates: TemplateInfo[] = [
      this.createDefaultBusinessTemplate(),
      this.createDefaultAcademicTemplate(),
      this.createDefaultMinimalTemplate()
    ];

    return {
      templates: defaultTemplates,
      categories: this.categorizeTemplates(defaultTemplates),
      searchIndex: this.buildSearchIndex(defaultTemplates),
      statistics: this.calculateLibraryStats(defaultTemplates)
    };
  }

  private createDefaultBusinessTemplate(): TemplateInfo {
    return {
      id: 'default-business',
      name: 'ビジネス標準',
      description: 'ビジネスプレゼンテーション用の標準テンプレート',
      category: 'business',
      metadata: {
        presentationStyle: 'formal',
        targetAudience: 'executive',
        slideCount: 10,
        colorSchemeType: 'corporate',
        layoutComplexity: 'moderate',
        contentDensity: 'medium',
        purpose: 'pitch',
        tags: ['business', 'formal', 'corporate'],
        registeredAt: new Date(),
        usageCount: 0
      },
      designPatterns: [],
      structure: {
        expectedSlideTypes: [
          { position: 'first', type: 'title', frequency: 1, required: true, variations: [] },
          { position: 'any', type: 'content', frequency: 0.8, required: true, variations: [] }
        ],
        contentFlow: [],
        navigationPattern: {
          hasAgenda: true,
          hasTableOfContents: false,
          sectionDividers: true,
          backNavigation: false
        },
        visualHierarchy: []
      },
      compatibility: {
        powerPointVersion: ['2019', '365'],
        supportedFeatures: ['themes', 'layouts'],
        limitations: [],
        apiRequirements: ['PowerPoint.js']
      }
    };
  }

  private createDefaultAcademicTemplate(): TemplateInfo {
    return {
      id: 'default-academic',
      name: '学術発表',
      description: '学術発表・研究発表用テンプレート',
      category: 'academic',
      metadata: {
        presentationStyle: 'formal',
        targetAudience: 'academic',
        slideCount: 15,
        colorSchemeType: 'academic',
        layoutComplexity: 'moderate',
        contentDensity: 'high',
        purpose: 'report',
        tags: ['academic', 'research', 'formal'],
        registeredAt: new Date(),
        usageCount: 0
      },
      designPatterns: [],
      structure: {
        expectedSlideTypes: [
          { position: 'first', type: 'title', frequency: 1, required: true, variations: [] },
          { position: 2, type: 'agenda', frequency: 1, required: true, variations: [] },
          { position: 'any', type: 'content', frequency: 0.9, required: true, variations: [] }
        ],
        contentFlow: [],
        navigationPattern: {
          hasAgenda: true,
          hasTableOfContents: true,
          sectionDividers: true,
          backNavigation: true
        },
        visualHierarchy: []
      },
      compatibility: {
        powerPointVersion: ['2019', '365'],
        supportedFeatures: ['themes', 'layouts'],
        limitations: [],
        apiRequirements: ['PowerPoint.js']
      }
    };
  }

  private createDefaultMinimalTemplate(): TemplateInfo {
    return {
      id: 'default-minimal',
      name: 'ミニマル',
      description: 'シンプルで洗練されたミニマルテンプレート',
      category: 'minimal',
      metadata: {
        presentationStyle: 'casual',
        targetAudience: 'general',
        slideCount: 8,
        colorSchemeType: 'minimal',
        layoutComplexity: 'simple',
        contentDensity: 'low',
        purpose: 'pitch',
        tags: ['minimal', 'simple', 'clean'],
        registeredAt: new Date(),
        usageCount: 0
      },
      designPatterns: [],
      structure: {
        expectedSlideTypes: [
          { position: 'first', type: 'title', frequency: 1, required: true, variations: [] },
          { position: 'any', type: 'content', frequency: 0.7, required: true, variations: [] }
        ],
        contentFlow: [],
        navigationPattern: {
          hasAgenda: false,
          hasTableOfContents: false,
          sectionDividers: false,
          backNavigation: false
        },
        visualHierarchy: []
      },
      compatibility: {
        powerPointVersion: ['2019', '365'],
        supportedFeatures: ['themes', 'layouts'],
        limitations: [],
        apiRequirements: ['PowerPoint.js']
      }
    };
  }

  private categorizeTemplates(templates: TemplateInfo[]): Record<TemplateCategory, TemplateInfo[]> {
    const categories: Record<TemplateCategory, TemplateInfo[]> = {
      business: [],
      academic: [],
      creative: [],
      technical: [],
      marketing: [],
      corporate: [],
      minimal: [],
      custom: []
    };

    templates.forEach(template => {
      categories[template.category].push(template);
    });

    return categories;
  }

  private buildSearchIndex(_templates: TemplateInfo[]): any {
    // 簡単な検索インデックス構築
    return {
      byIndustry: {},
      byStyle: {},
      byPurpose: {},
      byTags: {}
    };
  }

  private calculateLibraryStats(templates: TemplateInfo[]): any {
    return {
      totalTemplates: templates.length,
      byCategory: this.countByCategory(templates),
      mostUsed: [],
      recentlyAdded: templates.slice(-3).map(t => t.id),
      averageScore: 0
    };
  }

  private countByCategory(templates: TemplateInfo[]): Record<TemplateCategory, number> {
    const counts: Record<TemplateCategory, number> = {
      business: 0,
      academic: 0,
      creative: 0,
      technical: 0,
      marketing: 0,
      corporate: 0,
      minimal: 0,
      custom: 0
    };

    templates.forEach(template => {
      counts[template.category]++;
    });

    return counts;
  }

  private addTemplateToLibrary(template: TemplateInfo): void {
    this.templateLibrary.templates.push(template);
    this.templateLibrary.categories[template.category].push(template);
    this.templateLibrary.statistics = this.calculateLibraryStats(this.templateLibrary.templates);
  }

  private async registerTemplateFromFile(_request: TemplateRegistrationRequest): Promise<TemplateInfo> {
    // ファイルベースの登録（将来的な実装）
    throw new Error('File-based template registration not yet implemented');
  }

  private async registerCurrentPresentationAsTemplate(request: TemplateRegistrationRequest): Promise<TemplateInfo> {
    const detectedTemplate = await this.adaptationService.detectTemplate();
    
    if (!detectedTemplate) {
      throw new Error('現在のプレゼンテーションからテンプレートを検出できませんでした');
    }

    // ユーザー提供のメタデータで更新
    detectedTemplate.metadata = {
      ...detectedTemplate.metadata,
      ...request.metadata
    };

    return detectedTemplate;
  }

  private loadUsageStats(): void {
    try {
      const saved = localStorage.getItem('template-usage-stats');
      if (saved) {
        const parsed = JSON.parse(saved);
        this.usageStats = new Map(Object.entries(parsed));
      }
    } catch (error) {
      console.error('Failed to load usage stats:', error);
    }
  }

  private saveUsageStats(): void {
    try {
      const statsObject = Object.fromEntries(this.usageStats);
      localStorage.setItem('template-usage-stats', JSON.stringify(statsObject));
    } catch (error) {
      console.error('Failed to save usage stats:', error);
    }
  }

  private async saveTemplateLibrary(): Promise<void> {
    try {
      localStorage.setItem('template-library', JSON.stringify(this.templateLibrary));
    } catch (error) {
      console.error('Failed to save template library:', error);
    }
  }
}