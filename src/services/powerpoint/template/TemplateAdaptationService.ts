// TemplateAdaptationService.ts
// PowerPointテンプレート検出・適応サービス

import {
  TemplateInfo,
  TemplateAnalysisResult,
  DesignPattern,
  TemplateStructure,
  TemplateMetadata,
  AdaptedOutline,
  TemplateAdaptation,
  AdaptedSlide
} from '../template-types';
import { ThemeService } from '../theme/ThemeService';
import { ThemeAnalyzer } from '../theme/ThemeAnalyzer';

export class TemplateAdaptationService {
  private themeService: ThemeService;
  private themeAnalyzer: ThemeAnalyzer;

  constructor() {
    this.themeService = new ThemeService();
    this.themeAnalyzer = new ThemeAnalyzer();
  }

  /**
   * 現在開いているPowerPointプレゼンテーションからテンプレート情報を検出
   */
  async detectTemplate(): Promise<TemplateInfo | null> {
    try {
      return await PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const slides = presentation.slides;
        
        slides.load(['items']);
        await context.sync();

        if (slides.items.length === 0) {
          return null;
        }

        // 最初の数スライドを分析してテンプレート特性を抽出
        const analysisResult = await this.analyzeSlideStructures(context, slides.items.slice(0, 5));
        
        // テーマ情報も取得
        const themeInfo = await this.themeService.getCurrentThemeInfo();
        
        return this.buildTemplateInfo(analysisResult, themeInfo);
      });
    } catch (error) {
      console.error('Template detection failed:', error);
      return null;
    }
  }

  /**
   * テンプレートファイルまたはプレゼンテーションを分析してデザインパターンを抽出
   */
  async analyzeTemplateStructure(templateFile?: File): Promise<TemplateAnalysisResult> {
    try {
      if (templateFile) {
        // ファイルベースの分析（将来的な拡張）
        return await this.analyzeTemplateFile(templateFile);
      } else {
        // 現在のプレゼンテーションの分析
        return await this.analyzeCurrentPresentation();
      }
    } catch (error) {
      console.error('Template analysis failed:', error);
      return {
        detectedPatterns: [],
        extractedStructure: this.getDefaultTemplateStructure(),
        suggestedMetadata: {},
        confidence: 0,
        analysisNotes: [`Analysis failed: ${error.message}`]
      };
    }
  }

  /**
   * 既存のアウトラインをテンプレートに適応
   */
  async adaptOutlineToTemplate(
    outline: any,
    template: TemplateInfo
  ): Promise<AdaptedOutline> {
    try {
      const adaptations: TemplateAdaptation[] = [];
      const adaptedSlides: AdaptedSlide[] = [];

      // アウトラインの各スライドをテンプレートパターンに適応
      for (let i = 0; i < outline.slides.length; i++) {
        const slide = outline.slides[i];
        const adaptedSlide = await this.adaptSlideToTemplate(slide, template, i);
        adaptedSlides.push(adaptedSlide);
      }

      // 全体的な構造適応
      const structuralAdaptations = this.analyzeStructuralAdaptations(outline, template);
      adaptations.push(...structuralAdaptations);

      return {
        originalOutline: outline,
        selectedTemplate: template,
        adaptations,
        adaptedSlides,
        confidence: this.calculateAdaptationConfidence(adaptations)
      };
    } catch (error) {
      console.error('Outline adaptation failed:', error);
      throw error;
    }
  }

  /**
   * テンプレートに最適化されたコンテンツを生成
   */
  async generateAdaptedContent(
    adaptedOutline: AdaptedOutline,
    openAIService: any
  ): Promise<AdaptedSlide[]> {
    const enhancedSlides: AdaptedSlide[] = [];

    for (const slide of adaptedOutline.adaptedSlides) {
      try {
        const enhancedContent = await this.enhanceSlideContentForTemplate(
          slide,
          adaptedOutline.selectedTemplate,
          openAIService
        );
        
        enhancedSlides.push({
          ...slide,
          adaptedContent: enhancedContent
        });
      } catch (error) {
        console.error(`Failed to enhance slide ${slide.slideNumber}:`, error);
        // フォールバック: 元のコンテンツを使用
        enhancedSlides.push(slide);
      }
    }

    return enhancedSlides;
  }

  private async analyzeCurrentPresentation(): Promise<TemplateAnalysisResult> {
    return await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      
      slides.load(['items']);
      await context.sync();

      const patterns = await this.extractDesignPatterns(context, slides.items);
      const structure = await this.extractTemplateStructure(context, slides.items);
      const metadata = await this.suggestMetadata(patterns, structure);

      return {
        detectedPatterns: patterns,
        extractedStructure: structure,
        suggestedMetadata: metadata,
        confidence: this.calculateAnalysisConfidence(patterns, structure),
        analysisNotes: [
          `Analyzed ${slides.items.length} slides`,
          `Detected ${patterns.length} design patterns`,
          `Identified ${structure.expectedSlideTypes.length} slide type patterns`
        ]
      };
    });
  }

  private async analyzeSlideStructures(
    context: PowerPoint.RequestContext,
    slides: PowerPoint.Slide[]
  ): Promise<TemplateAnalysisResult> {
    const patterns: DesignPattern[] = [];
    const slideTypeFrequency: Record<string, number> = {};

    for (const slide of slides) {
      slide.load(['layout', 'shapes']);
      await context.sync();

      // レイアウトパターンの分析
      const layoutPattern = this.analyzeSlideLayout(slide);
      if (layoutPattern) patterns.push(layoutPattern);

      // スライドタイプの頻度計算
      const slideType = this.detectSlideType(slide);
      slideTypeFrequency[slideType] = (slideTypeFrequency[slideType] || 0) + 1;
    }

    const structure = this.buildTemplateStructure(slideTypeFrequency, slides.length);

    return {
      detectedPatterns: patterns,
      extractedStructure: structure,
      suggestedMetadata: this.inferMetadataFromPatterns(patterns),
      confidence: patterns.length > 0 ? 0.8 : 0.3,
      analysisNotes: [`Analyzed ${slides.length} slides for template patterns`]
    };
  }

  private analyzeSlideLayout(slide: PowerPoint.Slide): DesignPattern | null {
    try {
      // スライドレイアウトの分析ロジック
      // 実際の実装では slide.layout や slide.shapes を分析
      const layoutName = slide.layout?.name || 'unknown';
      
      return {
        type: 'layout',
        pattern: `layout-${layoutName.toLowerCase().replace(/\s+/g, '-')}`,
        frequency: 1,
        importance: 'important',
        description: `Layout pattern for ${layoutName}`,
        rules: [
          {
            condition: 'has title placeholder',
            action: 'apply title formatting',
            priority: 1
          }
        ]
      };
    } catch (error) {
      console.error('Layout analysis failed:', error);
      return null;
    }
  }

  private detectSlideType(slide: PowerPoint.Slide): string {
    try {
      // 簡単なスライドタイプ検出ロジック
      // 実際の実装では slide.shapes を分析してより詳細な検出を行う
      const layoutName = slide.layout?.name?.toLowerCase() || '';
      
      if (layoutName.includes('title')) return 'title';
      if (layoutName.includes('comparison') || layoutName.includes('two')) return 'comparison';
      if (layoutName.includes('blank')) return 'blank';
      
      return 'content';
    } catch (error) {
      console.error('Slide type detection failed:', error);
      return 'unknown';
    }
  }

  private buildTemplateStructure(
    slideTypeFrequency: Record<string, number>,
    totalSlides: number
  ): TemplateStructure {
    const expectedSlideTypes = Object.entries(slideTypeFrequency).map(([type, count]) => ({
      position: 'any' as const,
      type: type as any,
      frequency: count / totalSlides,
      required: count > totalSlides * 0.5,
      variations: []
    }));

    return {
      expectedSlideTypes,
      contentFlow: [],
      navigationPattern: {
        hasAgenda: false,
        hasTableOfContents: false,
        sectionDividers: false,
        backNavigation: false
      },
      visualHierarchy: []
    };
  }

  private async adaptSlideToTemplate(
    slide: any,
    template: TemplateInfo,
    index: number
  ): Promise<AdaptedSlide> {
    // スライドをテンプレートパターンに適応
    const relevantPatterns = template.designPatterns.filter(p => 
      p.type === 'layout' || p.type === 'typography'
    );

    const adaptedContent = {
      title: slide.title || `スライド ${index + 1}`,
      content: slide.content || slide.keyPoints || [],
      slideType: this.mapSlideTypeToTemplate(slide.contentType, template),
      layoutSuggestion: this.suggestLayoutForSlide(slide, template),
      styleOverrides: this.generateStyleOverrides(slide, template)
    };

    return {
      slideNumber: index + 1,
      originalSlide: slide,
      adaptedContent,
      templatePatterns: relevantPatterns,
      adaptationNotes: [
        `Adapted to ${template.name} template`,
        `Applied ${relevantPatterns.length} template patterns`
      ]
    };
  }

  private mapSlideTypeToTemplate(contentType: string, _template: TemplateInfo): string {
    // テンプレートの期待するスライドタイプにマッピング
    const mapping: Record<string, string> = {
      'title': 'title',
      'bullets': 'content',
      'comparison': 'comparison',
      'conclusion': 'conclusion'
    };

    return mapping[contentType] || 'content';
  }

  private suggestLayoutForSlide(slide: any, _template: TemplateInfo): string {
    // テンプレートの構造に基づいてレイアウトを提案
    const contentAmount = (slide.content?.join(' ') || '').length;
    
    if (slide.contentType === 'title') return 'Title Slide';
    if (slide.contentType === 'comparison') return 'Two Content';
    if (contentAmount > 500) return 'Content with Caption';
    
    return 'Title and Content';
  }

  private generateStyleOverrides(_slide: any, template: TemplateInfo): Record<string, any> {
    // テンプレートスタイルに基づいたオーバーライド
    const overrides: Record<string, any> = {};

    // テンプレートの視覚的階層を適用
    if (template.structure.visualHierarchy.length > 0) {
      const titleStyle = template.structure.visualHierarchy.find(h => h.element === 'title');
      if (titleStyle) {
        overrides.titleStyle = {
          fontSize: titleStyle.fontSize,
          fontWeight: titleStyle.fontWeight,
          color: titleStyle.color
        };
      }
    }

    return overrides;
  }

  private analyzeStructuralAdaptations(
    outline: any,
    template: TemplateInfo
  ): TemplateAdaptation[] {
    const adaptations: TemplateAdaptation[] = [];

    // スライド数の調整提案
    if (outline.slides.length !== template.metadata.slideCount) {
      adaptations.push({
        type: 'structure',
        description: `Recommend ${template.metadata.slideCount} slides (current: ${outline.slides.length})`,
        confidence: 0.7,
        changes: [{
          target: 'slide_count',
          from: outline.slides.length,
          to: template.metadata.slideCount,
          reason: 'Template optimized for specific slide count'
        }]
      });
    }

    // コンテンツ密度の調整
    if (template.metadata.contentDensity === 'low') {
      adaptations.push({
        type: 'content',
        description: 'Reduce content density to match template style',
        confidence: 0.8,
        changes: [{
          target: 'content_density',
          from: 'current',
          to: 'low',
          reason: 'Template designed for minimal content per slide'
        }]
      });
    }

    return adaptations;
  }

  private async enhanceSlideContentForTemplate(
    slide: AdaptedSlide,
    template: TemplateInfo,
    openAIService: any
  ): Promise<any> {
    const prompt = this.buildTemplateOptimizedPrompt(slide, template);
    
    try {
      const response = await openAIService.generateContent(prompt);
      return this.parseEnhancedContent(response, slide);
    } catch (error) {
      console.error('Template-optimized content generation failed:', error);
      return slide.adaptedContent;
    }
  }

  private buildTemplateOptimizedPrompt(slide: AdaptedSlide, template: TemplateInfo): string {
    return `
以下のスライドコンテンツを、${template.name}テンプレートのスタイルに最適化してください。

テンプレート特性:
- スタイル: ${template.metadata.presentationStyle}
- 対象聴衆: ${template.metadata.targetAudience}
- コンテンツ密度: ${template.metadata.contentDensity}
- 目的: ${template.metadata.purpose}

現在のスライド:
タイトル: ${slide.adaptedContent.title}
コンテンツ: ${slide.adaptedContent.content.join(', ')}

最適化要件:
1. テンプレートの${template.metadata.presentationStyle}スタイルに合わせる
2. ${template.metadata.contentDensity}コンテンツ密度を維持
3. ${template.metadata.targetAudience}向けの表現を使用

JSON形式で以下を返してください:
{
  "title": "最適化されたタイトル",
  "content": ["最適化されたコンテンツ項目の配列"],
  "notes": "最適化の説明"
}
`;
  }

  private parseEnhancedContent(response: string, fallbackSlide: AdaptedSlide): any {
    try {
      const parsed = JSON.parse(response);
      return {
        title: parsed.title || fallbackSlide.adaptedContent.title,
        content: parsed.content || fallbackSlide.adaptedContent.content,
        slideType: fallbackSlide.adaptedContent.slideType,
        layoutSuggestion: fallbackSlide.adaptedContent.layoutSuggestion,
        styleOverrides: fallbackSlide.adaptedContent.styleOverrides,
        enhancementNotes: parsed.notes || ''
      };
    } catch (error) {
      console.error('Failed to parse enhanced content:', error);
      return fallbackSlide.adaptedContent;
    }
  }

  private calculateAdaptationConfidence(adaptations: TemplateAdaptation[]): number {
    if (adaptations.length === 0) return 1.0;
    
    const avgConfidence = adaptations.reduce((sum, a) => sum + a.confidence, 0) / adaptations.length;
    return Math.max(0.1, Math.min(1.0, avgConfidence));
  }

  private async extractDesignPatterns(
    _context: PowerPoint.RequestContext,
    _slides: PowerPoint.Slide[]
  ): Promise<DesignPattern[]> {
    const patterns: DesignPattern[] = [];
    
    // 基本的なパターン抽出ロジック
    // 実際の実装では、より詳細な分析を行う
    
    return patterns;
  }

  private async extractTemplateStructure(
    _context: PowerPoint.RequestContext,
    _slides: PowerPoint.Slide[]
  ): Promise<TemplateStructure> {
    return this.getDefaultTemplateStructure();
  }

  private async suggestMetadata(
    _patterns: DesignPattern[],
    _structure: TemplateStructure
  ): Promise<Partial<TemplateMetadata>> {
    return {
      presentationStyle: 'formal',
      targetAudience: 'general',
      contentDensity: 'medium',
      purpose: 'report'
    };
  }

  private calculateAnalysisConfidence(
    patterns: DesignPattern[],
    structure: TemplateStructure
  ): number {
    return patterns.length > 0 && structure.expectedSlideTypes.length > 0 ? 0.8 : 0.3;
  }

  private getDefaultTemplateStructure(): TemplateStructure {
    return {
      expectedSlideTypes: [{
        position: 'any',
        type: 'content',
        frequency: 1,
        required: true,
        variations: []
      }],
      contentFlow: [],
      navigationPattern: {
        hasAgenda: false,
        hasTableOfContents: false,
        sectionDividers: false,
        backNavigation: false
      },
      visualHierarchy: []
    };
  }

  private buildTemplateInfo(
    analysisResult: TemplateAnalysisResult,
    themeInfo: any
  ): TemplateInfo {
    return {
      id: `template_${Date.now()}`,
      name: themeInfo?.name || 'Detected Template',
      description: 'Auto-detected from current presentation',
      category: 'custom',
      metadata: {
        presentationStyle: 'formal',
        targetAudience: 'general',
        slideCount: 0,
        colorSchemeType: 'corporate',
        layoutComplexity: 'moderate',
        contentDensity: 'medium',
        purpose: 'report',
        tags: ['auto-detected'],
        registeredAt: new Date(),
        usageCount: 0,
        ...analysisResult.suggestedMetadata
      },
      designPatterns: analysisResult.detectedPatterns,
      structure: analysisResult.extractedStructure,
      compatibility: {
        powerPointVersion: ['2019', '365'],
        supportedFeatures: ['basic-layouts', 'themes'],
        limitations: ['no-animations'],
        apiRequirements: ['PowerPoint.js']
      }
    };
  }

  private inferMetadataFromPatterns(patterns: DesignPattern[]): Partial<TemplateMetadata> {
    // パターンからメタデータを推測
    return {
      layoutComplexity: patterns.length > 5 ? 'complex' : 'simple',
      contentDensity: 'medium'
    };
  }

  private async analyzeTemplateFile(_file: File): Promise<TemplateAnalysisResult> {
    // ファイルベースの分析（将来的な実装）
    return {
      detectedPatterns: [],
      extractedStructure: this.getDefaultTemplateStructure(),
      suggestedMetadata: {},
      confidence: 0.5,
      analysisNotes: ['File-based analysis not yet implemented']
    };
  }
}