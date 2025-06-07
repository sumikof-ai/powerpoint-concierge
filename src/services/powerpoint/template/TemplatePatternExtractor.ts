// TemplatePatternExtractor.ts
// テンプレートのメタデータとパターン抽出システム

import {
  DesignPattern,
  TemplateStructure,
  TemplateMetadata,
  VisualHierarchyPattern,
  ContentFlowPattern,
  SlideTypePattern
} from '../template-types';

export class TemplatePatternExtractor {
  
  /**
   * スライドからデザインパターンを抽出
   */
  async extractPatternsFromSlides(slides: PowerPoint.Slide[]): Promise<DesignPattern[]> {
    const patterns: DesignPattern[] = [];
    
    try {
      await PowerPoint.run(async (context) => {
        // スライドデータの読み込み
        slides.forEach(slide => {
          slide.load(['layout', 'shapes', 'background']);
        });
        await context.sync();

        // 各種パターンの抽出
        patterns.push(...await this.extractLayoutPatterns(slides));
        patterns.push(...await this.extractColorPatterns(slides));
        patterns.push(...await this.extractTypographyPatterns(slides));
        patterns.push(...await this.extractSpacingPatterns(slides));
      });
    } catch (error) {
      console.error('Pattern extraction failed:', error);
    }

    return patterns;
  }

  /**
   * テンプレート構造の分析と抽出
   */
  async extractTemplateStructure(slides: PowerPoint.Slide[]): Promise<TemplateStructure> {
    try {
      const slideTypePatterns = await this.analyzeSlideTypePatterns(slides);
      const contentFlow = await this.analyzeContentFlow(slides);
      const navigationPattern = await this.analyzeNavigationPattern(slides);
      const visualHierarchy = await this.analyzeVisualHierarchy(slides);

      return {
        expectedSlideTypes: slideTypePatterns,
        contentFlow,
        navigationPattern,
        visualHierarchy
      };
    } catch (error) {
      console.error('Template structure extraction failed:', error);
      return this.getDefaultTemplateStructure();
    }
  }

  /**
   * メタデータの自動推測
   */
  async inferTemplateMetadata(
    patterns: DesignPattern[],
    structure: TemplateStructure,
    slides: PowerPoint.Slide[]
  ): Promise<Partial<TemplateMetadata>> {
    const metadata: Partial<TemplateMetadata> = {
      slideCount: slides.length,
      registeredAt: new Date(),
      usageCount: 0,
      tags: []
    };

    // プレゼンテーションスタイルの推測
    metadata.presentationStyle = this.inferPresentationStyle(patterns, structure);
    
    // 対象聴衆の推測
    metadata.targetAudience = this.inferTargetAudience(patterns, structure);
    
    // カラースキームタイプの推測
    metadata.colorSchemeType = this.inferColorSchemeType(patterns);
    
    // レイアウト複雑度の推測
    metadata.layoutComplexity = this.inferLayoutComplexity(patterns, structure);
    
    // コンテンツ密度の推測
    metadata.contentDensity = this.inferContentDensity(patterns, structure);
    
    // 目的の推測
    metadata.purpose = this.inferPurpose(patterns, structure);
    
    // タグの生成
    metadata.tags = this.generateTags(patterns, structure, metadata);

    return metadata;
  }

  /**
   * レイアウトパターンの抽出
   */
  private async extractLayoutPatterns(slides: PowerPoint.Slide[]): Promise<DesignPattern[]> {
    const patterns: DesignPattern[] = [];
    const layoutFrequency: Record<string, number> = {};

    for (const slide of slides) {
      try {
        const layoutName = await this.getSlideLayoutName(slide);
        layoutFrequency[layoutName] = (layoutFrequency[layoutName] || 0) + 1;
      } catch (error) {
        console.error('Layout pattern extraction failed for slide:', error);
      }
    }

    // 頻出レイアウトをパターンとして登録
    Object.entries(layoutFrequency).forEach(([layout, frequency]) => {
      if (frequency >= 2) { // 2回以上使用されたレイアウト
        patterns.push({
          type: 'layout',
          pattern: layout,
          frequency: frequency / slides.length,
          importance: frequency > slides.length * 0.5 ? 'critical' : 'important',
          description: `${layout} layout used in ${frequency} slides`,
          rules: [{
            condition: `slide requires ${layout} layout`,
            action: `apply ${layout} layout`,
            priority: 1
          }]
        });
      }
    });

    return patterns;
  }

  /**
   * カラーパターンの抽出
   */
  private async extractColorPatterns(slides: PowerPoint.Slide[]): Promise<DesignPattern[]> {
    const patterns: DesignPattern[] = [];
    const colorUsage: Record<string, number> = {};

    for (const slide of slides) {
      try {
        const colors = await this.extractColorsFromSlide(slide);
        colors.forEach(color => {
          colorUsage[color] = (colorUsage[color] || 0) + 1;
        });
      } catch (error) {
        console.error('Color pattern extraction failed for slide:', error);
      }
    }

    // 主要カラーをパターンとして登録
    Object.entries(colorUsage)
      .sort(([, a], [, b]) => b - a)
      .slice(0, 5) // 上位5色
      .forEach(([color, frequency]) => {
        patterns.push({
          type: 'color',
          pattern: color,
          frequency: frequency / slides.length,
          importance: frequency > slides.length * 0.7 ? 'critical' : 'important',
          description: `Primary color ${color} used in ${frequency} slides`,
          rules: [{
            condition: `element needs primary color`,
            action: `apply color ${color}`,
            priority: 1
          }]
        });
      });

    return patterns;
  }

  /**
   * タイポグラフィパターンの抽出
   */
  private async extractTypographyPatterns(slides: PowerPoint.Slide[]): Promise<DesignPattern[]> {
    const patterns: DesignPattern[] = [];
    const fontUsage: Record<string, number> = {};
    const sizeUsage: Record<number, number> = {};

    for (const slide of slides) {
      try {
        const typography = await this.extractTypographyFromSlide(slide);
        
        typography.fonts.forEach(font => {
          fontUsage[font] = (fontUsage[font] || 0) + 1;
        });
        
        typography.sizes.forEach(size => {
          sizeUsage[size] = (sizeUsage[size] || 0) + 1;
        });
      } catch (error) {
        console.error('Typography pattern extraction failed for slide:', error);
      }
    }

    // フォントパターン
    Object.entries(fontUsage)
      .sort(([, a], [, b]) => b - a)
      .slice(0, 3)
      .forEach(([font, frequency]) => {
        patterns.push({
          type: 'typography',
          pattern: `font-${font}`,
          frequency: frequency / slides.length,
          importance: 'important',
          description: `Font ${font} used in ${frequency} slides`,
          rules: [{
            condition: `text needs primary font`,
            action: `apply font ${font}`,
            priority: 1
          }]
        });
      });

    return patterns;
  }

  /**
   * スペーシングパターンの抽出
   */
  private async extractSpacingPatterns(_slides: PowerPoint.Slide[]): Promise<DesignPattern[]> {
    const patterns: DesignPattern[] = [];
    
    // 簡単なスペーシング分析（実際の実装ではより詳細な分析を行う）
    patterns.push({
      type: 'spacing',
      pattern: 'standard-margins',
      frequency: 1.0,
      importance: 'important',
      description: 'Standard margin spacing pattern',
      rules: [{
        condition: 'content needs margin',
        action: 'apply standard margins',
        priority: 1
      }]
    });

    return patterns;
  }

  /**
   * スライドタイプパターンの分析
   */
  private async analyzeSlideTypePatterns(slides: PowerPoint.Slide[]): Promise<SlideTypePattern[]> {
    const typeFrequency: Record<string, { count: number; positions: number[] }> = {};

    for (let i = 0; i < slides.length; i++) {
      try {
        const slideType = await this.detectSlideType(slides[i]);
        
        if (!typeFrequency[slideType]) {
          typeFrequency[slideType] = { count: 0, positions: [] };
        }
        
        typeFrequency[slideType].count++;
        typeFrequency[slideType].positions.push(i);
      } catch (error) {
        console.error(`Slide type analysis failed for slide ${i}:`, error);
      }
    }

    return Object.entries(typeFrequency).map(([type, data]) => ({
      position: this.determineSlidePosition(data.positions, slides.length),
      type: type as any,
      frequency: data.count / slides.length,
      required: data.count > 0,
      variations: []
    }));
  }

  /**
   * コンテンツフローの分析
   */
  private async analyzeContentFlow(slides: PowerPoint.Slide[]): Promise<ContentFlowPattern[]> {
    const flowPatterns: ContentFlowPattern[] = [];
    
    for (let i = 0; i < slides.length - 1; i++) {
      try {
        const currentType = await this.detectSlideType(slides[i]);
        const nextType = await this.detectSlideType(slides[i + 1]);
        
        flowPatterns.push({
          fromSlideType: currentType,
          toSlideType: nextType,
          probability: 1.0, // 簡単な実装
          transitionStyle: 'standard'
        });
      } catch (error) {
        console.error(`Content flow analysis failed for slides ${i}-${i+1}:`, error);
      }
    }

    return flowPatterns;
  }

  /**
   * ナビゲーションパターンの分析
   */
  private async analyzeNavigationPattern(slides: PowerPoint.Slide[]): Promise<any> {
    let hasAgenda = false;
    let hasTableOfContents = false;
    let sectionDividers = false;

    for (const slide of slides) {
      try {
        const content = await this.extractSlideTextContent(slide);
        const contentLower = content.toLowerCase();
        
        if (contentLower.includes('アジェンダ') || contentLower.includes('agenda')) {
          hasAgenda = true;
        }
        
        if (contentLower.includes('目次') || contentLower.includes('contents')) {
          hasTableOfContents = true;
        }
        
        if (contentLower.includes('セクション') || contentLower.includes('section')) {
          sectionDividers = true;
        }
      } catch (error) {
        console.error('Navigation pattern analysis failed for slide:', error);
      }
    }

    return {
      hasAgenda,
      hasTableOfContents,
      sectionDividers,
      backNavigation: false // 静的分析では困難
    };
  }

  /**
   * 視覚的階層の分析
   */
  private async analyzeVisualHierarchy(_slides: PowerPoint.Slide[]): Promise<VisualHierarchyPattern[]> {
    const hierarchy: VisualHierarchyPattern[] = [];
    
    // 基本的な階層パターン（実際の実装ではより詳細な分析を行う）
    hierarchy.push({
      level: 1,
      element: 'title',
      fontSize: 24,
      fontWeight: 'bold',
      color: '#000000',
      positioning: 'top'
    });

    hierarchy.push({
      level: 2,
      element: 'content',
      fontSize: 16,
      fontWeight: 'normal',
      color: '#333333',
      positioning: 'center'
    });

    return hierarchy;
  }

  /**
   * プレゼンテーションスタイルの推測
   */
  private inferPresentationStyle(
    patterns: DesignPattern[],
    structure: TemplateStructure
  ): 'formal' | 'casual' | 'technical' | 'creative' | 'minimal' {
    const colorPatterns = patterns.filter(p => p.type === 'color');
    const layoutPatterns = patterns.filter(p => p.type === 'layout');
    
    // カラーパターンから推測
    const hasMultipleColors = colorPatterns.length > 3;
    const hasNavigation = structure.navigationPattern.hasAgenda || 
                          structure.navigationPattern.hasTableOfContents;
    
    if (hasNavigation && !hasMultipleColors) {
      return 'formal';
    }
    
    if (hasMultipleColors && layoutPatterns.length > 3) {
      return 'creative';
    }
    
    if (layoutPatterns.length <= 2) {
      return 'minimal';
    }
    
    return 'casual';
  }

  /**
   * 対象聴衆の推測
   */
  private inferTargetAudience(
    _patterns: DesignPattern[],
    structure: TemplateStructure
  ): 'executive' | 'technical' | 'general' | 'academic' | 'sales' {
    const hasComplexNavigation = structure.navigationPattern.hasTableOfContents &&
                                structure.navigationPattern.sectionDividers;
    
    if (hasComplexNavigation) {
      return 'academic';
    }
    
    if (structure.expectedSlideTypes.length <= 2) {
      return 'executive';
    }
    
    return 'general';
  }

  /**
   * カラースキームタイプの推測
   */
  private inferColorSchemeType(
    patterns: DesignPattern[]
  ): 'corporate' | 'vibrant' | 'minimal' | 'academic' | 'creative' {
    const colorPatterns = patterns.filter(p => p.type === 'color');
    
    if (colorPatterns.length <= 2) {
      return 'minimal';
    }
    
    if (colorPatterns.length >= 5) {
      return 'vibrant';
    }
    
    return 'corporate';
  }

  /**
   * レイアウト複雑度の推測
   */
  private inferLayoutComplexity(
    patterns: DesignPattern[],
    structure: TemplateStructure
  ): 'simple' | 'moderate' | 'complex' {
    const layoutPatterns = patterns.filter(p => p.type === 'layout');
    const slideTypes = structure.expectedSlideTypes.length;
    
    if (layoutPatterns.length <= 2 && slideTypes <= 2) {
      return 'simple';
    }
    
    if (layoutPatterns.length >= 5 || slideTypes >= 5) {
      return 'complex';
    }
    
    return 'moderate';
  }

  /**
   * コンテンツ密度の推測
   */
  private inferContentDensity(
    patterns: DesignPattern[],
    structure: TemplateStructure
  ): 'low' | 'medium' | 'high' {
    // パターン数とスライドタイプ数から密度を推測
    const hasMultipleContentTypes = structure.expectedSlideTypes.length > 3;
    const hasComplexPatterns = patterns.filter(p => p.importance === 'critical').length > 2;
    
    if (hasMultipleContentTypes && hasComplexPatterns) {
      return 'high';
    }
    
    if (patterns.length <= 2 && structure.expectedSlideTypes.length <= 2) {
      return 'low';
    }
    
    return 'medium';
  }

  /**
   * 目的の推測
   */
  private inferPurpose(
    _patterns: DesignPattern[],
    structure: TemplateStructure
  ): 'pitch' | 'report' | 'training' | 'marketing' | 'analysis' {
    const hasAgenda = structure.navigationPattern.hasAgenda;
    const slideCount = structure.expectedSlideTypes.reduce((sum, type) => 
      sum + (type.frequency || 0), 0
    );
    
    if (hasAgenda && slideCount > 10) {
      return 'training';
    }
    
    if (slideCount <= 8) {
      return 'pitch';
    }
    
    return 'report';
  }

  /**
   * タグの生成
   */
  private generateTags(
    _patterns: DesignPattern[],
    _structure: TemplateStructure,
    metadata: Partial<TemplateMetadata>
  ): string[] {
    const tags: string[] = [];
    
    if (metadata.presentationStyle) {
      tags.push(metadata.presentationStyle);
    }
    
    if (metadata.targetAudience) {
      tags.push(metadata.targetAudience);
    }
    
    if (metadata.purpose) {
      tags.push(metadata.purpose);
    }
    
    if (metadata.layoutComplexity === 'simple') {
      tags.push('simple', 'clean');
    }
    
    if (metadata.contentDensity === 'low') {
      tags.push('minimal', 'spacious');
    }
    
    tags.push('auto-generated');
    
    return Array.from(new Set(tags)); // 重複除去
  }

  // ヘルパーメソッド
  private async getSlideLayoutName(slide: PowerPoint.Slide): Promise<string> {
    try {
      return slide.layout.name || 'Unknown Layout';
    } catch (error) {
      return 'Unknown Layout';
    }
  }

  private async extractColorsFromSlide(_slide: PowerPoint.Slide): Promise<string[]> {
    // 実際の実装では、スライドから色情報を抽出
    return ['#000000', '#FFFFFF']; // プレースホルダー
  }

  private async extractTypographyFromSlide(_slide: PowerPoint.Slide): Promise<{
    fonts: string[];
    sizes: number[];
  }> {
    // 実際の実装では、スライドからタイポグラフィ情報を抽出
    return {
      fonts: ['Arial', 'Calibri'],
      sizes: [16, 24]
    };
  }

  private async detectSlideType(_slide: PowerPoint.Slide): Promise<string> {
    // 実際の実装では、スライドの内容を分析してタイプを判定
    return 'content'; // プレースホルダー
  }

  private async extractSlideTextContent(_slide: PowerPoint.Slide): Promise<string> {
    // 実際の実装では、スライドからテキストを抽出
    return ''; // プレースホルダー
  }

  private determineSlidePosition(
    positions: number[],
    totalSlides: number
  ): 'any' | 'first' | 'last' | number {
    if (positions.includes(0)) {
      return 'first';
    }
    
    if (positions.includes(totalSlides - 1)) {
      return 'last';
    }
    
    if (positions.length === 1) {
      return positions[0] + 1; // 1ベースのインデックス
    }
    
    return 'any';
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
}