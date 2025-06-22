// Template Integration System Types
// テンプレート統合システムの型定義
/* global File */

export interface TemplateInfo {
  id: string;
  name: string;
  description?: string;
  category: TemplateCategory;
  metadata: TemplateMetadata;
  designPatterns: DesignPattern[];
  structure: TemplateStructure;
  compatibility: TemplateCompatibility;
}

export interface TemplateMetadata {
  industry?: string[];
  presentationStyle: "formal" | "casual" | "technical" | "creative" | "minimal";
  targetAudience: "executive" | "technical" | "general" | "academic" | "sales";
  slideCount: number;
  colorSchemeType: "corporate" | "vibrant" | "minimal" | "academic" | "creative";
  layoutComplexity: "simple" | "moderate" | "complex";
  contentDensity: "low" | "medium" | "high";
  purpose: "pitch" | "report" | "training" | "marketing" | "analysis";
  tags: string[];
  registeredAt: Date;
  lastUsed?: Date;
  usageCount: number;
}

export interface DesignPattern {
  type: "layout" | "color" | "typography" | "spacing" | "imagery";
  pattern: string;
  frequency: number;
  importance: "critical" | "important" | "optional";
  description: string;
  rules: PatternRule[];
}

export interface PatternRule {
  condition: string;
  action: string;
  priority: number;
}

export interface TemplateStructure {
  expectedSlideTypes: SlideTypePattern[];
  contentFlow: ContentFlowPattern[];
  navigationPattern: NavigationPattern;
  visualHierarchy: VisualHierarchyPattern[];
}

export interface SlideTypePattern {
  position: number | "any" | "first" | "last";
  type: "title" | "agenda" | "content" | "comparison" | "conclusion" | "transition";
  frequency: number;
  required: boolean;
  variations: string[];
}

export interface ContentFlowPattern {
  fromSlideType: string;
  toSlideType: string;
  probability: number;
  transitionStyle?: string;
}

export interface NavigationPattern {
  hasAgenda: boolean;
  hasTableOfContents: boolean;
  sectionDividers: boolean;
  backNavigation: boolean;
}

export interface VisualHierarchyPattern {
  level: number;
  element: "title" | "subtitle" | "content" | "accent";
  fontSize: number;
  fontWeight: string;
  color: string;
  positioning: "top" | "center" | "bottom" | "left" | "right";
}

export interface TemplateCompatibility {
  powerPointVersion: string[];
  supportedFeatures: string[];
  limitations: string[];
  apiRequirements: string[];
}

export type TemplateCategory =
  | "business"
  | "academic"
  | "creative"
  | "technical"
  | "marketing"
  | "corporate"
  | "minimal"
  | "custom";

export interface TemplateRecommendation {
  template: TemplateInfo;
  score: number;
  reasoning: string[];
  adaptations: TemplateAdaptation[];
}

export interface TemplateAdaptation {
  type: "content" | "layout" | "style" | "structure";
  description: string;
  confidence: number;
  changes: AdaptationChange[];
}

export interface AdaptationChange {
  target: string;
  from: any;
  to: any;
  reason: string;
}

export interface AdaptedOutline {
  originalOutline: any;
  selectedTemplate: TemplateInfo;
  adaptations: TemplateAdaptation[];
  adaptedSlides: AdaptedSlide[];
  confidence: number;
}

export interface AdaptedSlide {
  slideNumber: number;
  originalSlide: any;
  adaptedContent: {
    title: string;
    content: string[];
    slideType: string;
    layoutSuggestion: string;
    styleOverrides: Record<string, any>;
  };
  templatePatterns: DesignPattern[];
  adaptationNotes: string[];
}

export interface TemplateAnalysisResult {
  detectedPatterns: DesignPattern[];
  extractedStructure: TemplateStructure;
  suggestedMetadata: Partial<TemplateMetadata>;
  confidence: number;
  analysisNotes: string[];
}

export interface TemplateRegistrationRequest {
  file?: File;
  url?: string;
  metadata: Partial<TemplateMetadata>;
  autoAnalyze: boolean;
}

export interface TemplateLibrary {
  templates: TemplateInfo[];
  categories: Record<TemplateCategory, TemplateInfo[]>;
  searchIndex: TemplateSearchIndex;
  statistics: TemplateLibraryStats;
}

export interface TemplateSearchIndex {
  byIndustry: Record<string, string[]>;
  byStyle: Record<string, string[]>;
  byPurpose: Record<string, string[]>;
  byTags: Record<string, string[]>;
}

export interface TemplateLibraryStats {
  totalTemplates: number;
  byCategory: Record<TemplateCategory, number>;
  mostUsed: string[];
  recentlyAdded: string[];
  averageScore: number;
}

export interface TemplateSelectionCriteria {
  userInput: string;
  presentationContext: {
    audience?: string;
    purpose?: string;
    duration?: number;
    style?: string;
  };
  preferences: {
    categories?: TemplateCategory[];
    excludeCategories?: TemplateCategory[];
    minimumScore?: number;
    maxResults?: number;
  };
}

export interface TemplateUsageStats {
  templateId: string;
  usageCount: number;
  lastUsed: Date;
  averageRating?: number;
  userFeedback?: string[];
  successRate?: number;
}
