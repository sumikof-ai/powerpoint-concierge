// src/services/powerpoint/types.ts - 型定義
/* global PowerPoint */

export interface SlideInfo {
    id: string;
    title: string;
    content: string;
    index: number;
  }
  
  export interface SlideGenerationOptions {
    includeTransitions?: boolean;
    useTemplate?: string;
    slideLayout?: 'title' | 'content' | 'twoContent' | 'comparison' | 'blank';
    theme?: 'light' | 'dark' | 'colorful';
    fontSize?: 'small' | 'medium' | 'large';
  }
  
  export interface SlideContent {
    title: string;
    content: string[];
    slideType: 'title' | 'content' | 'conclusion';
    speakerNotes?: string;
  }
  
  export interface BulkSlideData {
    slides: SlideContent[];
    options?: SlideGenerationOptions;
  }
  
  export interface FontSizes {
    title: number;
    heading: number;
    subtitle: number;
    body: number;
    accent: number;
  }
  
  export interface PresentationStats {
    slideCount: number;
    estimatedDuration: number;
    wordCount: number;
  }
  
  export type SlideLayoutType = 'title' | 'content' | 'twoContent' | 'comparison' | 'blank';
  export type ThemeType = 'light' | 'dark' | 'colorful';
  export type FontSizeType = 'small' | 'medium' | 'large';
  export type ColorType = 'title' | 'heading' | 'subtitle' | 'body' | 'accent';
  
  export interface PowerPointContext {
    context: PowerPoint.RequestContext;
    slide: PowerPoint.Slide;
  }
  
  export interface ShapeOptions {
    left: number;
    top: number;
    width: number;
    height: number;
  }