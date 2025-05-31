// src/services/powerpoint.service.ts - PowerPoint.js API互換版
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
  
  export class PowerPointService {
    private defaultOptions: SlideGenerationOptions = {
      includeTransitions: false,
      slideLayout: 'content',
      theme: 'light',
      fontSize: 'medium',
    };
  
    /**
     * 現在のプレゼンテーションの全スライド情報を取得
     */
    public async getAllSlides(): Promise<SlideInfo[]> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
  
            const slideInfos: SlideInfo[] = [];
            
            for (let i = 0; i < slides.items.length; i++) {
              const slide = slides.items[i];
              slide.load("shapes");
              await context.sync();
  
              let title = `スライド ${i + 1}`;
              let content = '';
  
              // シェイプからテキストを抽出
              for (let j = 0; j < slide.shapes.items.length; j++) {
                const shape = slide.shapes.items[j];
                if (shape.type === PowerPoint.ShapeType.textBox || 
                    shape.type === PowerPoint.ShapeType.placeholder) {
                  shape.textFrame.load("textRange");
                  await context.sync();
                  
                  const text = shape.textFrame.textRange.text;
                  if (j === 0 && text) {
                    title = text.substring(0, 50);
                  }
                  content += text + '\n';
                }
              }
  
              slideInfos.push({
                id: slide.id,
                title: title,
                content: content.trim(),
                index: i
              });
            }
  
            resolve(slideInfos);
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * 複数のスライドを一括生成
     */
    public async generateBulkSlides(
      bulkData: BulkSlideData, 
      onProgress?: (current: number, total: number, slideName: string) => void
    ): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const { slides, options = {} } = bulkData;
            const mergedOptions = { ...this.defaultOptions, ...options };
            
            // 各スライドを順番に生成
            for (let i = 0; i < slides.length; i++) {
              const slideData = slides[i];
              
              // 進捗コールバック
              if (onProgress) {
                onProgress(i + 1, slides.length, slideData.title);
              }
  
              // スライドレイアウトの決定
              const layout = this.determineSlideLayout(slideData.slideType, mergedOptions.slideLayout);
              
              // スライドを作成
              await this.createSlideWithLayout(context, slideData, layout, mergedOptions);
            }
  
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * スライドタイプに基づいてレイアウトを決定
     */
    private determineSlideLayout(
      slideType: 'title' | 'content' | 'conclusion', 
      defaultLayout?: string
    ): 'title' | 'content' | 'twoContent' | 'comparison' | 'blank' {
      switch (slideType) {
        case 'title':
          return 'title';
        case 'conclusion':
          return 'content';
        case 'content':
        default:
          return defaultLayout as any || 'content';
      }
    }
  
    /**
     * 指定されたレイアウトでスライドを作成
     */
    private async createSlideWithLayout(
      context: PowerPoint.RequestContext,
      slideData: SlideContent,
      layout: 'title' | 'content' | 'twoContent' | 'comparison' | 'blank',
      options: SlideGenerationOptions
    ): Promise<void> {
      // 新しいスライドを追加
      context.presentation.slides.add();
      await context.sync();
      
      // 最後に追加されたスライドを取得
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      const slide = slides.items[slides.items.length - 1];
      
      switch (layout) {
        case 'title':
          await this.createTitleSlide(context, slide, slideData, options);
          break;
        case 'content':
          await this.createContentSlide(context, slide, slideData, options);
          break;
        case 'twoContent':
          await this.createTwoContentSlide(context, slide, slideData, options);
          break;
        case 'comparison':
          await this.createComparisonSlide(context, slide, slideData, options);
          break;
        case 'blank':
          await this.createBlankSlide(context, slide, slideData, options);
          break;
        default:
          await this.createContentSlide(context, slide, slideData, options);
      }
  
      // スピーカーノートは PowerPoint.js で直接サポートされていないため、
      // コメントとして残すか、スライド内にテキストとして追加する
      if (slideData.speakerNotes) {
        console.log(`スピーカーノート (${slideData.title}): ${slideData.speakerNotes}`);
        // 必要に応じて、スライドの下部にノート用テキストボックスを追加することも可能
      }
    }
  
    /**
     * タイトルスライドを作成
     */
    private async createTitleSlide(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      options: SlideGenerationOptions
    ): Promise<void> {
      const fontSize = this.getFontSize(options.fontSize);
      
      // メインタイトル（中央寄せのために位置を調整）
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 75,   // 中央寄せ効果
        top: 150,
        width: 600,
        height: 150
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.title;
      titleBox.textFrame.textRange.font.bold = true;
      this.applyThemeColors(titleBox, options.theme, 'title');
  
      // サブタイトル（コンテンツがある場合）
      if (slideData.content && slideData.content.length > 0) {
        const subtitleText = slideData.content.join(' | ');
        const subtitleBox = slide.shapes.addTextBox(subtitleText, {
          left: 100,  // 中央寄せ効果
          top: 320,
          width: 550,
          height: 100
        });
        
        await context.sync();
        
        subtitleBox.textFrame.textRange.font.size = fontSize.subtitle;
        this.applyThemeColors(subtitleBox, options.theme, 'subtitle');
      }
    }
  
    /**
     * 標準コンテンツスライドを作成
     */
    private async createContentSlide(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      options: SlideGenerationOptions
    ): Promise<void> {
      const fontSize = this.getFontSize(options.fontSize);
      
      // タイトル
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 50,
        top: 40,
        width: 650,
        height: 80
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
      this.applyThemeColors(titleBox, options.theme, 'heading');
  
      // コンテンツ（箇条書き）
      if (slideData.content && slideData.content.length > 0) {
        // 箇条書きを改行で区切って表示
        const contentText = slideData.content.map(item => `• ${item}`).join('\n\n');
        const contentBox = slide.shapes.addTextBox(contentText, {
          left: 80,
          top: 140,
          width: 580,
          height: 350
        });
        
        await context.sync();
        
        contentBox.textFrame.textRange.font.size = fontSize.body;
        this.applyThemeColors(contentBox, options.theme, 'body');
      }
    }
  
    /**
     * 2カラムコンテンツスライドを作成
     */
    private async createTwoContentSlide(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      options: SlideGenerationOptions
    ): Promise<void> {
      const fontSize = this.getFontSize(options.fontSize);
      
      // タイトル
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 50,
        top: 40,
        width: 650,
        height: 80
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
      this.applyThemeColors(titleBox, options.theme, 'heading');
  
      // コンテンツを2つに分割
      if (slideData.content && slideData.content.length > 0) {
        const midPoint = Math.ceil(slideData.content.length / 2);
        const leftContent = slideData.content.slice(0, midPoint);
        const rightContent = slideData.content.slice(midPoint);
  
        // 左カラム
        if (leftContent.length > 0) {
          const leftText = leftContent.map(item => `• ${item}`).join('\n\n');
          const leftBox = slide.shapes.addTextBox(leftText, {
            left: 50,
            top: 140,
            width: 300,
            height: 350
          });
          
          await context.sync();
          
          leftBox.textFrame.textRange.font.size = fontSize.body;
          this.applyThemeColors(leftBox, options.theme, 'body');
        }
  
        // 右カラム
        if (rightContent.length > 0) {
          const rightText = rightContent.map(item => `• ${item}`).join('\n\n');
          const rightBox = slide.shapes.addTextBox(rightText, {
            left: 380,
            top: 140,
            width: 300,
            height: 350
          });
          
          await context.sync();
          
          rightBox.textFrame.textRange.font.size = fontSize.body;
          this.applyThemeColors(rightBox, options.theme, 'body');
        }
      }
    }
  
    /**
     * 比較スライドを作成
     */
    private async createComparisonSlide(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      options: SlideGenerationOptions
    ): Promise<void> {
      const fontSize = this.getFontSize(options.fontSize);
      
      // タイトル
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 50,
        top: 40,
        width: 650,
        height: 80
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
      this.applyThemeColors(titleBox, options.theme, 'heading');
  
      // 比較表形式でコンテンツを配置
      if (slideData.content && slideData.content.length > 0) {
        // 左側ヘッダー
        const leftHeaderBox = slide.shapes.addTextBox("項目", {
          left: 50,
          top: 140,
          width: 300,
          height: 40
        });
        
        await context.sync();
        
        leftHeaderBox.textFrame.textRange.font.bold = true;
        leftHeaderBox.textFrame.textRange.font.size = fontSize.accent;
        this.applyThemeColors(leftHeaderBox, options.theme, 'accent');
  
        // 右側ヘッダー
        const rightHeaderBox = slide.shapes.addTextBox("詳細", {
          left: 380,
          top: 140,
          width: 300,
          height: 40
        });
        
        await context.sync();
        
        rightHeaderBox.textFrame.textRange.font.bold = true;
        rightHeaderBox.textFrame.textRange.font.size = fontSize.accent;
        this.applyThemeColors(rightHeaderBox, options.theme, 'accent');
  
        // コンテンツを交互に配置
        for (let index = 0; index < slideData.content.length && index < 8; index++) {
          const item = slideData.content[index];
          const yPos = 200 + (index * 35);
          const isLeft = index % 2 === 0;
          
          const contentBox = slide.shapes.addTextBox(`• ${item}`, {
            left: isLeft ? 50 : 380,
            top: yPos,
            width: 300,
            height: 30
          });
          
          await context.sync();
          
          contentBox.textFrame.textRange.font.size = fontSize.body;
          this.applyThemeColors(contentBox, options.theme, 'body');
        }
      }
    }
  
    /**
     * 空白スライドを作成（タイトルのみ）
     */
    private async createBlankSlide(
      context: PowerPoint.RequestContext,
      slide: PowerPoint.Slide,
      slideData: SlideContent,
      options: SlideGenerationOptions
    ): Promise<void> {
      const fontSize = this.getFontSize(options.fontSize);
      
      // タイトルのみ
      const titleBox = slide.shapes.addTextBox(slideData.title, {
        left: 50,
        top: 40,
        width: 650,
        height: 80
      });
      
      await context.sync();
      
      titleBox.textFrame.textRange.font.size = fontSize.heading;
      titleBox.textFrame.textRange.font.bold = true;
      this.applyThemeColors(titleBox, options.theme, 'heading');
    }
  
    /**
     * フォントサイズを取得
     */
    private getFontSize(size?: 'small' | 'medium' | 'large') {
      switch (size) {
        case 'small':
          return { title: 32, heading: 20, subtitle: 16, body: 12, accent: 14 };
        case 'large':
          return { title: 44, heading: 28, subtitle: 22, body: 16, accent: 18 };
        case 'medium':
        default:
          return { title: 38, heading: 24, subtitle: 18, body: 14, accent: 16 };
      }
    }
  
    /**
     * テーマに基づいて色を適用
     * PowerPoint.js で確実に動作する基本的な色設定のみ使用
     */
    private applyThemeColors(
      shape: PowerPoint.Shape, 
      theme?: 'light' | 'dark' | 'colorful', 
      type?: 'title' | 'heading' | 'subtitle' | 'body' | 'accent'
    ): void {
      try {
        switch (theme) {
          case 'dark':
            shape.fill.setSolidColor('#2D2D30');
            shape.textFrame.textRange.font.color = '#FFFFFF';
            if (type === 'accent') {
              shape.fill.setSolidColor('#007ACC');
            }
            break;
          case 'colorful':
            switch (type) {
              case 'title':
                shape.fill.setSolidColor('#FF6B6B');
                shape.textFrame.textRange.font.color = '#FFFFFF';
                break;
              case 'heading':
                shape.fill.setSolidColor('#4ECDC4');
                shape.textFrame.textRange.font.color = '#FFFFFF';
                break;
              case 'accent':
                shape.fill.setSolidColor('#45B7D1');
                shape.textFrame.textRange.font.color = '#FFFFFF';
                break;
              default:
                shape.fill.setSolidColor('#FFFFFF');
                shape.textFrame.textRange.font.color = '#2C3E50';
            }
            break;
          case 'light':
          default:
            shape.fill.setSolidColor('#FFFFFF');
            shape.textFrame.textRange.font.color = '#333333';
            if (type === 'accent') {
              shape.fill.setSolidColor('#F8F9FA');
              if (shape.lineFormat) {
                shape.lineFormat.color = '#DEE2E6';
                shape.lineFormat.weight = 1;
              }
            }
        }
      } catch (error) {
        console.warn('テーマ色の適用に失敗しました:', error);
        // フォールバック: 基本的な白背景、黒文字
        try {
          shape.fill.setSolidColor('#FFFFFF');
          shape.textFrame.textRange.font.color = '#000000';
        } catch (fallbackError) {
          console.warn('フォールバック色の適用も失敗しました:', fallbackError);
        }
      }
    }
  
    /**
     * 新しいスライドを追加（従来版、互換性のため保持）
     */
    public async addSlide(title: string, content: string): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            // 新しいスライドを追加
            context.presentation.slides.add();
            await context.sync();
            
            // 最後に追加されたスライドを取得
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
            const slide = slides.items[slides.items.length - 1];
            
            // タイトルテキストボックスを追加
            const titleTextBox = slide.shapes.addTextBox(title, {
              left: 50,
              top: 50,
              width: 600,
              height: 100
            });
            
            await context.sync();
            
            // タイトルのフォーマット設定
            titleTextBox.textFrame.textRange.font.size = 24;
            titleTextBox.textFrame.textRange.font.bold = true;
            titleTextBox.fill.setSolidColor("white");
            if (titleTextBox.lineFormat) {
              titleTextBox.lineFormat.color = "black";
              titleTextBox.lineFormat.weight = 2;
            }
  
            // コンテンツテキストボックスを追加
            if (content) {
              const contentTextBox = slide.shapes.addTextBox(content, {
                left: 50,
                top: 180,
                width: 600,
                height: 400
              });
              
              await context.sync();
              
              // コンテンツのフォーマット設定
              contentTextBox.textFrame.textRange.font.size = 16;
              contentTextBox.fill.setSolidColor("white");
              if (contentTextBox.lineFormat) {
                contentTextBox.lineFormat.color = "gray";
                contentTextBox.lineFormat.weight = 1;
              }
            }
  
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * 指定したスライドのコンテンツを更新
     */
    public async updateSlide(slideIndex: number, title: string, content: string): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
  
            if (slideIndex >= slides.items.length) {
              throw new Error(`スライド ${slideIndex + 1} が見つかりません`);
            }
  
            const slide = slides.items[slideIndex];
            slide.shapes.load("items");
            await context.sync();
  
            // 既存のテキストボックスをクリア
            for (let i = slide.shapes.items.length - 1; i >= 0; i--) {
              const shape = slide.shapes.items[i];
              if (shape.type === PowerPoint.ShapeType.textBox) {
                shape.delete();
              }
            }
  
            await context.sync();
  
            // 新しいコンテンツで再作成
            const slideData: SlideContent = {
              title,
              content: content.split('\n• ').filter(item => item.trim() !== ''),
              slideType: 'content'
            };
  
            await this.createContentSlide(context, slide, slideData, this.defaultOptions);
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * 指定したスライドを削除
     */
    public async deleteSlide(slideIndex: number): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
  
            if (slideIndex >= slides.items.length) {
              throw new Error(`スライド ${slideIndex + 1} が見つかりません`);
            }
  
            slides.items[slideIndex].delete();
            await context.sync();
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * スライド間にトランジションを追加
     * 注意: PowerPoint.js では現在トランジション機能のサポートが限定的
     */
    public async addTransitions(transitionType: 'fade' | 'slide' | 'none' = 'fade'): Promise<void> {
      return new Promise((resolve) => {
        console.log(`トランジション設定をリクエストしました: ${transitionType}`);
        console.log('注意: PowerPoint.js では現在トランジション機能のサポートが限定的です');
        resolve();
      });
    }
  
    /**
     * プレゼンテーション全体の統計を取得
     */
    public async getPresentationStats(): Promise<{
      slideCount: number;
      estimatedDuration: number;
      wordCount: number;
    }> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
  
            let totalWords = 0;
            
            for (let i = 0; i < slides.items.length; i++) {
              const slide = slides.items[i];
              slide.shapes.load("items");
              await context.sync();
  
              for (let j = 0; j < slide.shapes.items.length; j++) {
                const shape = slide.shapes.items[j];
                if (shape.type === PowerPoint.ShapeType.textBox || 
                    shape.type === PowerPoint.ShapeType.placeholder) {
                  shape.textFrame.load("textRange");
                  await context.sync();
                  
                  const text = shape.textFrame.textRange.text;
                  totalWords += text.split(/\s+/).filter(word => word.length > 0).length;
                }
              }
            }
  
            // 1スライドあたり平均2分として概算
            const estimatedDuration = slides.items.length * 2;
  
            resolve({
              slideCount: slides.items.length,
              estimatedDuration,
              wordCount: totalWords
            });
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  
    /**
     * テキストボックスを追加（テスト用）
     */
    public async addTextBox(text: string): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            let slide;
            try {
              slide = context.presentation.getSelectedSlides().getItemAt(0);
            } catch {
              slide = context.presentation.slides.getItemAt(0);
            }
  
            const textBox = slide.shapes.addTextBox(text, {
              left: 100,
              top: 200,
              width: 500,
              height: 200
            });
            
            await context.sync();
            
            textBox.textFrame.textRange.font.size = 14;
            textBox.fill.setSolidColor("white");
            if (textBox.lineFormat) {
              textBox.lineFormat.color = "blue";
              textBox.lineFormat.weight = 1;
              textBox.lineFormat.dashStyle = "solid";
            }
  
            await context.sync();
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  }