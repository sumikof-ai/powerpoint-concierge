// src/taskpane/services/powerpoint.ts
/* global PowerPoint */

export interface SlideInfo {
    id: string;
    title: string;
    content: string;
    index: number;
  }
  
  export class PowerPointService {
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
                    title = text.substring(0, 50); // 最初のテキストをタイトルとして使用
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
     * 新しいスライドを追加
     */
    public async addSlide(title: string, content: string): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            // 新しいスライドを追加
            // const slide = context.presentation.slides.add();
            const pageCount = context.presentation.slides.getCount();
            const slide = context.presentation.slides.getItemAt(pageCount.value);
                        
            // タイトルテキストボックスを追加
            const titleTextBox = slide.shapes.addTextBox(title, {
              left: 50,
              top: 50,
              width: 600,
              height: 100
            });
            
            // タイトルのフォーマット設定
            titleTextBox.textFrame.textRange.font.size = 24;
            titleTextBox.textFrame.textRange.font.bold = true;
            titleTextBox.fill.setSolidColor("white");
            titleTextBox.lineFormat.color = "black";
            titleTextBox.lineFormat.weight = 2;
  
            // コンテンツテキストボックスを追加
            if (content) {
              const contentTextBox = slide.shapes.addTextBox(content, {
                left: 50,
                top: 180,
                width: 600,
                height: 400
              });
              
              // コンテンツのフォーマット設定
              contentTextBox.textFrame.textRange.font.size = 16;
              contentTextBox.fill.setSolidColor("white");
              contentTextBox.lineFormat.color = "gray";
              contentTextBox.lineFormat.weight = 1;
            }
  
            await context.sync();
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
  
            // 新しいタイトルテキストボックスを追加
            const titleTextBox = slide.shapes.addTextBox(title, {
              left: 50,
              top: 50,
              width: 600,
              height: 100
            });
            
            titleTextBox.textFrame.textRange.font.size = 24;
            titleTextBox.textFrame.textRange.font.bold = true;
            titleTextBox.fill.setSolidColor("white");
            titleTextBox.lineFormat.color = "black";
            titleTextBox.lineFormat.weight = 2;
  
            // 新しいコンテンツテキストボックスを追加
            if (content) {
              const contentTextBox = slide.shapes.addTextBox(content, {
                left: 50,
                top: 180,
                width: 600,
                height: 400
              });
              
              contentTextBox.textFrame.textRange.font.size = 16;
              contentTextBox.fill.setSolidColor("white");
              contentTextBox.lineFormat.color = "gray";
              contentTextBox.lineFormat.weight = 1;
            }
  
            await context.sync();
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
     * テキストボックスを追加（テスト用）
     */
    public async addTextBox(text: string): Promise<void> {
      return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
          try {
            // 現在選択されているスライドを取得、なければ最初のスライドを使用
            let slide;
            try {
              slide = context.presentation.getSelectedSlides().getItemAt(0);
            } catch {
              // 選択されたスライドがない場合、最初のスライドを使用
              slide = context.presentation.slides.getItemAt(0);
            }
  
            // テキストボックスを追加
            const textBox = slide.shapes.addTextBox(text, {
              left: 100,
              top: 200,
              width: 500,
              height: 200
            });
            
            // フォーマット設定
            textBox.textFrame.textRange.font.size = 14;
            textBox.fill.setSolidColor("white");
            textBox.lineFormat.color = "blue";
            textBox.lineFormat.weight = 1;
            textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
  
            await context.sync();
            resolve();
          } catch (error) {
            reject(error);
          }
        });
      });
    }
  }