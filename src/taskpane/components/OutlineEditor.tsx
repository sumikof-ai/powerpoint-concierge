import * as React from "react";
import { useState, useCallback } from "react";
import {
  Button,
  Card,
  CardHeader,
  CardPreview,
  Text,
  Input,
  Field,
  Textarea,
  Divider,
  tokens,
  makeStyles,
  Badge,
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
} from "@fluentui/react-components";
import {
  Edit24Regular,
  Delete24Regular,
  Add24Regular,
  ArrowUp24Regular,
  ArrowDown24Regular,
  DocumentAdd24Regular,
  Clock24Regular,
} from "@fluentui/react-icons";

// アウトライン関連の型定義
export interface SlideOutline {
  slideNumber: number;
  title: string;
  content: string[];
  slideType: 'title' | 'content' | 'conclusion';
  speakerNotes?: string;
}

export interface PresentationOutline {
  title: string;
  slides: SlideOutline[];
  estimatedDuration: number;
}

interface OutlineEditorProps {
  outline: PresentationOutline | null;
  onOutlineUpdate: (outline: PresentationOutline) => void;
  onGenerateSlides: (outline: PresentationOutline) => void;
  onRegenerateOutline: (instruction: string) => void;
  isLoading?: boolean;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  headerCard: {
    marginBottom: "8px",
  },
  titleSection: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "16px",
  },
  titleInput: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
  },
  metaInfo: {
    display: "flex",
    alignItems: "center",
    gap: "16px",
    marginBottom: "16px",
  },
  slideCard: {
    marginBottom: "8px",
  },
  slideHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  slideControls: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
  contentList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "12px",
  },
  contentItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  contentInput: {
    flex: 1,
  },
  addButton: {
    alignSelf: "flex-start",
    marginTop: "8px",
  },
  actionButtons: {
    display: "flex",
    gap: "12px",
    justifyContent: "flex-end",
    marginTop: "24px",
    paddingTop: "16px",
    borderTop: "1px solid " + tokens.colorNeutralStroke2,
  },
  regenerateSection: {
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "16px",
  },
  slideTypeDisplay: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
});

const OutlineEditor: React.FC<OutlineEditorProps> = ({
  outline,
  onOutlineUpdate,
  onGenerateSlides,
  onRegenerateOutline,
  isLoading = false,
}) => {
  const styles = useStyles();
  const [editingOutline, setEditingOutline] = useState<PresentationOutline | null>(outline);
  const [regenerateInstruction, setRegenerateInstruction] = useState<string>("");

  // アウトラインが更新された時にローカル状態も更新
  React.useEffect(() => {
    setEditingOutline(outline);
  }, [outline]);

  // プレゼンテーションタイトルの更新
  const updateTitle = useCallback((newTitle: string) => {
    if (!editingOutline) return;
    const updated = { ...editingOutline, title: newTitle };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // 予想時間の更新
  const updateDuration = useCallback((newDuration: number) => {
    if (!editingOutline) return;
    const updated = { ...editingOutline, estimatedDuration: newDuration };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // スライドタイトルの更新
  const updateSlideTitle = useCallback((slideIndex: number, newTitle: string) => {
    if (!editingOutline) return;
    const updatedSlides = [...editingOutline.slides];
    updatedSlides[slideIndex] = { ...updatedSlides[slideIndex], title: newTitle };
    const updated = { ...editingOutline, slides: updatedSlides };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // スライドコンテンツの更新
  const updateSlideContent = useCallback((slideIndex: number, contentIndex: number, newContent: string) => {
    if (!editingOutline) return;
    const updatedSlides = [...editingOutline.slides];
    const updatedContent = [...updatedSlides[slideIndex].content];
    updatedContent[contentIndex] = newContent;
    updatedSlides[slideIndex] = { ...updatedSlides[slideIndex], content: updatedContent };
    const updated = { ...editingOutline, slides: updatedSlides };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // コンテンツアイテムの追加
  const addContentItem = useCallback((slideIndex: number) => {
    if (!editingOutline) return;
    const updatedSlides = [...editingOutline.slides];
    updatedSlides[slideIndex] = {
      ...updatedSlides[slideIndex],
      content: [...updatedSlides[slideIndex].content, "新しいポイント"]
    };
    const updated = { ...editingOutline, slides: updatedSlides };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // コンテンツアイテムの削除
  const removeContentItem = useCallback((slideIndex: number, contentIndex: number) => {
    if (!editingOutline) return;
    const updatedSlides = [...editingOutline.slides];
    const updatedContent = updatedSlides[slideIndex].content.filter((_, idx) => idx !== contentIndex);
    updatedSlides[slideIndex] = { ...updatedSlides[slideIndex], content: updatedContent };
    const updated = { ...editingOutline, slides: updatedSlides };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // スライドの削除
  const removeSlide = useCallback((slideIndex: number) => {
    if (!editingOutline) return;
    const updatedSlides = editingOutline.slides.filter((_, idx) => idx !== slideIndex);
    // スライド番号を再調整
    const reindexedSlides = updatedSlides.map((slide, idx) => ({
      ...slide,
      slideNumber: idx + 1
    }));
    const updated = { ...editingOutline, slides: reindexedSlides };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // 新しいスライドの追加
  const addSlide = useCallback(() => {
    if (!editingOutline) return;
    const newSlide: SlideOutline = {
      slideNumber: editingOutline.slides.length + 1,
      title: "新しいスライド",
      content: ["新しいポイント"],
      slideType: 'content'
    };
    const updated = { 
      ...editingOutline, 
      slides: [...editingOutline.slides, newSlide] 
    };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // スライドの順序変更
  const moveSlide = useCallback((slideIndex: number, direction: 'up' | 'down') => {
    if (!editingOutline) return;
    const slides = [...editingOutline.slides];
    const newIndex = direction === 'up' ? slideIndex - 1 : slideIndex + 1;
    
    if (newIndex < 0 || newIndex >= slides.length) return;
    
    // スライドを交換
    [slides[slideIndex], slides[newIndex]] = [slides[newIndex], slides[slideIndex]];
    
    // スライド番号を再調整
    const reindexedSlides = slides.map((slide, idx) => ({
      ...slide,
      slideNumber: idx + 1
    }));
    
    const updated = { ...editingOutline, slides: reindexedSlides };
    setEditingOutline(updated);
    onOutlineUpdate(updated);
  }, [editingOutline, onOutlineUpdate]);

  // スライドタイプの表示文字列
  const getSlideTypeDisplay = (type: string) => {
    switch (type) {
      case 'title': return 'タイトルスライド';
      case 'content': return 'コンテンツスライド';
      case 'conclusion': return 'まとめスライド';
      default: return 'その他';
    }
  };

  // 再生成の実行
  const handleRegenerate = () => {
    if (regenerateInstruction.trim()) {
      onRegenerateOutline(regenerateInstruction.trim());
      setRegenerateInstruction("");
    }
  };

  if (!editingOutline) {
    return (
      <div className={styles.container}>
        <Text>アウトラインが生成されていません。チャットでプレゼンテーションの要件を入力してください。</Text>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      {/* プレゼンテーション全体情報 */}
      <Card className={styles.headerCard}>
        <CardHeader 
          header={<Text weight="semibold">プレゼンテーション概要</Text>}
        />
        <CardPreview>
          <div className={styles.titleSection}>
            <Field label="タイトル" style={{ flex: 1 }}>
              <Input
                value={editingOutline.title}
                onChange={(e) => updateTitle(e.target.value)}
                className={styles.titleInput}
                disabled={isLoading}
              />
            </Field>
          </div>
          
          <div className={styles.metaInfo}>
            <Badge appearance="outline" icon={<DocumentAdd24Regular />}>
              {editingOutline.slides.length} スライド
            </Badge>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Clock24Regular />
              <Input
                type="number"
                value={editingOutline.estimatedDuration.toString()}
                onChange={(e) => updateDuration(parseInt(e.target.value) || 0)}
                disabled={isLoading}
                style={{ width: '80px' }}
              />
              <Text>分</Text>
            </div>
          </div>
        </CardPreview>
      </Card>

      {/* スライド一覧 */}
      <Accordion multiple collapsible>
        {editingOutline.slides.map((slide, slideIndex) => (
          <AccordionItem key={slide.slideNumber} value={`slide-${slideIndex}`}>
            <AccordionHeader>
              <div className={styles.slideHeader}>
                <div style={{ flex: 1 }}>
                  <Text weight="semibold">
                    スライド {slide.slideNumber}: {slide.title}
                  </Text>
                  <div className={styles.slideTypeDisplay}>
                    {getSlideTypeDisplay(slide.slideType)}
                  </div>
                </div>
                
                <div className={styles.slideControls}>
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowUp24Regular />}
                    onClick={() => moveSlide(slideIndex, 'up')}
                    disabled={slideIndex === 0 || isLoading}
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<ArrowDown24Regular />}
                    onClick={() => moveSlide(slideIndex, 'down')}
                    disabled={slideIndex === editingOutline.slides.length - 1 || isLoading}
                  />
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<Delete24Regular />}
                    onClick={() => removeSlide(slideIndex)}
                    disabled={editingOutline.slides.length <= 1 || isLoading}
                  />
                </div>
              </div>
            </AccordionHeader>
            
            <AccordionPanel>
              <div style={{ padding: '12px 0' }}>
                <Field label="スライドタイトル" style={{ marginBottom: '16px' }}>
                  <Input
                    value={slide.title}
                    onChange={(e) => updateSlideTitle(slideIndex, e.target.value)}
                    disabled={isLoading}
                  />
                </Field>

                <Field label="コンテンツ">
                  <div className={styles.contentList}>
                    {slide.content.map((content, contentIndex) => (
                      <div key={contentIndex} className={styles.contentItem}>
                        <Text>•</Text>
                        <Input
                          className={styles.contentInput}
                          value={content}
                          onChange={(e) => updateSlideContent(slideIndex, contentIndex, e.target.value)}
                          disabled={isLoading}
                        />
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<Delete24Regular />}
                          onClick={() => removeContentItem(slideIndex, contentIndex)}
                          disabled={slide.content.length <= 1 || isLoading}
                        />
                      </div>
                    ))}
                    
                    <Button
                      className={styles.addButton}
                      size="small"
                      appearance="subtle"
                      icon={<Add24Regular />}
                      onClick={() => addContentItem(slideIndex)}
                      disabled={isLoading}
                    >
                      ポイントを追加
                    </Button>
                  </div>
                </Field>
              </div>
            </AccordionPanel>
          </AccordionItem>
        ))}
      </Accordion>

      {/* 新しいスライド追加 */}
      <Button
        appearance="outline"
        icon={<Add24Regular />}
        onClick={addSlide}
        disabled={isLoading}
      >
        スライドを追加
      </Button>

      <Divider />

      {/* 再生成セクション */}
      <div className={styles.regenerateSection}>
        <Text weight="semibold" style={{ marginBottom: '12px' }}>
          AIによる再生成
        </Text>
        <Field label="追加指示・修正要望">
          <Textarea
            placeholder="例: もっと詳しい内容にして、図表を多用したい、技術的な内容を減らして..."
            value={regenerateInstruction}
            onChange={(e) => setRegenerateInstruction(e.target.value)}
            rows={3}
            disabled={isLoading}
          />
        </Field>
        <Button
          appearance="secondary"
          icon={<Edit24Regular />}
          onClick={handleRegenerate}
          disabled={!regenerateInstruction.trim() || isLoading}
          style={{ marginTop: '8px' }}
        >
          アウトラインを再生成
        </Button>
      </div>

      {/* アクションボタン */}
      <div className={styles.actionButtons}>
        <Button
          appearance="primary"
          icon={<DocumentAdd24Regular />}
          onClick={() => onGenerateSlides(editingOutline)}
          disabled={isLoading}
          size="large"
        >
          このアウトラインでスライドを生成
        </Button>
      </div>
    </div>
  );
};

export default OutlineEditor;