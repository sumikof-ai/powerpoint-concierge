// src/taskpane/components/outline/OutlineEditor.tsx - リファクタリング版メインコンポーネント
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
  Badge,
  Accordion,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import {
  Edit24Regular,
  Add24Regular,
  DocumentAdd24Regular,
  Clock24Regular,
} from "@fluentui/react-icons";
import { PresentationOutline, SlideOutline } from "../types";
import SlideEditor from "./SlideEditor";

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
  durationInput: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  durationField: {
    width: "80px",
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
  regenerateField: {
    marginBottom: "12px",
  },
  addSlideButton: {
    marginBottom: "16px",
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

  const updateOutline = useCallback(
    (updatedOutline: PresentationOutline) => {
      setEditingOutline(updatedOutline);
      onOutlineUpdate(updatedOutline);
    },
    [onOutlineUpdate]
  );

  // プレゼンテーションタイトルの更新
  const updateTitle = useCallback(
    (newTitle: string) => {
      if (!editingOutline) return;
      const updated = { ...editingOutline, title: newTitle };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // 予想時間の更新
  const updateDuration = useCallback(
    (newDuration: number) => {
      if (!editingOutline) return;
      const updated = { ...editingOutline, estimatedDuration: newDuration };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // スライドタイトルの更新
  const updateSlideTitle = useCallback(
    (slideIndex: number, newTitle: string) => {
      if (!editingOutline) return;
      const updatedSlides = [...editingOutline.slides];
      updatedSlides[slideIndex] = { ...updatedSlides[slideIndex], title: newTitle };
      const updated = { ...editingOutline, slides: updatedSlides };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // スライドコンテンツの更新
  const updateSlideContent = useCallback(
    (slideIndex: number, contentIndex: number, newContent: string) => {
      if (!editingOutline) return;
      const updatedSlides = [...editingOutline.slides];
      const updatedContent = [...updatedSlides[slideIndex].content];
      updatedContent[contentIndex] = newContent;
      updatedSlides[slideIndex] = { ...updatedSlides[slideIndex], content: updatedContent };
      const updated = { ...editingOutline, slides: updatedSlides };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // コンテンツアイテムの追加
  const addContentItem = useCallback(
    (slideIndex: number) => {
      if (!editingOutline) return;
      const updatedSlides = [...editingOutline.slides];
      updatedSlides[slideIndex] = {
        ...updatedSlides[slideIndex],
        content: [...updatedSlides[slideIndex].content, "新しいポイント"],
      };
      const updated = { ...editingOutline, slides: updatedSlides };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // コンテンツアイテムの削除
  const removeContentItem = useCallback(
    (slideIndex: number, contentIndex: number) => {
      if (!editingOutline) return;
      const updatedSlides = [...editingOutline.slides];
      const updatedContent = updatedSlides[slideIndex].content.filter(
        (_, idx) => idx !== contentIndex
      );
      updatedSlides[slideIndex] = { ...updatedSlides[slideIndex], content: updatedContent };
      const updated = { ...editingOutline, slides: updatedSlides };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // スライドの削除
  const removeSlide = useCallback(
    (slideIndex: number) => {
      if (!editingOutline) return;
      const updatedSlides = editingOutline.slides.filter((_, idx) => idx !== slideIndex);
      // スライド番号を再調整
      const reindexedSlides = updatedSlides.map((slide, idx) => ({
        ...slide,
        slideNumber: idx + 1,
      }));
      const updated = { ...editingOutline, slides: reindexedSlides };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // 新しいスライドの追加
  const addSlide = useCallback(() => {
    if (!editingOutline) return;
    const newSlide: SlideOutline = {
      slideNumber: editingOutline.slides.length + 1,
      title: "新しいスライド",
      content: ["新しいポイント"],
      slideType: "content",
    };
    const updated = {
      ...editingOutline,
      slides: [...editingOutline.slides, newSlide],
    };
    updateOutline(updated);
  }, [editingOutline, updateOutline]);

  // スライドの順序変更
  const moveSlide = useCallback(
    (slideIndex: number, direction: "up" | "down") => {
      if (!editingOutline) return;
      const slides = [...editingOutline.slides];
      const newIndex = direction === "up" ? slideIndex - 1 : slideIndex + 1;

      if (newIndex < 0 || newIndex >= slides.length) return;

      // スライドを交換
      [slides[slideIndex], slides[newIndex]] = [slides[newIndex], slides[slideIndex]];

      // スライド番号を再調整
      const reindexedSlides = slides.map((slide, idx) => ({
        ...slide,
        slideNumber: idx + 1,
      }));

      const updated = { ...editingOutline, slides: reindexedSlides };
      updateOutline(updated);
    },
    [editingOutline, updateOutline]
  );

  // 再生成の実行
  const handleRegenerate = useCallback(() => {
    if (regenerateInstruction.trim()) {
      onRegenerateOutline(regenerateInstruction.trim());
      setRegenerateInstruction("");
    }
  }, [regenerateInstruction, onRegenerateOutline]);

  // アウトラインが生成されていない場合
  if (!editingOutline) {
    return (
      <div className={styles.container}>
        <Text>
          アウトラインが生成されていません。チャットでプレゼンテーションの要件を入力してください。
        </Text>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      {/* プレゼンテーション全体情報 */}
      <Card className={styles.headerCard}>
        <CardHeader header={<Text weight="semibold">プレゼンテーション概要</Text>} />
        <CardPreview>
          <div className={styles.titleSection}>
            <Field label="タイトル" style={{ flex: 1 }}>
              <Input
                value={editingOutline.title}
                onChange={(e) => updateTitle(e.target.value)}
                className={styles.titleInput}
                disabled={isLoading}
                placeholder="プレゼンテーションタイトル"
              />
            </Field>
          </div>

          <div className={styles.metaInfo}>
            <Badge appearance="outline" icon={<DocumentAdd24Regular />}>
              {editingOutline.slides.length} スライド
            </Badge>
            <div className={styles.durationInput}>
              <Clock24Regular />
              <Input
                className={styles.durationField}
                type="number"
                value={editingOutline.estimatedDuration.toString()}
                onChange={(e) => updateDuration(parseInt(e.target.value) || 0)}
                disabled={isLoading}
                min="1"
                max="120"
              />
              <Text>分</Text>
            </div>
          </div>
        </CardPreview>
      </Card>

      {/* スライド一覧 */}
      <Accordion multiple collapsible>
        {editingOutline.slides.map((slide, slideIndex) => (
          <SlideEditor
            key={slide.slideNumber}
            slide={slide}
            slideIndex={slideIndex}
            totalSlides={editingOutline.slides.length}
            onUpdateSlideTitle={updateSlideTitle}
            onUpdateSlideContent={updateSlideContent}
            onAddContentItem={addContentItem}
            onRemoveContentItem={removeContentItem}
            onRemoveSlide={removeSlide}
            onMoveSlide={moveSlide}
            isLoading={isLoading}
          />
        ))}
      </Accordion>

      {/* 新しいスライド追加 */}
      <Button
        className={styles.addSlideButton}
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
        <Text weight="semibold" style={{ marginBottom: "12px" }}>
          AIによる再生成
        </Text>
        <Field label="追加指示・修正要望" className={styles.regenerateField}>
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
