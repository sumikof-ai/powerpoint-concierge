// src/taskpane/components/outline/SlideEditor.tsx - 個別スライド編集コンポーネント
import * as React from "react";
import { useCallback } from "react";
import {
  Button,
  Text,
  Input,
  Field,
  tokens,
  makeStyles,
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
} from "@fluentui/react-icons";
import { SlideOutline } from '../types';

interface SlideEditorProps {
  slide: SlideOutline;
  slideIndex: number;
  totalSlides: number;
  onUpdateSlideTitle: (slideIndex: number, newTitle: string) => void;
  onUpdateSlideContent: (slideIndex: number, contentIndex: number, newContent: string) => void;
  onAddContentItem: (slideIndex: number) => void;
  onRemoveContentItem: (slideIndex: number, contentIndex: number) => void;
  onRemoveSlide: (slideIndex: number) => void;
  onMoveSlide: (slideIndex: number, direction: 'up' | 'down') => void;
  isLoading: boolean;
}

const useStyles = makeStyles({
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
  slideTypeDisplay: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  accordionContent: {
    padding: '12px 0',
  },
  fieldMargin: {
    marginBottom: '16px',
  },
});

const SlideEditor: React.FC<SlideEditorProps> = ({
  slide,
  slideIndex,
  totalSlides,
  onUpdateSlideTitle,
  onUpdateSlideContent,
  onAddContentItem,
  onRemoveContentItem,
  onRemoveSlide,
  onMoveSlide,
  isLoading,
}) => {
  const styles = useStyles();

  const getSlideTypeDisplay = useCallback((type: string): string => {
    switch (type) {
      case 'title': return 'タイトルスライド';
      case 'content': return 'コンテンツスライド';
      case 'conclusion': return 'まとめスライド';
      default: return 'その他';
    }
  }, []);

  const handleSlideMove = useCallback((direction: 'up' | 'down') => {
    onMoveSlide(slideIndex, direction);
  }, [slideIndex, onMoveSlide]);

  const handleSlideRemove = useCallback(() => {
    if (totalSlides > 1) {
      onRemoveSlide(slideIndex);
    }
  }, [slideIndex, totalSlides, onRemoveSlide]);

  const handleContentAdd = useCallback(() => {
    onAddContentItem(slideIndex);
  }, [slideIndex, onAddContentItem]);

  const handleContentRemove = useCallback((contentIndex: number) => {
    if (slide.content.length > 1) {
      onRemoveContentItem(slideIndex, contentIndex);
    }
  }, [slideIndex, slide.content.length, onRemoveContentItem]);

  const renderSlideHeader = () => (
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
          onClick={() => handleSlideMove('up')}
          disabled={slideIndex === 0 || isLoading}
          title="スライドを上に移動"
        />
        <Button
          size="small"
          appearance="subtle"
          icon={<ArrowDown24Regular />}
          onClick={() => handleSlideMove('down')}
          disabled={slideIndex === totalSlides - 1 || isLoading}
          title="スライドを下に移動"
        />
        <Button
          size="small"
          appearance="subtle"
          icon={<Delete24Regular />}
          onClick={handleSlideRemove}
          disabled={totalSlides <= 1 || isLoading}
          title="スライドを削除"
        />
      </div>
    </div>
  );

  const renderContentItems = () => (
    <div className={styles.contentList}>
      {slide.content.map((content, contentIndex) => (
        <div key={contentIndex} className={styles.contentItem}>
          <Text>•</Text>
          <Input
            className={styles.contentInput}
            value={content}
            onChange={(e) => onUpdateSlideContent(slideIndex, contentIndex, e.target.value)}
            disabled={isLoading}
            placeholder="コンテンツを入力してください"
          />
          <Button
            size="small"
            appearance="subtle"
            icon={<Delete24Regular />}
            onClick={() => handleContentRemove(contentIndex)}
            disabled={slide.content.length <= 1 || isLoading}
            title="この項目を削除"
          />
        </div>
      ))}
      
      <Button
        className={styles.addButton}
        size="small"
        appearance="subtle"
        icon={<Add24Regular />}
        onClick={handleContentAdd}
        disabled={isLoading}
      >
        ポイントを追加
      </Button>
    </div>
  );

  return (
    <AccordionItem value={`slide-${slideIndex}`}>
      <AccordionHeader>
        {renderSlideHeader()}
      </AccordionHeader>
      
      <AccordionPanel>
        <div className={styles.accordionContent}>
          <Field label="スライドタイトル" className={styles.fieldMargin}>
            <Input
              value={slide.title}
              onChange={(e) => onUpdateSlideTitle(slideIndex, e.target.value)}
              disabled={isLoading}
              placeholder="スライドタイトルを入力してください"
            />
          </Field>

          <Field label="コンテンツ">
            {renderContentItems()}
          </Field>
        </div>
      </AccordionPanel>
    </AccordionItem>
  );
};

export default SlideEditor;