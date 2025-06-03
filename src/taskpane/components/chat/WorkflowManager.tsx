// src/taskpane/components/chat/WorkflowManager.tsx - ワークフロー管理コンポーネント
import * as React from "react";
import {
  Button,
  Text,
  Card,
  CardPreview,
  Spinner,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import {
  DocumentAdd24Regular,
  Edit24Regular,
  ArrowRight24Regular,
} from "@fluentui/react-icons";
import { PresentationOutline } from '../types';

export type WorkflowStep = 'chat' | 'outline' | 'generating' | 'completed';

interface WorkflowManagerProps {
  currentStep: WorkflowStep;
  generationProgress: string;
  currentOutline: PresentationOutline | null;
  selectedTheme: 'light' | 'dark' | 'colorful';
  selectedFontSize: 'small' | 'medium' | 'large';
  onStartNewPresentation: () => void;
  onEditOutline: () => void;
}

const useStyles = makeStyles({
  stepIndicator: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
  progressSection: {
    padding: "24px",
    textAlign: "center",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
  },
  completedSection: {
    padding: "24px",
    textAlign: "center",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
  },
  actionButtons: {
    display: 'flex',
    gap: '12px',
    marginTop: '16px',
  },
  presentationCard: {
    maxWidth: "400px",
    width: "100%",
  },
});

export const WorkflowManager: React.FC<WorkflowManagerProps> = ({
  currentStep,
  generationProgress,
  currentOutline,
  selectedTheme,
  selectedFontSize,
  onStartNewPresentation,
  onEditOutline,
}) => {
  const styles = useStyles();

  const getStepTitle = (step: WorkflowStep): string => {
    switch (step) {
      case 'chat': return 'チャット';
      case 'outline': return 'アウトライン編集';
      case 'generating': return 'スライド生成中';
      case 'completed': return '完了';
      default: return '';
    }
  };

  const getStepDescription = (step: WorkflowStep): string => {
    switch (step) {
      case 'chat': 
        return 'プレゼンテーションの要件を入力してください';
      case 'outline': 
        return 'AIが生成したアウトラインを確認・編集してください';
      case 'generating': 
        return 'PowerPointスライドを生成しています';
      case 'completed': 
        return 'スライドの生成が完了しました';
      default: 
        return '';
    }
  };

  const renderStepIndicator = () => (
    <div className={styles.stepIndicator}>
      <Text weight="semibold">
        現在のステップ: {getStepTitle(currentStep)}
      </Text>
      {currentStep === 'outline' && <ArrowRight24Regular />}
      <Text size={300} style={{ marginLeft: 'auto' }}>
        {getStepDescription(currentStep)}
      </Text>
    </div>
  );

  const renderGeneratingStep = () => (
    <div className={styles.progressSection}>
      <Spinner size="large" />
      <Text size={500} weight="semibold">スライドを生成中...</Text>
      <Text>{generationProgress}</Text>
      {currentOutline && (
        <Card className={styles.presentationCard}>
          <CardPreview>
            <Text weight="semibold">生成中のプレゼンテーション</Text>
            <Text>タイトル: {currentOutline.title}</Text>
            <Text>スライド数: {currentOutline.slides.length}枚</Text>
            <Text>テーマ: {selectedTheme.toUpperCase()}</Text>
            <Text>フォントサイズ: {selectedFontSize.toUpperCase()}</Text>
          </CardPreview>
        </Card>
      )}
    </div>
  );

  const renderCompletedStep = () => (
    <div className={styles.completedSection}>
      <Text size={600} weight="semibold">🎉 スライド生成完了！</Text>
      {currentOutline && (
        <Card className={styles.presentationCard}>
          <CardPreview>
            <Text weight="semibold">生成されたプレゼンテーション</Text>
            <Text>タイトル: {currentOutline.title}</Text>
            <Text>スライド数: {currentOutline.slides.length}枚</Text>
            <Text>予想時間: {currentOutline.estimatedDuration}分</Text>
            <Text>適用テーマ: {selectedTheme.toUpperCase()}</Text>
            <Text>フォントサイズ: {selectedFontSize.toUpperCase()}</Text>
          </CardPreview>
        </Card>
      )}
      
      <div className={styles.actionButtons}>
        <Button
          appearance="primary"
          icon={<DocumentAdd24Regular />}
          onClick={onStartNewPresentation}
        >
          新しいプレゼンテーションを作成
        </Button>
        <Button
          appearance="secondary"
          icon={<Edit24Regular />}
          onClick={onEditOutline}
        >
          アウトラインを再編集
        </Button>
      </div>
    </div>
  );

  return (
    <>
      {/* ステップインジケーター */}
      {renderStepIndicator()}

      {/* ステップ別コンテンツ */}
      {currentStep === 'generating' && renderGeneratingStep()}
      {currentStep === 'completed' && renderCompletedStep()}
    </>
  );
};

export default WorkflowManager;