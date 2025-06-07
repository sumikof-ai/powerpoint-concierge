// src/taskpane/components/chat/WorkflowManager.tsx - 詳細化機能対応版
import * as React from "react";
import {
  Button,
  Text,
  Card,
  CardPreview,
  Spinner,
  Badge,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import {
  DocumentAdd24Regular,
  Edit24Regular,
  ArrowRight24Regular,
  ArrowRight16Regular,
  CheckmarkCircle24Regular,
  Settings24Regular,
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
    flexDirection: "column",
    gap: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
  stepInfo: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  stepChain: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    flexWrap: "wrap",
    justifyContent: "center",
  },
  stepBadge: {
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "nowrap",
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
    flexDirection: 'column',
    gap: '8px',
    marginTop: '16px',
    width: '100%',
    '@media (min-width: 400px)': {
      flexDirection: 'row',
      gap: '12px',
    },
  },
  actionButton: {
    flex: 1,
    minWidth: '0',
  },
  presentationCard: {
    maxWidth: "100%",
    width: "100%",
  },
  enhancementHighlight: {
    backgroundColor: tokens.colorBrandBackground2,
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "12px",
  },
  featureList: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    textAlign: "left",
    marginTop: "8px",
  },
  featureItem: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    lineHeight: "1.4",
  },
  featureIcon: {
    marginTop: "2px",
    flexShrink: 0,
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
      case 'generating': return 'AI詳細化 & スライド生成中';
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
        return '各スライドを詳細化し、PowerPointを生成しています';
      case 'completed': 
        return '詳細化されたスライドの生成が完了しました';
      default: 
        return '';
    }
  };

  const renderStepChain = () => {
    const steps: { key: WorkflowStep; label: string }[] = [
      { key: 'chat', label: 'チャット' },
      { key: 'outline', label: 'アウトライン' },
      { key: 'generating', label: '詳細化' },
      { key: 'completed', label: '完了' }
    ];

    return (
      <div className={styles.stepChain}>
        {steps.map((step, index) => (
          <React.Fragment key={step.key}>
            <Badge 
              appearance={
                currentStep === step.key ? 'filled' : 
                steps.findIndex(s => s.key === currentStep) > index ? 'ghost' : 'tint'
              }
              color={
                currentStep === step.key ? 'brand' :
                steps.findIndex(s => s.key === currentStep) > index ? 'success' : 'subtle'
              }
              className={styles.stepBadge}
            >
              {step.label}
            </Badge>
            {index < steps.length - 1 && <ArrowRight16Regular />}
          </React.Fragment>
        ))}
      </div>
    );
  };

  const renderStepIndicator = () => (
    <div className={styles.stepIndicator}>
      <div className={styles.stepInfo}>
        <Text weight="semibold">
          現在のステップ: {getStepTitle(currentStep)}
        </Text>
        <Text size={300}>
          {getStepDescription(currentStep)}
        </Text>
      </div>
      {renderStepChain()}
    </div>
  );

  const renderGeneratingStep = () => (
    <div className={styles.progressSection}>
      <Spinner size="large" />
      <Text size={500} weight="semibold">AI詳細化 & スライド生成中...</Text>
      <Text>{generationProgress}</Text>
      
      {currentOutline && (
        <Card className={styles.presentationCard}>
          <CardPreview>
            <Text weight="semibold">生成中のプレゼンテーション</Text>
            <Text>タイトル: {currentOutline.title}</Text>
            <Text>スライド数: {currentOutline.slides.length}枚</Text>
            <Text>テーマ: {selectedTheme.toUpperCase()}</Text>
            <Text>フォントサイズ: {selectedFontSize.toUpperCase()}</Text>
            
            <div className={styles.enhancementHighlight}>
              <Text weight="semibold" size={300}>🚀 AI詳細化機能</Text>
              <div className={styles.featureList}>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>各スライドを個別に詳細化</Text>
                </div>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>説明資料レベルのコンテンツ</Text>
                </div>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>前後スライドとの一貫性を保持</Text>
                </div>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>具体例・データを自動追加</Text>
                </div>
              </div>
            </div>
          </CardPreview>
        </Card>
      )}
    </div>
  );

  const renderCompletedStep = () => (
    <div className={styles.completedSection}>
      <Text size={600} weight="semibold">🎉 詳細化スライド生成完了！</Text>
      
      {currentOutline && (
        <Card className={styles.presentationCard}>
          <CardPreview>
            <Text weight="semibold">生成されたプレゼンテーション</Text>
            <Text>タイトル: {currentOutline.title}</Text>
            <Text>スライド数: {currentOutline.slides.length}枚</Text>
            <Text>予想時間: {currentOutline.estimatedDuration}分</Text>
            <Text>適用テーマ: {selectedTheme.toUpperCase()}</Text>
            <Text>フォントサイズ: {selectedFontSize.toUpperCase()}</Text>
            
            <div className={styles.enhancementHighlight}>
              <Text weight="semibold" size={300}>✨ 詳細化の成果</Text>
              <div className={styles.featureList}>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>全スライドが説明資料レベルに詳細化</Text>
                </div>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>聴衆の自立理解を促進する内容</Text>
                </div>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>ビジネス現場で即使用可能</Text>
                </div>
                <div className={styles.featureItem}>
                  <CheckmarkCircle24Regular className={styles.featureIcon} />
                  <Text size={200}>一貫性のある高品質なデザイン</Text>
                </div>
              </div>
            </div>
          </CardPreview>
        </Card>
      )}
      
      <div className={styles.actionButtons}>
        <Button
          appearance="primary"
          icon={<DocumentAdd24Regular />}
          onClick={onStartNewPresentation}
          className={styles.actionButton}
        >
          新しいプレゼンテーション
        </Button>
        <Button
          appearance="secondary"
          icon={<Edit24Regular />}
          onClick={onEditOutline}
          className={styles.actionButton}
        >
          アウトライン編集
        </Button>
        <Button
          appearance="subtle"
          icon={<Settings24Regular />}
          onClick={() => {
            // 設定画面への遷移や詳細設定の表示
            console.log('詳細設定を開く');
          }}
          className={styles.actionButton}
        >
          詳細設定
        </Button>
      </div>
      
      <div style={{ marginTop: '16px', textAlign: 'center' }}>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          💡 PowerPoint内でスライドの編集・調整が可能です
        </Text>
      </div>
    </div>
  );

  const renderOutlineStepInfo = () => {
    if (currentStep !== 'outline') return null;

    return (
      <Card style={{ marginBottom: '16px', backgroundColor: tokens.colorNeutralBackground3 }}>
        <CardPreview style={{ padding: '16px' }}>
          <Text weight="semibold" style={{ marginBottom: '8px' }}>
            📋 アウトライン編集のポイント
          </Text>
          <div className={styles.featureList}>
            <div className={styles.featureItem}>
              <Text size={200}>• スライドの順序や内容を自由に調整できます</Text>
            </div>
            <div className={styles.featureItem}>
              <Text size={200}>• 「スライド生成」実行時に各スライドが自動詳細化されます</Text>
            </div>
            <div className={styles.featureItem}>
              <Text size={200}>• 簡潔なキーワードでも詳細な説明資料に変換されます</Text>
            </div>
            <div className={styles.featureItem}>
              <Text size={200}>• AI再生成で別のアプローチでアウトラインを作り直せます</Text>
            </div>
          </div>
        </CardPreview>
      </Card>
    );
  };

  const renderChatStepInfo = () => {
    if (currentStep !== 'chat') return null;

    return (
      <Card style={{ marginBottom: '16px', backgroundColor: tokens.colorNeutralBackground3 }}>
        <CardPreview style={{ padding: '16px' }}>
          <Text weight="semibold" style={{ marginBottom: '8px' }}>
            🚀 PowerPoint Concierge の特徴
          </Text>
          <div className={styles.featureList}>
            <div className={styles.featureItem}>
              <CheckmarkCircle24Regular className={styles.featureIcon} />
              <Text size={200}>AIによる3段階生成（アウトライン→詳細化→PowerPoint）</Text>
            </div>
            <div className={styles.featureItem}>
              <CheckmarkCircle24Regular className={styles.featureIcon} />
              <Text size={200}>説明資料として使える詳細なコンテンツ</Text>
            </div>
            <div className={styles.featureItem}>
              <CheckmarkCircle24Regular className={styles.featureIcon} />
              <Text size={200}>テーマ設定（色・フォント）の自動適用</Text>
            </div>
            <div className={styles.featureItem}>
              <CheckmarkCircle24Regular className={styles.featureIcon} />
              <Text size={200}>テンプレート選択（レイアウト）も選択可能</Text>
            </div>
          </div>
        </CardPreview>
      </Card>
    );
  };

  return (
    <>
      {/* ステップインジケーター */}
      {renderStepIndicator()}

      {/* ステップ別の情報カード */}
      {renderChatStepInfo()}
      {renderOutlineStepInfo()}

      {/* ステップ別コンテンツ */}
      {currentStep === 'generating' && renderGeneratingStep()}
      {currentStep === 'completed' && renderCompletedStep()}
    </>
  );
};

export default WorkflowManager;