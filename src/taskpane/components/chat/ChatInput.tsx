// src/taskpane/components/chat/ChatInput.tsx - SlideContentGenerator統合版
import * as React from "react";
import { useState, useEffect } from "react";
import { 
  Button, 
  Field, 
  Textarea, 
  Text,
  Card,
  CardHeader,
  CardPreview,
  Divider,
  Spinner,
  MessageBar,
  ProgressBar,
  tokens, 
  makeStyles,
} from "@fluentui/react-components";
import { 
  Send24Regular, 
  Chat24Regular,
} from "@fluentui/react-icons";
import { OpenAIService } from '../../../services/openai.service';
import { PowerPointService } from '../../../services/powerpoint'; 
import { ChatMessage, OpenAISettings, PresentationOutline } from '../types';
import ThemeSettingsComponent from './ThemeSettings';
import WorkflowManager, { WorkflowStep } from './WorkflowManager';
import OutlineEditor from '../outline/OutlineEditor';

interface ChatInputProps {
  onSendMessage: (message: string) => Promise<void>;
  settings: OpenAISettings | null;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    padding: "16px",
  },
  chatContainer: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
  },
  chatHeader: {
    display: "flex",
    alignItems: "center",
    marginBottom: "16px",
    gap: "8px",
  },
  chatMessages: {
    flex: 1,
    overflowY: "auto",
    marginBottom: "16px",
    maxHeight: "300px",
    border: "1px solid " + tokens.colorNeutralStroke1,
    borderRadius: tokens.borderRadiusMedium,
    padding: "8px",
  },
  messageCard: {
    marginBottom: "8px",
  },
  userMessage: {
    backgroundColor: tokens.colorBrandBackground2,
    marginLeft: "20px",
  },
  assistantMessage: {
    backgroundColor: tokens.colorNeutralBackground2,
    marginRight: "20px",
  },
  inputArea: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  textareaField: {
    width: "100%",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    justifyContent: "space-between",
    alignItems: "center",
  },
  primaryButton: {
    marginLeft: "auto",
  },
  timestamp: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  emptyState: {
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    padding: "20px",
  },
  loadingContainer: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px",
  },
  messageContent: {
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  progressSection: {
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
  progressDetails: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  phaseIndicator: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
  },
  errorCard: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    border: "1px solid " + tokens.colorPaletteRedBorder1,
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
});

type GenerationPhase = 'analyzing' | 'detailing' | 'creating';

const ChatInput: React.FC<ChatInputProps> = ({ onSendMessage, settings }) => {
  const [message, setMessage] = useState<string>("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [openAIService, setOpenAIService] = useState<OpenAIService | null>(null);
  const [powerPointService] = useState<PowerPointService>(new PowerPointService());
  const [error, setError] = useState<string>("");
  const [currentStep, setCurrentStep] = useState<WorkflowStep>('chat');
  const [currentOutline, setCurrentOutline] = useState<PresentationOutline | null>(null);
  const [generationProgress, setGenerationProgress] = useState<string>("");
  const [generationPhase, setGenerationPhase] = useState<GenerationPhase>('analyzing');
  const [progressPercentage, setProgressPercentage] = useState<number>(0);
  
  // テーマ設定
  const [selectedTheme, setSelectedTheme] = useState<'light' | 'dark' | 'colorful'>('light');
  const [selectedFontSize, setSelectedFontSize] = useState<'small' | 'medium' | 'large'>('medium');
  const [showThemeSettings, setShowThemeSettings] = useState<boolean>(false);
  
  const styles = useStyles();

  // OpenAI設定が変更されたときにサービスを更新
  useEffect(() => {
    if (settings && settings.apiKey) {
      setOpenAIService(new OpenAIService(settings));
      setError("");
    } else {
      setOpenAIService(null);
    }
  }, [settings]);

  const handleSendMessage = async () => {
    if (!message.trim() || isLoading) return;

    if (!openAIService) {
      setError("OpenAI APIの設定を完了してください。設定タブでAPIキーを入力してください。");
      return;
    }

    const userMessage: ChatMessage = {
      id: Date.now().toString(),
      content: message.trim(),
      timestamp: new Date(),
      type: 'user'
    };

    setMessages(prev => [...prev, userMessage]);
    setMessage("");
    setIsLoading(true);
    setError("");

    try {
      const outline = await openAIService.generateStructuredOutline(userMessage.content);
      
      const assistantMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        content: `✅ アウトラインを生成しました！\n\nタイトル: ${outline.title}\nスライド数: ${outline.slides.length}\n予想時間: ${outline.estimatedDuration}分\n\n「アウトライン編集」タブで内容を確認・編集してください。\n\n💡 スライド生成時は、各スライドのコンテンツがAIによって詳細化され、説明資料として使用できるレベルに拡張されます。`,
        timestamp: new Date(),
        type: 'assistant'
      };
      
      setMessages(prev => [...prev, assistantMessage]);
      setCurrentOutline(outline);
      setCurrentStep('outline');

      await onSendMessage(userMessage.content);
    } catch (error) {
      console.error("Error calling OpenAI API:", error);
      const errorMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        content: `エラーが発生しました: ${error instanceof Error ? error.message : '不明なエラー'}`,
        timestamp: new Date(),
        type: 'assistant'
      };
      setMessages(prev => [...prev, errorMessage]);
      setError(error instanceof Error ? error.message : '不明なエラーが発生しました');
    } finally {
      setIsLoading(false);
    }
  };

  const handleOutlineUpdate = (updatedOutline: PresentationOutline) => {
    setCurrentOutline(updatedOutline);
  };

  const handleRegenerateOutline = async (instruction: string) => {
    if (!openAIService || !currentOutline) return;

    setIsLoading(true);
    try {
      const newOutline = await openAIService.regenerateOutline(currentOutline, instruction);
      setCurrentOutline(newOutline);
      
      const regenerationMessage: ChatMessage = {
        id: Date.now().toString(),
        content: `🔄 アウトラインを再生成しました！\n\n指示: ${instruction}\n\n新しいタイトル: ${newOutline.title}\nスライド数: ${newOutline.slides.length}`,
        timestamp: new Date(),
        type: 'assistant'
      };
      setMessages(prev => [...prev, regenerationMessage]);
    } catch (error) {
      setError(error instanceof Error ? error.message : 'アウトライン再生成でエラーが発生しました');
    } finally {
      setIsLoading(false);
    }
  };

  const handleGenerateSlides = async (outline: PresentationOutline) => {
    if (!outline || !openAIService) return;

    setCurrentStep('generating');
    setIsLoading(true);
    setGenerationProgress("詳細化プロセスを開始します...");
    setProgressPercentage(0);

    try {
      // 詳細な進捗管理機能を使用
      await powerPointService.generateSlidesWithDetailedProgress(
        outline,
        openAIService,
        {
          slideLayout: 'content' as const,
          theme: selectedTheme,
          fontSize: selectedFontSize,
          includeTransitions: false,
          useThemeAwareGeneration: true
        },
        (phase, current, total, message) => {
          setGenerationPhase(phase);
          setGenerationProgress(message);
          
          // フェーズに基づく進捗計算
          let baseProgress = 0;
          switch (phase) {
            case 'analyzing':
              baseProgress = 0;
              break;
            case 'detailing':
              baseProgress = 10;
              break;
            case 'creating':
              baseProgress = 60;
              break;
          }
          
          const phaseProgress = (current / total) * (phase === 'detailing' ? 50 : phase === 'creating' ? 40 : 10);
          setProgressPercentage(baseProgress + phaseProgress);
        }
      );

      setGenerationProgress("✅ スライド生成完了！");
      setProgressPercentage(100);
      setCurrentStep('completed');
      
      const completionMessage: ChatMessage = {
        id: Date.now().toString(),
        content: `🎉 詳細化されたPowerPointスライドの生成が完了しました！\n\n生成されたスライド: ${outline.slides.length}枚\nタイトル: ${outline.title}\nテーマ: ${selectedTheme.toUpperCase()}\nフォントサイズ: ${selectedFontSize.toUpperCase()}\n\n✨ 各スライドのコンテンツはAIによって詳細化され、説明資料として使用できるレベルに拡張されました。\n\n📋 詳細化の特徴:\n• 具体例とデータを含む詳細な説明\n• 前後のスライドとの一貫性を考慮\n• ビジネス現場で実用的な内容\n• 聴衆が自立して理解できるレベル`,
        timestamp: new Date(),
        type: 'assistant'
      };
      setMessages(prev => [...prev, completionMessage]);

    } catch (error) {
      console.error("Error generating slides:", error);
      setError(error instanceof Error ? error.message : 'スライド生成でエラーが発生しました');
      setCurrentStep('outline');
      setProgressPercentage(0);
    } finally {
      setIsLoading(false);
    }
  };

  const handleStartNewPresentation = () => {
    setCurrentStep('chat');
    setCurrentOutline(null);
    setGenerationProgress("");
    setProgressPercentage(0);
    setError("");
  };

  const handleTestTheme = async () => {
    if (!powerPointService) return;
    
    setIsLoading(true);
    setGenerationProgress("テーマテストスライドを生成中...");
    
    try {
      await powerPointService.testThemeApplication();
      setGenerationProgress("テーマテスト完了！各テーマのスライドが作成されました。");
      
      setTimeout(() => {
        setGenerationProgress("");
      }, 3000);
    } catch (error) {
      setError("テーマテストでエラーが発生しました: " + (error instanceof Error ? error.message : '不明なエラー'));
    } finally {
      setIsLoading(false);
    }
  };

  const handleTestDetailedGeneration = async () => {
    if (!openAIService) {
      setError("OpenAI APIの設定を完了してください。");
      return;
    }

    setIsLoading(true);
    setGenerationProgress("詳細化テストを実行中...");
    
    try {
      await powerPointService.testDetailedGeneration(openAIService);
      setGenerationProgress("詳細化テスト完了！テスト用スライドが作成されました。");
      
      setTimeout(() => {
        setGenerationProgress("");
      }, 3000);
    } catch (error) {
      setError("詳細化テストでエラーが発生しました: " + (error instanceof Error ? error.message : '不明なエラー'));
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      handleSendMessage();
    }
  };

  const formatTimestamp = (timestamp: Date) => {
    return timestamp.toLocaleTimeString('ja-JP', { 
      hour: '2-digit', 
      minute: '2-digit' 
    });
  };

  const getPhaseDisplayName = (phase: GenerationPhase): string => {
    switch (phase) {
      case 'analyzing': return '📊 アウトライン分析';
      case 'detailing': return '📝 コンテンツ詳細化';
      case 'creating': return '🎨 スライド作成';
      default: return '処理中';
    }
  };

  const renderChatMessages = () => (
    <div className={styles.chatMessages}>
      {messages.length === 0 ? (
        <div className={styles.emptyState}>
          <Text>PowerPointプレゼンテーションの作成についてお聞かせください。</Text>
          <br />
          <Text size={200}>例: "営業戦略についてのプレゼンテーションを作成してください"</Text>
          <br />
          <Text size={200}>💡 各スライドは自動的に詳細化され、説明資料として使えるレベルになります</Text>
        </div>
      ) : (
        messages.map((msg) => (
          <Card 
            key={msg.id} 
            className={`${styles.messageCard} ${msg.type === 'user' ? styles.userMessage : styles.assistantMessage}`}
          >
            <CardHeader
              header={
                <Text weight="semibold" size={300}>
                  {msg.type === 'user' ? 'あなた' : 'AI アシスタント'}
                </Text>
              }
              description={
                <Text className={styles.timestamp}>
                  {formatTimestamp(msg.timestamp)}
                </Text>
              }
            />
            <CardPreview>
              <Text className={styles.messageContent}>{msg.content}</Text>
            </CardPreview>
          </Card>
        ))
      )}
      {isLoading && currentStep === 'chat' && (
        <div className={styles.loadingContainer}>
          <Spinner size="tiny" />
          <Text>AI がアウトラインを生成中...</Text>
        </div>
      )}
    </div>
  );

  const renderProgressSection = () => {
    if (!isLoading || currentStep !== 'generating') return null;

    return (
      <div className={styles.progressSection}>
        <div className={styles.progressDetails}>
          <div className={styles.phaseIndicator}>
            <Text weight="semibold">
              {getPhaseDisplayName(generationPhase)}
            </Text>
            <Text size={200}>
              ({Math.round(progressPercentage)}%)
            </Text>
          </div>
          
          <ProgressBar value={progressPercentage} max={100} />
          
          <Text size={300}>
            {generationProgress}
          </Text>
          
          {currentOutline && (
            <div style={{ marginTop: '12px' }}>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                処理中: {currentOutline.title} ({currentOutline.slides.length}スライド)
              </Text>
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderChatInput = () => (
    <div className={styles.inputArea}>
      <Field 
        className={styles.textareaField}
        label="メッセージを入力してください"
      >
        <Textarea
          placeholder="例: 営業戦略についてのプレゼンテーションを作成してください"
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          onKeyDown={handleKeyDown}
          rows={3}
          resize="vertical"
          disabled={isLoading}
        />
      </Field>
      
      <div className={styles.buttonGroup}>
        <ThemeSettingsComponent
          selectedTheme={selectedTheme}
          selectedFontSize={selectedFontSize}
          onThemeChange={setSelectedTheme}
          onFontSizeChange={setSelectedFontSize}
          onTestTheme={handleTestTheme}
          showSettings={showThemeSettings}
          onToggleSettings={() => setShowThemeSettings(!showThemeSettings)}
          isLoading={isLoading}
        />
        
        {/* 詳細化テストボタン（開発用） */}
        {process.env.NODE_ENV === 'development' && (
          <Button
            size="small"
            appearance="subtle"
            onClick={handleTestDetailedGeneration}
            disabled={isLoading || !openAIService}
          >
            詳細化テスト
          </Button>
        )}
        
        <Button
          className={styles.primaryButton}
          appearance="primary"
          icon={<Send24Regular />}
          onClick={handleSendMessage}
          disabled={!message.trim() || isLoading || !openAIService}
        >
          {isLoading ? "生成中..." : "アウトライン生成"}
        </Button>
      </div>
    </div>
  );

  return (
    <div className={styles.container}>
      <WorkflowManager
        currentStep={currentStep}
        generationProgress={generationProgress}
        currentOutline={currentOutline}
        selectedTheme={selectedTheme}
        selectedFontSize={selectedFontSize}
        onStartNewPresentation={handleStartNewPresentation}
        onEditOutline={() => setCurrentStep('outline')}
      />

      {error && (
        <div className={styles.errorCard}>
          <Text weight="semibold" style={{ color: tokens.colorPaletteRedForeground1 }}>
            エラーが発生しました
          </Text>
          <Text style={{ color: tokens.colorPaletteRedForeground1 }}>
            {error}
          </Text>
        </div>
      )}

      {generationProgress && currentStep !== 'generating' && (
        <MessageBar intent="info" style={{ marginBottom: '16px' }}>
          {generationProgress}
        </MessageBar>
      )}

      {/* 詳細化進捗セクション */}
      {renderProgressSection()}

      {/* チャットセクション */}
      {currentStep === 'chat' && (
        <div className={styles.chatContainer}>
          <div className={styles.chatHeader}>
            <Chat24Regular />
            <Text weight="semibold" size={400}>PowerPoint Concierge</Text>
            <Text size={200} style={{ marginLeft: '8px', color: tokens.colorNeutralForeground3 }}>
              (詳細化機能搭載)
            </Text>
          </div>

          {renderChatMessages()}
          <Divider />
          {renderChatInput()}
        </div>
      )}

      {/* アウトライン編集セクション */}
      {currentStep === 'outline' && (
        <OutlineEditor
          outline={currentOutline}
          onOutlineUpdate={handleOutlineUpdate}
          onGenerateSlides={handleGenerateSlides}
          onRegenerateOutline={handleRegenerateOutline}
          isLoading={isLoading}
        />
      )}
    </div>
  );
};

export default ChatInput;