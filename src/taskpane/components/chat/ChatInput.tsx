// src/taskpane/components/chat/ChatInput.tsx - リファクタリング版メインコンポーネント
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
});

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
        content: `✅ アウトラインを生成しました！\n\nタイトル: ${outline.title}\nスライド数: ${outline.slides.length}\n予想時間: ${outline.estimatedDuration}分\n\n「アウトライン編集」タブで内容を確認・編集してください。`,
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
    if (!outline) return;

    setCurrentStep('generating');
    setIsLoading(true);
    setGenerationProgress("スライド生成を開始します...");

    try {
      const bulkData = {
        slides: outline.slides.map(slide => ({
          title: slide.title,
          content: slide.content,
          slideType: slide.slideType,
          speakerNotes: slide.speakerNotes
        })),
        options: {
          slideLayout: 'content' as const,
          theme: selectedTheme,
          fontSize: selectedFontSize,
          includeTransitions: false,
          useThemeAwareGeneration: true
        }
      };

      await powerPointService.generateBulkSlides(
        bulkData,
        (current, total, slideName) => {
          setGenerationProgress(`スライド ${current}/${total} を生成中: ${slideName}`);
        }
      );

      setGenerationProgress("スライド生成完了！");
      setCurrentStep('completed');
      
      const completionMessage: ChatMessage = {
        id: Date.now().toString(),
        content: `🎉 PowerPointスライドの生成が完了しました！\n\n生成されたスライド: ${outline.slides.length}枚\nタイトル: ${outline.title}\nテーマ: ${selectedTheme.toUpperCase()}\nフォントサイズ: ${selectedFontSize.toUpperCase()}`,
        timestamp: new Date(),
        type: 'assistant'
      };
      setMessages(prev => [...prev, completionMessage]);

    } catch (error) {
      console.error("Error generating slides:", error);
      setError(error instanceof Error ? error.message : 'スライド生成でエラーが発生しました');
      setCurrentStep('outline');
    } finally {
      setIsLoading(false);
    }
  };

  const handleStartNewPresentation = () => {
    setCurrentStep('chat');
    setCurrentOutline(null);
    setGenerationProgress("");
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

  const renderChatMessages = () => (
    <div className={styles.chatMessages}>
      {messages.length === 0 ? (
        <div className={styles.emptyState}>
          <Text>PowerPointプレゼンテーションの作成についてお聞かせください。</Text>
          <br />
          <Text size={200}>例: "営業戦略についてのプレゼンテーションを作成してください"</Text>
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
        <MessageBar intent="error" style={{ marginBottom: '16px' }}>
          {error}
        </MessageBar>
      )}

      {generationProgress && (
        <MessageBar intent="info" style={{ marginBottom: '16px' }}>
          {generationProgress}
        </MessageBar>
      )}

      {/* チャットセクション */}
      {currentStep === 'chat' && (
        <div className={styles.chatContainer}>
          <div className={styles.chatHeader}>
            <Chat24Regular />
            <Text weight="semibold" size={400}>PowerPoint Concierge</Text>
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