// src/taskpane/components/ChatInput.tsx - 修正版（テーマオプション追加）
import * as React from "react";
import { useState, useEffect } from "react";
import { 
  Button, 
  Field, 
  Textarea, 
  tokens, 
  makeStyles,
  Card,
  CardHeader,
  CardPreview,
  Text,
  Divider,
  Spinner,
  MessageBar,
  Dropdown,
  Option,
  Label,
} from "@fluentui/react-components";
import { 
  Send24Regular, 
  Chat24Regular, 
  DocumentAdd24Regular,
  Edit24Regular,
  ArrowRight24Regular,
  Settings24Regular,
  Play24Regular,
} from "@fluentui/react-icons";
import { OpenAIService } from '../../services/openai.service';
import { PowerPointService } from '../../services/powerpoint'; 
import { ChatMessage, OpenAISettings } from './types';
import OutlineEditor, { PresentationOutline } from './OutlineEditor';

interface ChatInputProps {
  onSendMessage: (message: string) => Promise<void>;
  settings: OpenAISettings | null;
}

type WorkflowStep = 'chat' | 'outline' | 'generating' | 'completed';

// テーマとサイズのオプション
const themeOptions = [
  { key: 'light', text: 'ライト（白背景）' },
  { key: 'dark', text: 'ダーク（黒背景）' },
  { key: 'colorful', text: 'カラフル（多色）' },
];

const fontSizeOptions = [
  { key: 'small', text: '小（12-32pt）' },
  { key: 'medium', text: '中（16-42pt）' },
  { key: 'large', text: '大（18-48pt）' },
];

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    padding: "16px",
  },
  stepTabs: {
    marginBottom: "16px",
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
  // 新規追加：テーマ設定セクション
  themeSection: {
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
  themeControls: {
    display: "flex",
    gap: "16px",
    alignItems: "end",
    marginTop: "12px",
  },
  themePreview: {
    display: "flex",
    gap: "8px",
    marginTop: "8px",
  },
  previewBox: {
    width: "20px",
    height: "20px",
    borderRadius: "4px",
    border: "1px solid " + tokens.colorNeutralStroke1,
  },
  testSection: {
    marginTop: "12px",
    paddingTop: "12px",
    borderTop: "1px solid " + tokens.colorNeutralStroke2,
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
  
  // 新規追加：テーマ設定
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

  // テーマプレビューカラーを取得
  const getThemePreviewColors = (theme: string) => {
    switch (theme) {
      case 'light':
        return { bg: '#FFFFFF', text: '#000000', accent: '#0078D4' };
      case 'dark':
        return { bg: '#1F1F1F', text: '#FFFFFF', accent: '#0078D4' };
      case 'colorful':
        return { bg: '#FFFFFF', text: '#323130', accent: '#FF6B35' };
      default:
        return { bg: '#FFFFFF', text: '#000000', accent: '#0078D4' };
    }
  };

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
      // 構造化されたアウトラインを生成
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
      
      // チャットにメッセージを追加
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
      // テーマ設定を含むBulkSlideData形式に変換
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

      // 新しいPowerPointServiceを使用してスライドを一括生成
      await powerPointService.generateBulkSlides(
        bulkData,
        (current, total, slideName) => {
          setGenerationProgress(`スライド ${current}/${total} を生成中: ${slideName}`);
        }
      );

      setGenerationProgress("スライド生成完了！");
      setCurrentStep('completed');
      
      // 完了メッセージをチャットに追加
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
    // チャット履歴はクリアしない（参考として残す）
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

  const getStepTitle = (step: WorkflowStep) => {
    switch (step) {
      case 'chat': return 'チャット';
      case 'outline': return 'アウトライン編集';
      case 'generating': return 'スライド生成中';
      case 'completed': return '完了';
      default: return '';
    }
  };

  return (
    <div className={styles.container}>
      {/* ステップインジケーター */}
      <div className={styles.stepIndicator}>
        <Text weight="semibold">
          現在のステップ: {getStepTitle(currentStep)}
        </Text>
        {currentStep === 'outline' && (
          <ArrowRight24Regular />
        )}
      </div>

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
            <Button
              size="small"
              appearance="subtle"
              icon={<Settings24Regular />}
              onClick={() => setShowThemeSettings(!showThemeSettings)}
            >
              テーマ設定
            </Button>
          </div>

          {/* テーマ設定セクション */}
          {showThemeSettings && (
            <div className={styles.themeSection}>
              <Label weight="semibold">プレゼンテーションの外観設定</Label>
              
              <div className={styles.themeControls}>
                <Field label="テーマ">
                  <Dropdown
                    value={selectedTheme}
                    selectedOptions={[selectedTheme]}
                    onOptionSelect={(_, data) => {
                      if (data.optionValue) {
                        setSelectedTheme(data.optionValue as 'light' | 'dark' | 'colorful');
                      }
                    }}
                  >
                    {themeOptions.map(option => (
                      <Option key={option.key} value={option.key}>
                        {option.text}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="フォントサイズ">
                  <Dropdown
                    value={selectedFontSize}
                    selectedOptions={[selectedFontSize]}
                    onOptionSelect={(_, data) => {
                      if (data.optionValue) {
                        setSelectedFontSize(data.optionValue as 'small' | 'medium' | 'large');
                      }
                    }}
                  >
                    {fontSizeOptions.map(option => (
                      <Option key={option.key} value={option.key}>
                        {option.text}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>

              {/* テーマプレビュー */}
              <div style={{ marginTop: '12px' }}>
                <Label size="small">プレビュー:</Label>
                <div className={styles.themePreview}>
                  {(() => {
                    const colors = getThemePreviewColors(selectedTheme);
                    return (
                      <>
                        <div 
                          className={styles.previewBox}
                          style={{ backgroundColor: colors.bg }}
                          title="背景色"
                        />
                        <div 
                          className={styles.previewBox}
                          style={{ backgroundColor: colors.text }}
                          title="テキスト色"
                        />
                        <div 
                          className={styles.previewBox}
                          style={{ backgroundColor: colors.accent }}
                          title="アクセント色"
                        />
                      </>
                    );
                  })()}
                  <Text size={200} style={{ marginLeft: '8px' }}>
                    {selectedTheme.toUpperCase()} / {selectedFontSize.toUpperCase()}
                  </Text>
                </div>
              </div>

              {/* テーマテスト機能 */}
              <div className={styles.testSection}>
                <Button
                  size="small"
                  appearance="secondary"
                  icon={<Play24Regular />}
                  onClick={handleTestTheme}
                  disabled={isLoading}
                >
                  テーマをテスト（3つのテーマで同じスライドを作成）
                </Button>
              </div>
            </div>
          )}
          
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

          <Divider />

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
              <Text size={200}>
                選択中: {selectedTheme.toUpperCase()}テーマ / {selectedFontSize.toUpperCase()}サイズ
              </Text>
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

      {/* スライド生成中セクション */}
      {currentStep === 'generating' && (
        <div className={styles.progressSection}>
          <Spinner size="large" />
          <Text size={500} weight="semibold">スライドを生成中...</Text>
          <Text>{generationProgress}</Text>
          {currentOutline && (
            <Card>
              <CardPreview>
                <Text>タイトル: {currentOutline.title}</Text>
                <Text>スライド数: {currentOutline.slides.length}枚</Text>
                <Text>テーマ: {selectedTheme.toUpperCase()}</Text>
                <Text>フォントサイズ: {selectedFontSize.toUpperCase()}</Text>
              </CardPreview>
            </Card>
          )}
        </div>
      )}

      {/* 完了セクション */}
      {currentStep === 'completed' && (
        <div className={styles.completedSection}>
          <Text size={600} weight="semibold">🎉 スライド生成完了！</Text>
          {currentOutline && (
            <Card>
              <CardHeader header={<Text weight="semibold">生成されたプレゼンテーション</Text>} />
              <CardPreview>
                <Text>タイトル: {currentOutline.title}</Text>
                <Text>スライド数: {currentOutline.slides.length}枚</Text>
                <Text>予想時間: {currentOutline.estimatedDuration}分</Text>
                <Text>適用テーマ: {selectedTheme.toUpperCase()}</Text>
                <Text>フォントサイズ: {selectedFontSize.toUpperCase()}</Text>
              </CardPreview>
            </Card>
          )}
          
          <div style={{ display: 'flex', gap: '12px', marginTop: '16px' }}>
            <Button
              appearance="primary"
              icon={<DocumentAdd24Regular />}
              onClick={handleStartNewPresentation}
            >
              新しいプレゼンテーションを作成
            </Button>
            <Button
              appearance="secondary"
              icon={<Edit24Regular />}
              onClick={() => setCurrentStep('outline')}
            >
              アウトラインを再編集
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};

export default ChatInput;