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
} from "@fluentui/react-components";
import { Send24Regular, Chat24Regular, DocumentAdd24Regular } from "@fluentui/react-icons";
import { OpenAIService } from '../../services/openai.service';
import { PowerPointService } from '../../services/powerpoint.service';
import { ChatMessage, OpenAISettings } from './types';

interface ChatInputProps {
  onSendMessage: (message: string) => Promise<void>;
  settings: OpenAISettings | null;
}

const useStyles = makeStyles({
  chatContainer: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    padding: "16px",
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
    maxHeight: "400px",
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
    alignSelf: "flex-end",
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
      // OpenAI APIを呼び出してアウトラインを生成
      const response = await openAIService.generateOutline(userMessage.content);
      
      const assistantMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        content: response,
        timestamp: new Date(),
        type: 'assistant'
      };
      
      setMessages(prev => [...prev, assistantMessage]);

      // 通知
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

  const handleAddToPowerPoint = async (content: string) => {
    if (!content.trim()) return;

    setIsLoading(true);
    try {
      // OpenAIの応答をPowerPointのテキストボックスとして追加
      await powerPointService.addTextBox(content);
      
      const successMessage: ChatMessage = {
        id: Date.now().toString(),
        content: "✅ PowerPointにテキストボックスを追加しました！",
        timestamp: new Date(),
        type: 'assistant'
      };
      setMessages(prev => [...prev, successMessage]);
    } catch (error) {
      console.error("Error adding to PowerPoint:", error);
      const errorMessage: ChatMessage = {
        id: Date.now().toString(),
        content: `PowerPointへの追加でエラーが発生しました: ${error instanceof Error ? error.message : '不明なエラー'}`,
        timestamp: new Date(),
        type: 'assistant'
      };
      setMessages(prev => [...prev, errorMessage]);
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

  return (
    <div className={styles.chatContainer}>
      <div className={styles.chatHeader}>
        <Chat24Regular />
        <Text weight="semibold" size={400}>PowerPoint Concierge</Text>
      </div>

      {error && (
        <MessageBar intent="error" className="mb-4">
          {error}
        </MessageBar>
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
                action={
                  msg.type === 'assistant' && msg.content && !msg.content.startsWith('✅') && !msg.content.startsWith('エラー') ? (
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<DocumentAdd24Regular />}
                      onClick={() => handleAddToPowerPoint(msg.content)}
                      disabled={isLoading}
                    >
                      PowerPointに追加
                    </Button>
                  ) : null
                }
              />
              <CardPreview>
                <Text className={styles.messageContent}>{msg.content}</Text>
              </CardPreview>
            </Card>
          ))
        )}
        {isLoading && (
          <div className={styles.loadingContainer}>
            <Spinner size="tiny" />
            <Text>AI が応答を生成中...</Text>
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
          <Button
            appearance="primary"
            icon={<Send24Regular />}
            onClick={handleSendMessage}
            disabled={!message.trim() || isLoading || !openAIService}
          >
            {isLoading ? "送信中..." : "送信"}
          </Button>
        </div>
      </div>
    </div>
  );
};

export default ChatInput;