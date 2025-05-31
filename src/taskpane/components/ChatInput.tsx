import * as React from "react";
import { useState } from "react";
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
  Divider
} from "@fluentui/react-components";
import { Send24Regular, Chat24Regular } from "@fluentui/react-icons";

interface ChatMessage {
  id: string;
  content: string;
  timestamp: Date;
  type: 'user' | 'assistant';
}

interface ChatInputProps {
  onSendMessage: (message: string) => Promise<void>;
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
  sendButton: {
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
});

const ChatInput: React.FC<ChatInputProps> = ({ onSendMessage }) => {
  const [message, setMessage] = useState<string>("");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const styles = useStyles();

  const handleSendMessage = async () => {
    if (!message.trim() || isLoading) return;

    const userMessage: ChatMessage = {
      id: Date.now().toString(),
      content: message.trim(),
      timestamp: new Date(),
      type: 'user'
    };

    setMessages(prev => [...prev, userMessage]);
    setMessage("");
    setIsLoading(true);

    try {
      await onSendMessage(userMessage.content);
      
      // 現在は仮の応答を追加（後でOpenAI APIと連携）
      const assistantMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        content: "メッセージを受信しました。OpenAI APIとの連携は次のステップで実装します。",
        timestamp: new Date(),
        type: 'assistant'
      };
      
      setMessages(prev => [...prev, assistantMessage]);
    } catch (error) {
      console.error("Error sending message:", error);
      const errorMessage: ChatMessage = {
        id: (Date.now() + 1).toString(),
        content: "エラーが発生しました。もう一度お試しください。",
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
      
      <div className={styles.chatMessages}>
        {messages.length === 0 ? (
          <div className={styles.emptyState}>
            <Text>PowerPointプレゼンテーションの作成についてお聞かせください。</Text>
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
                <Text>{msg.content}</Text>
              </CardPreview>
            </Card>
          ))
        )}
        {isLoading && (
          <Card className={`${styles.messageCard} ${styles.assistantMessage}`}>
            <CardPreview>
              <Text>入力中...</Text>
            </CardPreview>
          </Card>
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
        
        <Button
          className={styles.sendButton}
          appearance="primary"
          icon={<Send24Regular />}
          onClick={handleSendMessage}
          disabled={!message.trim() || isLoading}
        >
          {isLoading ? "送信中..." : "送信"}
        </Button>
      </div>
    </div>
  );
};

export default ChatInput;