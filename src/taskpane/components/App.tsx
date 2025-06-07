import * as React from "react";
import { useState, useEffect } from "react";
import {
  Tab,
  TabList,
  TabValue,
  SelectTabData,
  SelectTabEvent,
  makeStyles
} from "@fluentui/react-components";
import { Chat24Regular, Settings24Regular, Apps24Regular } from "@fluentui/react-icons";
import Header from "./Header";
import ChatInput from "./chat/ChatInput";
import Settings from "./Settings";
import TemplateManager from "./template/TemplateManager";
import { OpenAISettings } from "./types";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
  },
  tabContainer: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    padding: "16px",
  },
  tabContent: {
    flex: 1,
    paddingTop: "16px",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = useState<TabValue>("chat");
  const [openAISettings, setOpenAISettings] = useState<OpenAISettings | null>(null);

  // コンポーネントマウント時に保存された設定を読み込み
  useEffect(() => {
    const loadSettings = () => {
      try {
        const savedSettings = localStorage.getItem('powerpoint-concierge-settings');
        if (savedSettings) {
          const parsed = JSON.parse(savedSettings);
          setOpenAISettings(parsed);
        }
      } catch (error) {
        console.error("Failed to load saved settings:", error);
      }
    };

    loadSettings();
  }, []);

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    console.log(event);
    setSelectedTab(data.value);
  };

  const handleSendMessage = async (message: string) => {
    console.log("Sending message:", message);
    console.log("Current settings:", openAISettings);
    
    // この関数は現在ChatInputコンポーネント内で処理されているため、
    // ここでは簡単なログ出力のみ行う
    return new Promise<void>((resolve) => {
      setTimeout(() => {
        resolve();
      }, 100);
    });
  };

  const handleSettingsChange = (settings: OpenAISettings) => {
    setOpenAISettings(settings);
    console.log("Settings updated:", settings);
  };

  return (
    <div className={styles.root}>
      <Header 
        logo="assets/logo-filled.png" 
        title={props.title} 
        message="PowerPoint Concierge" 
      />
      
      <div className={styles.tabContainer}>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab id="chat" value="chat" icon={<Chat24Regular />}>
            チャット
          </Tab>
          <Tab id="templates" value="templates" icon={<Apps24Regular />}>
            テンプレート
          </Tab>
          <Tab id="settings" value="settings" icon={<Settings24Regular />}>
            設定
          </Tab>
        </TabList>

        <div className={styles.tabContent}>
          {selectedTab === "chat" && (
            <ChatInput 
              onSendMessage={handleSendMessage}
              settings={openAISettings}
            />
          )}
          {selectedTab === "templates" && (
            <TemplateManager
              onTemplateCreated={(template) => {
                console.log('Template created:', template);
              }}
              onTemplateDeleted={(templateId) => {
                console.log('Template deleted:', templateId);
              }}
            />
          )}
          {selectedTab === "settings" && (
            <Settings onSettingsChange={handleSettingsChange} />
          )}
        </div>
      </div>
    </div>
  );
};

export default App;