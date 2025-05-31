import * as React from "react";
import { useState } from "react";
import {
  Tab,
  TabList,
  TabValue,
  SelectTabData,
  SelectTabEvent,
  makeStyles
} from "@fluentui/react-components";
import { Chat24Regular, Settings24Regular } from "@fluentui/react-icons";
import Header from "./Header";
import ChatInput from "./ChatInput";
import Settings from "./Settings";

interface AppProps {
  title: string;
}

interface OpenAISettings {
  apiKey: string;
  baseUrl: string;
  model: string;
  temperature: number;
  maxTokens: number;
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

  const onTabSelect = (event: SelectTabEvent,data: SelectTabData) => {
    console.log(event);
    setSelectedTab(data.value);
  };

  const handleSendMessage = async (message: string) => {
    // TODO: OpenAI APIとの連携を実装
    console.log("Sending message:", message);
    console.log("Current settings:", openAISettings);
    
    // 現在は仮の処理として、プロミスを返すだけ
    return new Promise<void>((resolve) => {
      setTimeout(() => {
        resolve();
      }, 1000);
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
          <Tab id="settings" value="settings" icon={<Settings24Regular />}>
            設定
          </Tab>
        </TabList>

        <div className={styles.tabContent}>
          {selectedTab === "chat" && (
            <ChatInput onSendMessage={handleSendMessage} />
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