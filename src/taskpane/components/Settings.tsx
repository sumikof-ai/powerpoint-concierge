import * as React from "react";
import { useState, useEffect } from "react";
import {
  Button,
  Field,
  Input,
  Text,
  Card,
  CardHeader,
  CardPreview,
  Dropdown,
  Option,
  tokens,
  makeStyles,
  Switch,
  Label,
  Divider,
} from "@fluentui/react-components";
import { Settings24Regular, Save24Regular } from "@fluentui/react-icons";

interface OpenAISettings {
  apiKey: string;
  baseUrl: string;
  model: string;
  temperature: number;
  maxTokens: number;
}

interface SettingsProps {
  onSettingsChange: (settings: OpenAISettings) => void;
}

const useStyles = makeStyles({
  settingsContainer: {
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  settingsHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
  },
  settingCard: {
    padding: "16px",
  },
  fieldGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  advancedSettings: {
    marginTop: "8px",
  },
  saveButton: {
    marginTop: "16px",
    alignSelf: "flex-end",
  },
  warning: {
    backgroundColor: tokens.colorPaletteYellowBackground1,
    border: "1px solid " + tokens.colorPaletteYellowBorder1,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    marginBottom: "16px",
  },
});

const defaultSettings: OpenAISettings = {
  apiKey: "",
  baseUrl: "https://api.openai.com/v1",
  model: "gpt-4",
  temperature: 0.7,
  maxTokens: 2000,
};

const availableModels = [
  { key: "gpt-4", text: "GPT-4" },
  { key: "gpt-4-turbo-preview", text: "GPT-4 Turbo" },
  { key: "gpt-3.5-turbo", text: "GPT-3.5 Turbo" },
];

const Settings: React.FC<SettingsProps> = ({ onSettingsChange }) => {
  const [settings, setSettings] = useState<OpenAISettings>(defaultSettings);
  const [showAdvanced, setShowAdvanced] = useState<boolean>(false);
  const [isSaved, setIsSaved] = useState<boolean>(false);
  const styles = useStyles();

  // コンポーネントマウント時に保存された設定を読み込み
  useEffect(() => {
    const savedSettings = localStorage.getItem("powerpoint-concierge-settings");
    if (savedSettings) {
      try {
        const parsed = JSON.parse(savedSettings);
        setSettings({ ...defaultSettings, ...parsed });
      } catch (error) {
        console.error("Failed to parse saved settings:", error);
      }
    }
  }, []);

  const handleSettingChange = (key: keyof OpenAISettings, value: string | number) => {
    setSettings((prev) => ({
      ...prev,
      [key]: value,
    }));
    setIsSaved(false);
  };

  const handleSaveSettings = () => {
    try {
      localStorage.setItem("powerpoint-concierge-settings", JSON.stringify(settings));
      onSettingsChange(settings);
      setIsSaved(true);
      setTimeout(() => setIsSaved(false), 2000);
    } catch (error) {
      console.error("Failed to save settings:", error);
    }
  };

  const isValidUrl = (url: string) => {
    try {
      new URL(url);
      return true;
    } catch {
      return false;
    }
  };

  const canSave = settings.apiKey.trim() !== "" && isValidUrl(settings.baseUrl);

  return (
    <div className={styles.settingsContainer}>
      <div className={styles.settingsHeader}>
        <Settings24Regular />
        <Text weight="semibold" size={400}>
          設定
        </Text>
      </div>

      <div className={styles.warning}>
        <Text size={300}>
          ⚠️ APIキーは安全に管理してください。このアドインはローカルストレージに設定を保存します。
        </Text>
      </div>

      <Card className={styles.settingCard}>
        <CardHeader header={<Text weight="semibold">OpenAI API設定</Text>} />
        <CardPreview>
          <div className={styles.fieldGroup}>
            <Field label="APIキー" required>
              <Input
                type="password"
                placeholder="sk-..."
                value={settings.apiKey}
                onChange={(e) => handleSettingChange("apiKey", e.target.value)}
              />
            </Field>

            <Field label="ベースURL">
              <Input
                placeholder="https://api.openai.com/v1"
                value={settings.baseUrl}
                onChange={(e) => handleSettingChange("baseUrl", e.target.value)}
              />
            </Field>

            <Field label="モデル">
              <Dropdown
                value={settings.model}
                selectedOptions={[settings.model]}
                onOptionSelect={(e, data) => {
                  if (data.optionValue) {
                    console.log(e);
                    handleSettingChange("model", data.optionValue);
                  }
                }}
              >
                {availableModels.map((model) => (
                  <Option key={model.key} value={model.key}>
                    {model.text}
                  </Option>
                ))}
              </Dropdown>
            </Field>
          </div>
        </CardPreview>
      </Card>

      <div>
        <Label>
          <Switch
            checked={showAdvanced}
            onChange={(e) => setShowAdvanced(e.currentTarget.checked)}
          />
          詳細設定を表示
        </Label>
      </div>

      {showAdvanced && (
        <Card className={styles.settingCard}>
          <CardHeader header={<Text weight="semibold">詳細設定</Text>} />
          <CardPreview>
            <div className={styles.fieldGroup}>
              <Field
                label={`Temperature (${settings.temperature})`}
                hint="0.0-2.0の範囲で指定。高いほど創造的な回答になります。"
              >
                <Input
                  type="text"
                  min="0"
                  max="2"
                  step="0.1"
                  value={settings.temperature.toString()}
                  onChange={(e) => handleSettingChange("temperature", parseFloat(e.target.value))}
                />
              </Field>

              <Field label="最大トークン数" hint="回答の最大長を制御します。">
                <Input
                  type="number"
                  min="100"
                  max="4000"
                  value={settings.maxTokens.toString()}
                  onChange={(e) => handleSettingChange("maxTokens", parseInt(e.target.value))}
                />
              </Field>
            </div>
          </CardPreview>
        </Card>
      )}

      <Divider />

      <Button
        className={styles.saveButton}
        appearance="primary"
        icon={<Save24Regular />}
        onClick={handleSaveSettings}
        disabled={!canSave}
      >
        {isSaved ? "保存済み" : "設定を保存"}
      </Button>
    </div>
  );
};

export default Settings;
