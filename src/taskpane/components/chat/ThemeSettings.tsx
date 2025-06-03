// src/taskpane/components/chat/ThemeSettings.tsx - テーマ設定コンポーネント
import * as React from "react";
import {
  Button,
  Dropdown,
  Option,
  Field,
  Label,
  Text,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import {
  Settings24Regular,
  Play24Regular,
} from "@fluentui/react-icons";

export interface ThemeSettings {
  theme: 'light' | 'dark' | 'colorful';
  fontSize: 'small' | 'medium' | 'large';
}

interface ThemeSettingsProps {
  selectedTheme: 'light' | 'dark' | 'colorful';
  selectedFontSize: 'small' | 'medium' | 'large';
  onThemeChange: (theme: 'light' | 'dark' | 'colorful') => void;
  onFontSizeChange: (size: 'small' | 'medium' | 'large') => void;
  onTestTheme: () => void;
  showSettings: boolean;
  onToggleSettings: () => void;
  isLoading: boolean;
}

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
  toggleButton: {
    marginLeft: "auto",
  },
  currentSettings: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
});

export const ThemeSettingsComponent: React.FC<ThemeSettingsProps> = ({
  selectedTheme,
  selectedFontSize,
  onThemeChange,
  onFontSizeChange,
  onTestTheme,
  showSettings,
  onToggleSettings,
  isLoading,
}) => {
  const styles = useStyles();

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

  const renderThemePreview = () => {
    const colors = getThemePreviewColors(selectedTheme);
    return (
      <div className={styles.themePreview}>
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
        <Text size={200} style={{ marginLeft: '8px' }}>
          {selectedTheme.toUpperCase()} / {selectedFontSize.toUpperCase()}
        </Text>
      </div>
    );
  };

  return (
    <>
      {/* テーマ設定トグルボタン */}
      <Button
        className={styles.toggleButton}
        size="small"
        appearance="subtle"
        icon={<Settings24Regular />}
        onClick={onToggleSettings}
      >
        テーマ設定
      </Button>

      {/* 現在の設定表示 */}
      {!showSettings && (
        <Text className={styles.currentSettings}>
          選択中: {selectedTheme.toUpperCase()}テーマ / {selectedFontSize.toUpperCase()}サイズ
        </Text>
      )}

      {/* テーマ設定パネル */}
      {showSettings && (
        <div className={styles.themeSection}>
          <Label weight="semibold">プレゼンテーションの外観設定</Label>
          
          <div className={styles.themeControls}>
            <Field label="テーマ">
              <Dropdown
                value={selectedTheme}
                selectedOptions={[selectedTheme]}
                onOptionSelect={(_, data) => {
                  if (data.optionValue) {
                    onThemeChange(data.optionValue as 'light' | 'dark' | 'colorful');
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
                    onFontSizeChange(data.optionValue as 'small' | 'medium' | 'large');
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
            {renderThemePreview()}
          </div>

          {/* テーマテスト機能 */}
          <div className={styles.testSection}>
            <Button
              size="small"
              appearance="secondary"
              icon={<Play24Regular />}
              onClick={onTestTheme}
              disabled={isLoading}
            >
              テーマをテスト（3つのテーマで同じスライドを作成）
            </Button>
          </div>
        </div>
      )}
    </>
  );
};

export default ThemeSettingsComponent;