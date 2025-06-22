import React, { useState, useEffect } from "react";
import {
  Card,
  CardHeader,
  Button,
  Text,
  Title3,
  Input,
  Textarea,
  Field,
  Radio,
  RadioGroup,
  Dropdown,
  Option,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  Add24Regular,
  Save24Regular,
  Delete24Regular,
  Info24Regular,
  Star24Regular,
} from "@fluentui/react-icons";
import {
  TemplateInfo,
  TemplateMetadata,
  TemplateCategory,
  TemplateRegistrationRequest,
} from "../../../services/powerpoint/template-types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  templateList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxHeight: "400px",
    overflowY: "auto",
  },
  templateItem: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: "4px",
  },
  templateDetails: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    flexGrow: 1,
  },
  templateActions: {
    display: "flex",
    gap: "8px",
  },
  formGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "16px",
    marginBottom: "16px",
  },
  formField: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  dialogContent: {
    minWidth: "500px",
  },
  successMessage: {
    marginBottom: "16px",
  },
});

interface TemplateManagerProps {
  onTemplateCreated?: (template: TemplateInfo) => void;
  onTemplateDeleted?: (templateId: string) => void;
}

const TemplateManager: React.FC<TemplateManagerProps> = ({
  onTemplateCreated,
  onTemplateDeleted,
}) => {
  const styles = useStyles();
  const [templates, setTemplates] = useState<TemplateInfo[]>([]);
  const [isCreateDialogOpen, setIsCreateDialogOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [successMessage, setSuccessMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  // フォーム状態
  const [formData, setFormData] = useState({
    name: "",
    description: "",
    category: "business" as TemplateCategory,
    presentationStyle: "formal" as const,
    targetAudience: "general" as const,
    purpose: "report" as const,
    tags: "",
  });

  useEffect(() => {
    loadTemplates();
  }, []);

  const loadTemplates = async () => {
    try {
      const { PowerPointService } = await import("../../../services/powerpoint");
      const powerPointService = new PowerPointService();
      const library = powerPointService.getTemplateLibrary();
      setTemplates(library.templates);
    } catch (error) {
      console.error("Failed to load templates:", error);
      setErrorMessage("テンプレートの読み込みに失敗しました");
    }
  };

  const handleCreateTemplate = async () => {
    if (!formData.name.trim()) {
      setErrorMessage("テンプレート名を入力してください");
      return;
    }

    setIsLoading(true);
    setErrorMessage("");

    try {
      const { PowerPointService } = await import("../../../services/powerpoint");
      const powerPointService = new PowerPointService();

      const metadata: Partial<TemplateRegistrationRequest["metadata"]> = {
        presentationStyle: formData.presentationStyle,
        targetAudience: formData.targetAudience,
        purpose: formData.purpose,
        tags: formData.tags
          .split(",")
          .map((tag) => tag.trim())
          .filter((tag) => tag),
      };

      const newTemplate = await powerPointService.registerCurrentPresentationAsTemplate(
        formData.name,
        formData.description,
        metadata
      );

      newTemplate.category = formData.category;

      setTemplates((prev) => [...prev, newTemplate]);
      setSuccessMessage(`テンプレート「${newTemplate.name}」を作成しました`);
      setIsCreateDialogOpen(false);
      resetForm();

      if (onTemplateCreated) {
        onTemplateCreated(newTemplate);
      }

      // 成功メッセージを3秒後に消す
      setTimeout(() => setSuccessMessage(""), 3000);
    } catch (error) {
      console.error("Failed to create template:", error);
      setErrorMessage(`テンプレートの作成に失敗しました: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteTemplate = async (templateId: string) => {
    if (!confirm("このテンプレートを削除しますか？")) {
      return;
    }

    try {
      // テンプレートをローカルリストから削除
      setTemplates((prev) => prev.filter((t) => t.id !== templateId));
      setSuccessMessage("テンプレートを削除しました");

      if (onTemplateDeleted) {
        onTemplateDeleted(templateId);
      }

      // 成功メッセージを3秒後に消す
      setTimeout(() => setSuccessMessage(""), 3000);
    } catch (error) {
      console.error("Failed to delete template:", error);
      setErrorMessage("テンプレートの削除に失敗しました");
    }
  };

  const resetForm = () => {
    setFormData({
      name: "",
      description: "",
      category: "business",
      presentationStyle: "formal",
      targetAudience: "general",
      purpose: "report",
      tags: "",
    });
  };

  const handleFormChange = (field: string, value: string) => {
    setFormData((prev) => ({ ...prev, [field]: value }));
  };

  const renderTemplateItem = (template: TemplateInfo) => (
    <div key={template.id} className={styles.templateItem}>
      <div className={styles.templateDetails}>
        <Text weight="medium">{template.name}</Text>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          {template.description}
        </Text>
        <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
          カテゴリ: {template.category} | スタイル: {template.metadata.presentationStyle} |
          利用回数: {template.metadata.usageCount}
        </Text>
      </div>

      <div className={styles.templateActions}>
        <Button
          size="small"
          icon={<Star24Regular />}
          onClick={() => {
            // テンプレートの詳細表示や評価機能
            console.log("Template details:", template);
          }}
        >
          詳細
        </Button>

        <Button
          size="small"
          icon={<Delete24Regular />}
          onClick={() => handleDeleteTemplate(template.id)}
          disabled={template.id.startsWith("default-")} // デフォルトテンプレートは削除不可
        >
          削除
        </Button>
      </div>
    </div>
  );

  return (
    <div className={styles.container}>
      {/* 使い方ガイド */}
      <Card style={{ marginBottom: "16px", backgroundColor: tokens.colorNeutralBackground3 }}>
        <CardHeader>
          <Text weight="semibold">📋 テンプレートの用意の仕方</Text>
        </CardHeader>
        <div style={{ padding: "16px" }}>
          <Text size={200}>
            <strong>Step 1:</strong> PowerPointで理想的なプレゼンテーションを作成
            <br />
            • デザイン、レイアウト、色合いを設定
            <br />
            • スライドマスターやテーマを調整
            <br />
            • 数枚のサンプルスライドを作成
            <br />
            <br />
            <strong>Step 2:</strong> 「現在のプレゼンテーションをテンプレート化」をクリック
            <br />
            • システムが自動的にデザインパターンを分析
            <br />
            • レイアウト構造とスタイルを抽出
            <br />
            <br />
            <strong>Step 3:</strong> テンプレート情報を入力
            <br />
            • わかりやすい名前と説明を入力
            <br />
            • カテゴリ、スタイル、対象聴衆を選択
            <br />
            • 検索用のタグを設定
            <br />
            <br />
            <strong>Step 4:</strong> 今後の利用
            <br />
            • アウトライン生成後に推奨テンプレートとして表示
            <br />
            • AI詳細化の代わりにテンプレート最適化を実行
            <br />• 一貫性のあるデザインでプレゼンテーションを生成
          </Text>
        </div>
      </Card>

      <div className={styles.header}>
        <Title3>テンプレート管理</Title3>

        <Dialog open={isCreateDialogOpen}>
          <DialogTrigger disableButtonEnhancement>
            <Button
              icon={<Add24Regular />}
              appearance="primary"
              onClick={() => setIsCreateDialogOpen(true)}
            >
              現在のプレゼンテーションをテンプレート化
            </Button>
          </DialogTrigger>

          <DialogSurface className={styles.dialogContent}>
            <DialogBody>
              <DialogTitle>新しいテンプレートを作成</DialogTitle>

              <DialogContent>
                {errorMessage && (
                  <MessageBar intent="error">
                    <MessageBarBody>{errorMessage}</MessageBarBody>
                  </MessageBar>
                )}

                <div className={styles.formGrid}>
                  <Field label="テンプレート名 *" required>
                    <Input
                      value={formData.name}
                      onChange={(_, data) => handleFormChange("name", data.value)}
                      placeholder="例: 営業提案書テンプレート"
                    />
                  </Field>

                  <Field label="カテゴリ">
                    <Dropdown
                      value={formData.category}
                      onOptionSelect={(_, data) =>
                        handleFormChange("category", data.optionValue as string)
                      }
                    >
                      <Option value="business">ビジネス</Option>
                      <Option value="academic">学術</Option>
                      <Option value="marketing">マーケティング</Option>
                      <Option value="technical">技術</Option>
                      <Option value="minimal">ミニマル</Option>
                      <Option value="creative">クリエイティブ</Option>
                      <Option value="corporate">コーポレート</Option>
                    </Dropdown>
                  </Field>
                </div>

                <Field label="説明">
                  <Textarea
                    value={formData.description}
                    onChange={(_, data) => handleFormChange("description", data.value)}
                    placeholder="このテンプレートの特徴や用途を説明してください"
                    rows={3}
                  />
                </Field>

                <div className={styles.formGrid}>
                  <Field label="プレゼンテーションスタイル">
                    <RadioGroup
                      value={formData.presentationStyle}
                      onChange={(_, data) => handleFormChange("presentationStyle", data.value)}
                    >
                      <Radio value="formal" label="フォーマル" />
                      <Radio value="casual" label="カジュアル" />
                      <Radio value="technical" label="技術的" />
                      <Radio value="creative" label="クリエイティブ" />
                    </RadioGroup>
                  </Field>

                  <Field label="対象聴衆">
                    <RadioGroup
                      value={formData.targetAudience}
                      onChange={(_, data) => handleFormChange("targetAudience", data.value)}
                    >
                      <Radio value="executive" label="経営層" />
                      <Radio value="technical" label="技術者" />
                      <Radio value="general" label="一般" />
                      <Radio value="academic" label="学術" />
                      <Radio value="sales" label="営業" />
                    </RadioGroup>
                  </Field>
                </div>

                <div className={styles.formGrid}>
                  <Field label="目的">
                    <RadioGroup
                      value={formData.purpose}
                      onChange={(_, data) => handleFormChange("purpose", data.value)}
                    >
                      <Radio value="pitch" label="提案・ピッチ" />
                      <Radio value="report" label="報告書" />
                      <Radio value="training" label="研修・教育" />
                      <Radio value="marketing" label="マーケティング" />
                      <Radio value="analysis" label="分析・調査" />
                    </RadioGroup>
                  </Field>

                  <Field label="タグ（カンマ区切り）">
                    <Input
                      value={formData.tags}
                      onChange={(_, data) => handleFormChange("tags", data.value)}
                      placeholder="例: 営業, 提案書, 月次報告"
                    />
                  </Field>
                </div>
              </DialogContent>
            </DialogBody>

            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button
                  appearance="secondary"
                  onClick={() => {
                    setIsCreateDialogOpen(false);
                    setErrorMessage("");
                    resetForm();
                  }}
                >
                  キャンセル
                </Button>
              </DialogTrigger>

              <Button
                appearance="primary"
                icon={<Save24Regular />}
                onClick={handleCreateTemplate}
                disabled={isLoading || !formData.name.trim()}
              >
                {isLoading ? "作成中..." : "テンプレートを作成"}
              </Button>
            </DialogActions>
          </DialogSurface>
        </Dialog>
      </div>

      {successMessage && (
        <MessageBar intent="success" className={styles.successMessage}>
          <MessageBarBody>{successMessage}</MessageBarBody>
        </MessageBar>
      )}

      <div>
        <Text size={300} weight="medium">
          登録済みテンプレート ({templates.length}個)
        </Text>
        {templates.length === 0 ? (
          <div
            style={{
              textAlign: "center",
              padding: "32px",
              color: tokens.colorNeutralForeground3,
            }}
          >
            <Info24Regular style={{ marginBottom: "8px" }} />
            <Text>テンプレートがありません</Text>
            <Text size={200} style={{ display: "block", marginTop: "4px" }}>
              現在のプレゼンテーションをテンプレートとして保存できます
            </Text>
          </div>
        ) : (
          <div className={styles.templateList}>{templates.map(renderTemplateItem)}</div>
        )}
      </div>
    </div>
  );
};

export default TemplateManager;
