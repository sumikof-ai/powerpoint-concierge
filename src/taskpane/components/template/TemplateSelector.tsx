import React, { useState, useEffect } from "react";
import {
  Card,
  CardHeader,
  Button,
  Text,
  Title3,
  Subtitle2,
  Badge,
  Spinner,
  SearchBox,
  Dropdown,
  Option,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  Search24Regular,
  Filter24Regular,
  Star24Filled,
  Star24Regular,
  CheckmarkCircle24Filled,
} from "@fluentui/react-icons";
import {
  TemplateInfo,
  TemplateRecommendation,
  TemplateCategory,
} from "../../../services/powerpoint/template-types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  searchSection: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  filtersSection: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
  templateGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))",
    gap: "16px",
    maxHeight: "400px",
    overflowY: "auto",
  },
  templateCard: {
    cursor: "pointer",
    transition: "transform 0.2s ease, box-shadow 0.2s ease",
    ":hover": {
      transform: "translateY(-2px)",
      boxShadow: tokens.shadow8,
    },
  },
  selectedCard: {
    border: `2px solid ${tokens.colorBrandBackground}`,
    boxShadow: tokens.shadow8,
  },
  templateHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: "8px",
  },
  templateInfo: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  templateTags: {
    display: "flex",
    gap: "4px",
    flexWrap: "wrap",
    marginTop: "8px",
  },
  templateStats: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginTop: "8px",
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
  },
  recommendationReasons: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "4px",
  },
  loadingContainer: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    height: "200px",
  },
  emptyState: {
    textAlign: "center",
    padding: "32px",
    color: tokens.colorNeutralForeground3,
  },
});

interface TemplateSelectorProps {
  userInput?: string;
  onTemplateSelect: (template: TemplateInfo) => void;
  onTemplateRecommendations?: (recommendations: TemplateRecommendation[]) => void;
  selectedTemplateId?: string;
  isVisible: boolean;
}

const TemplateSelector: React.FC<TemplateSelectorProps> = ({
  userInput = "",
  onTemplateSelect,
  onTemplateRecommendations,
  selectedTemplateId,
  isVisible,
}) => {
  const styles = useStyles();
  const [recommendations, setRecommendations] = useState<TemplateRecommendation[]>([]);
  const [allTemplates, setAllTemplates] = useState<TemplateInfo[]>([]);
  const [filteredTemplates, setFilteredTemplates] = useState<TemplateInfo[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState<TemplateCategory | "all">("all");
  const [showRecommendationsOnly, setShowRecommendationsOnly] = useState(true);

  useEffect(() => {
    if (isVisible) {
      loadTemplates();
      if (userInput) {
        loadRecommendations();
      }
    }
  }, [isVisible, userInput]);

  useEffect(() => {
    filterTemplates();
  }, [allTemplates, recommendations, searchQuery, selectedCategory, showRecommendationsOnly]);

  const loadTemplates = async () => {
    try {
      // PowerPointServiceからテンプレートライブラリを取得
      const { PowerPointService } = await import("../../../services/powerpoint");
      const powerPointService = new PowerPointService();
      const library = powerPointService.getTemplateLibrary();
      setAllTemplates(library.templates);
    } catch (error) {
      console.error("Failed to load templates:", error);
    }
  };

  const loadRecommendations = async () => {
    if (!userInput.trim()) return;

    setIsLoading(true);
    try {
      const { PowerPointService } = await import("../../../services/powerpoint");
      const powerPointService = new PowerPointService();
      const recs = await powerPointService.getTemplateRecommendations(userInput);
      setRecommendations(recs);

      if (onTemplateRecommendations) {
        onTemplateRecommendations(recs);
      }
    } catch (error) {
      console.error("Failed to load recommendations:", error);
      setRecommendations([]);
    } finally {
      setIsLoading(false);
    }
  };

  const filterTemplates = () => {
    let templates = showRecommendationsOnly ? recommendations.map((r) => r.template) : allTemplates;

    // 検索フィルター
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase();
      templates = templates.filter(
        (template) =>
          template.name.toLowerCase().includes(query) ||
          template.description?.toLowerCase().includes(query) ||
          template.metadata.tags.some((tag) => tag.toLowerCase().includes(query))
      );
    }

    // カテゴリフィルター
    if (selectedCategory !== "all") {
      templates = templates.filter((template) => template.category === selectedCategory);
    }

    setFilteredTemplates(templates);
  };

  const handleTemplateClick = (template: TemplateInfo) => {
    onTemplateSelect(template);
  };

  const getTemplateScore = (template: TemplateInfo): number => {
    const recommendation = recommendations.find((r) => r.template.id === template.id);
    return recommendation?.score || 0;
  };

  const getRecommendationReasons = (template: TemplateInfo): string[] => {
    const recommendation = recommendations.find((r) => r.template.id === template.id);
    return recommendation?.reasoning || [];
  };

  const renderTemplateCard = (template: TemplateInfo) => {
    const isSelected = template.id === selectedTemplateId;
    const score = getTemplateScore(template);
    const reasons = getRecommendationReasons(template);
    const isRecommended = score > 0;

    return (
      <Card
        key={template.id}
        className={`${styles.templateCard} ${isSelected ? styles.selectedCard : ""}`}
        onClick={() => handleTemplateClick(template)}
      >
        <CardHeader>
          <div className={styles.templateHeader}>
            <div className={styles.templateInfo}>
              <Title3>{template.name}</Title3>
              <Text size={200}>{template.description}</Text>
            </div>
            <div>
              {isSelected && <CheckmarkCircle24Filled color={tokens.colorBrandBackground} />}
              {isRecommended && !isSelected && (
                <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                  <Star24Filled color={tokens.colorPaletteYellowForeground1} />
                  <Text size={100}>{Math.round(score * 100)}%</Text>
                </div>
              )}
            </div>
          </div>
        </CardHeader>

        <div style={{ padding: "12px" }}>
          <div className={styles.templateTags}>
            <Badge color="brand" size="small">
              {template.category}
            </Badge>
            <Badge color="informative" size="small">
              {template.metadata.presentationStyle}
            </Badge>
            <Badge color="subtle" size="small">
              {template.metadata.targetAudience}
            </Badge>
          </div>

          <div className={styles.templateStats}>
            <Text size={100}>
              スライド数: {template.metadata.slideCount} | 用途: {template.metadata.purpose} |
              利用回数: {template.metadata.usageCount}
            </Text>
          </div>

          {reasons.length > 0 && (
            <div className={styles.recommendationReasons}>
              <Text size={100} weight="medium">
                推奨理由:
              </Text>
              {reasons.map((reason, index) => (
                <Text key={index} size={100} style={{ display: "block", marginTop: "2px" }}>
                  • {reason}
                </Text>
              ))}
            </div>
          )}
        </div>
      </Card>
    );
  };

  if (!isVisible) {
    return null;
  }

  return (
    <div className={styles.container}>
      <div className={styles.searchSection}>
        <SearchBox
          placeholder="テンプレートを検索..."
          value={searchQuery}
          onChange={(_, data) => setSearchQuery(data.value)}
          contentBefore={<Search24Regular />}
          style={{ flexGrow: 1 }}
        />

        <Button
          icon={<Filter24Regular />}
          appearance={showRecommendationsOnly ? "primary" : "secondary"}
          onClick={() => setShowRecommendationsOnly(!showRecommendationsOnly)}
        >
          {showRecommendationsOnly ? "推奨のみ" : "すべて表示"}
        </Button>
      </div>

      <div className={styles.filtersSection}>
        <Dropdown
          placeholder="カテゴリで絞り込み"
          value={selectedCategory === "all" ? "すべて" : selectedCategory}
          onOptionSelect={(_, data) =>
            setSelectedCategory(data.optionValue as TemplateCategory | "all")
          }
        >
          <Option value="all">すべて</Option>
          <Option value="business">ビジネス</Option>
          <Option value="academic">学術</Option>
          <Option value="marketing">マーケティング</Option>
          <Option value="technical">技術</Option>
          <Option value="minimal">ミニマル</Option>
          <Option value="creative">クリエイティブ</Option>
          <Option value="corporate">コーポレート</Option>
        </Dropdown>

        {userInput && (
          <Button appearance="secondary" onClick={loadRecommendations} disabled={isLoading}>
            推奨を更新
          </Button>
        )}
      </div>

      {isLoading ? (
        <div className={styles.loadingContainer}>
          <Spinner label="テンプレートを分析中..." />
        </div>
      ) : filteredTemplates.length > 0 ? (
        <div className={styles.templateGrid}>{filteredTemplates.map(renderTemplateCard)}</div>
      ) : (
        <div className={styles.emptyState}>
          <Text size={300}>
            {showRecommendationsOnly && recommendations.length === 0
              ? "ユーザー入力に基づく推奨テンプレートがありません"
              : "条件に一致するテンプレートが見つかりません"}
          </Text>
          {showRecommendationsOnly && (
            <Button
              appearance="secondary"
              onClick={() => setShowRecommendationsOnly(false)}
              style={{ marginTop: "16px" }}
            >
              すべてのテンプレートを表示
            </Button>
          )}
        </div>
      )}

      {recommendations.length > 0 && showRecommendationsOnly && (
        <Text size={200} style={{ textAlign: "center", color: tokens.colorNeutralForeground3 }}>
          AI分析により{recommendations.length}個のテンプレートを推奨しています
        </Text>
      )}
    </div>
  );
};

export default TemplateSelector;
