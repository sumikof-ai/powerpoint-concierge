/* global describe, it, beforeEach, expect */
import { SlideFactory } from "./SlideFactory";

describe("SlideFactory", () => {
  let slideFactory: SlideFactory;

  beforeEach(() => {
    slideFactory = new SlideFactory();
  });

  describe("determineSlideLayout", () => {
    it("テンプレートが指定されている場合、titleスライドではtitleレイアウトを返す", () => {
      const result = slideFactory.determineSlideLayout("title", 100, undefined, true);
      expect(result).toBe("title");
    });

    it("テンプレートが指定されていない場合、titleスライドではblankレイアウトを返す", () => {
      const result = slideFactory.determineSlideLayout("title", 100, undefined, false);
      expect(result).toBe("blank");
    });

    it("hasTemplateが未定義の場合、titleスライドではblankレイアウトを返す", () => {
      const result = slideFactory.determineSlideLayout("title", 100, undefined, undefined);
      expect(result).toBe("blank");
    });

    it("contentスライドでは常にcontentレイアウトを返す（テンプレート有無に関係なく）", () => {
      const resultWithTemplate = slideFactory.determineSlideLayout("content", 100, undefined, true);
      const resultWithoutTemplate = slideFactory.determineSlideLayout(
        "content",
        100,
        undefined,
        false
      );

      expect(resultWithTemplate).toBe("content");
      expect(resultWithoutTemplate).toBe("content");
    });

    it("contentスライドでコンテンツ量が多い場合はtwoContentレイアウトを返す", () => {
      const result = slideFactory.determineSlideLayout("content", 350, undefined, false);
      expect(result).toBe("twoContent");
    });

    it("defaultLayoutが指定されている場合はそれを優先する（contentスライド）", () => {
      const result = slideFactory.determineSlideLayout("content", 100, "comparison", false);
      expect(result).toBe("comparison");
    });

    it("conclusionスライドでは常にcontentレイアウトを返す", () => {
      const result = slideFactory.determineSlideLayout("conclusion", 100, undefined, true);
      expect(result).toBe("content");
    });

    it("不明なスライドタイプではcontentレイアウトを返す", () => {
      // @ts-ignore - テスト用に不正な値を渡す
      const result = slideFactory.determineSlideLayout("unknown", 100, undefined, true);
      expect(result).toBe("content");
    });
  });

  describe("calculateContentAmount", () => {
    it("タイトルとコンテンツの文字数を合計して返す", () => {
      const slideData = {
        title: "テストタイトル", // 7文字
        content: ["項目1", "項目2"], // 3 + 3 = 6文字
        slideType: "content" as const,
        speakerNotes: "",
      };

      // プライベートメソッドにアクセスするため@ts-ignoreを使用
      // @ts-ignore
      const result = slideFactory.calculateContentAmount(slideData);
      expect(result).toBe(13); // 7 + 6 = 13
    });

    it("コンテンツが空の場合はタイトルの文字数のみ返す", () => {
      const slideData = {
        title: "テストタイトル", // 7文字
        content: [],
        slideType: "title" as const,
        speakerNotes: "",
      };

      // @ts-ignore
      const result = slideFactory.calculateContentAmount(slideData);
      expect(result).toBe(7);
    });
  });

  describe("validateSlideContent", () => {
    it("正常なスライドデータでは有効と判定される", () => {
      const slideData = {
        title: "適切な長さのタイトル",
        content: ["項目1", "項目2", "項目3"],
        slideType: "content" as const,
        speakerNotes: "",
      };

      const result = slideFactory.validateSlideContent(slideData);
      expect(result.isValid).toBe(true);
      expect(result.warnings).toHaveLength(0);
    });

    it("タイトルが長すぎる場合は警告が出る", () => {
      const longTitle = "あ".repeat(101); // 101文字
      const slideData = {
        title: longTitle,
        content: ["項目1"],
        slideType: "content" as const,
        speakerNotes: "",
      };

      const result = slideFactory.validateSlideContent(slideData);
      expect(result.isValid).toBe(false);
      expect(result.warnings).toContain("タイトルが長すぎます（100文字以内を推奨）");
    });

    it("コンテンツ項目が多すぎる場合は警告が出る", () => {
      const slideData = {
        title: "タイトル",
        content: Array(8).fill("項目"), // 8項目
        slideType: "content" as const,
        speakerNotes: "",
      };

      const result = slideFactory.validateSlideContent(slideData);
      expect(result.isValid).toBe(false);
      expect(result.warnings).toContain("コンテンツ項目が多すぎます（7項目以内を推奨）");
      expect(result.suggestions).toContain("複数のスライドに分割することを検討してください");
    });

    it("個別のコンテンツ項目が長すぎる場合は警告が出る", () => {
      const longContent = "あ".repeat(151); // 151文字
      const slideData = {
        title: "タイトル",
        content: [longContent],
        slideType: "content" as const,
        speakerNotes: "",
      };

      const result = slideFactory.validateSlideContent(slideData);
      expect(result.isValid).toBe(false);
      expect(result.warnings).toContain("項目 1 が長すぎます（150文字以内を推奨）");
    });

    it("titleスライド以外でコンテンツが空の場合は警告が出る", () => {
      const slideData = {
        title: "タイトル",
        content: [],
        slideType: "content" as const,
        speakerNotes: "",
      };

      const result = slideFactory.validateSlideContent(slideData);
      expect(result.isValid).toBe(false);
      expect(result.warnings).toContain("コンテンツが空です");
    });
  });

  describe("suggestOptimizations", () => {
    it("コンテンツ量が多い場合は分割を提案する", () => {
      const slideData = {
        title: "あ".repeat(100), // 100文字
        content: Array(10).fill("あ".repeat(50)), // 50文字 × 10 = 500文字
        slideType: "content" as const,
        speakerNotes: "",
      };

      const suggestions = slideFactory.suggestOptimizations(slideData);
      expect(suggestions).toContain(
        "コンテンツ量が多いため、2カラムレイアウトまたは複数スライドへの分割を推奨"
      );
    });

    it("タイトルスライドでコンテンツが多い場合はシンプル化を提案する", () => {
      const slideData = {
        title: "タイトル",
        content: ["項目1", "項目2", "項目3"], // 3項目
        slideType: "title" as const,
        speakerNotes: "",
      };

      const suggestions = slideFactory.suggestOptimizations(slideData);
      expect(suggestions).toContain("タイトルスライドはシンプルに保つことを推奨（2項目以内）");
    });

    it("まとめスライドでコンテンツが多い場合は要点を絞ることを提案する", () => {
      const slideData = {
        title: "まとめ",
        content: Array(6).fill("項目"), // 6項目
        slideType: "conclusion" as const,
        speakerNotes: "",
      };

      const suggestions = slideFactory.suggestOptimizations(slideData);
      expect(suggestions).toContain("まとめスライドは要点を絞ることを推奨（5項目以内）");
    });
  });

  describe("getLayoutTemplates", () => {
    it("すべてのレイアウトテンプレートを返す", () => {
      const templates = slideFactory.getLayoutTemplates();

      expect(templates).toHaveProperty("title");
      expect(templates).toHaveProperty("content");
      expect(templates).toHaveProperty("twoContent");
      expect(templates).toHaveProperty("comparison");

      // タイトルテンプレートの構造確認
      expect(templates.title).toHaveProperty("titlePosition");
      expect(templates.title).toHaveProperty("subtitlePosition");
      expect(templates.title.titlePosition).toMatchObject({
        left: 75,
        top: 150,
        width: 600,
        height: 150,
      });
    });
  });
});
