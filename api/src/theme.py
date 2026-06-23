"""
テーマ定義モジュール
プレゼンテーションの配色・フォント・スタイルを管理する。
"""

from dataclasses import dataclass, field
from pptx.util import Pt
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_color: str) -> RGBColor:
    """HEXカラーコードをRGBColorに変換する。"""
    hex_color = hex_color.lstrip("#")
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


@dataclass
class ThemeColors:
    """テーマの配色定義。"""
    bg_primary: str        # メイン背景色
    bg_secondary: str      # セカンダリ背景（グラデーション終点など）
    title_color: str       # タイトルテキスト色
    text_color: str        # 本文テキスト色
    text_light: str        # 薄いテキスト（サブ情報）
    accent: str            # アクセントカラー（バー、ライン、アイコン）
    accent_secondary: str  # セカンダリアクセント
    overlay: str           # 半透明オーバーレイ色
    overlay_opacity: float # オーバーレイの不透明度 (0.0~1.0)


@dataclass
class ThemeFonts:
    """テーマのフォント定義。"""
    title_ja: str = "游ゴシック"
    title_en: str = "Segoe UI"
    body_ja: str = "游ゴシック"
    body_en: str = "Segoe UI"
    title_size: Pt = field(default_factory=lambda: Pt(36))
    subtitle_size: Pt = field(default_factory=lambda: Pt(18))
    heading_size: Pt = field(default_factory=lambda: Pt(28))
    body_size: Pt = field(default_factory=lambda: Pt(16))
    caption_size: Pt = field(default_factory=lambda: Pt(12))
    page_number_size: Pt = field(default_factory=lambda: Pt(10))


@dataclass
class Theme:
    """プレゼンテーション全体のテーマ。"""
    name: str
    description: str
    colors: ThemeColors
    fonts: ThemeFonts = field(default_factory=ThemeFonts)

    def rgb(self, color_attr: str) -> RGBColor:
        """テーマカラー属性名からRGBColorを取得する。"""
        return hex_to_rgb(getattr(self.colors, color_attr))


# =========================================================================
# テーマプリセット
# =========================================================================

THEMES: dict[str, Theme] = {

    "midnight": Theme(
        name="midnight",
        description="テクノロジー・未来・AI・サイバー系に最適",
        colors=ThemeColors(
            bg_primary="#0d1117",
            bg_secondary="#161b22",
            title_color="#f0f6fc",
            text_color="#c9d1d9",
            text_light="#8b949e",
            accent="#58a6ff",
            accent_secondary="#1f6feb",
            overlay="#0d1117",
            overlay_opacity=0.70,
        ),
    ),

    "elegance": Theme(
        name="elegance",
        description="高級感・ラグジュアリー・プレミアム系に最適",
        colors=ThemeColors(
            bg_primary="#1a1a2e",
            bg_secondary="#16213e",
            title_color="#ffffff",
            text_color="#e0e0e0",
            text_light="#a0a0b0",
            accent="#e94560",
            accent_secondary="#c23152",
            overlay="#1a1a2e",
            overlay_opacity=0.72,
        ),
    ),

    "corporate": Theme(
        name="corporate",
        description="ビジネス・会議・提案書・信頼感を重視",
        colors=ThemeColors(
            bg_primary="#ffffff",
            bg_secondary="#f8fafc",
            title_color="#1e293b",
            text_color="#475569",
            text_light="#94a3b8",
            accent="#2563eb",
            accent_secondary="#1d4ed8",
            overlay="#1e293b",
            overlay_opacity=0.65,
        ),
    ),

    "nature": Theme(
        name="nature",
        description="環境・ヘルスケア・ウェルネス・オーガニック系に最適",
        colors=ThemeColors(
            bg_primary="#f0fdf4",
            bg_secondary="#ecfdf5",
            title_color="#14532d",
            text_color="#365314",
            text_light="#6b7280",
            accent="#22c55e",
            accent_secondary="#16a34a",
            overlay="#14532d",
            overlay_opacity=0.65,
        ),
    ),

    "creative": Theme(
        name="creative",
        description="デザイン・アート・クリエイティブ系に最適",
        colors=ThemeColors(
            bg_primary="#faf5ff",
            bg_secondary="#f5f3ff",
            title_color="#581c87",
            text_color="#6b21a8",
            text_light="#9ca3af",
            accent="#a855f7",
            accent_secondary="#9333ea",
            overlay="#581c87",
            overlay_opacity=0.68,
        ),
    ),

    "warm": Theme(
        name="warm",
        description="教育・コミュニティ・食・ホスピタリティ系に最適",
        colors=ThemeColors(
            bg_primary="#fffbeb",
            bg_secondary="#fef3c7",
            title_color="#78350f",
            text_color="#92400e",
            text_light="#9ca3af",
            accent="#f59e0b",
            accent_secondary="#d97706",
            overlay="#78350f",
            overlay_opacity=0.65,
        ),
    ),

    "minimal": Theme(
        name="minimal",
        description="クリーン・モダン・ミニマリスト・汎用",
        colors=ThemeColors(
            bg_primary="#fafafa",
            bg_secondary="#f4f4f5",
            title_color="#18181b",
            text_color="#3f3f46",
            text_light="#a1a1aa",
            accent="#06b6d4",
            accent_secondary="#0891b2",
            overlay="#18181b",
            overlay_opacity=0.65,
        ),
    ),
}


def get_theme(name: str) -> Theme:
    """テーマ名からThemeオブジェクトを取得する。"""
    if name not in THEMES:
        available = ", ".join(THEMES.keys())
        raise ValueError(f"テーマ '{name}' は存在しません。利用可能: {available}")
    return THEMES[name]


def list_themes() -> list[str]:
    """利用可能なテーマ名のリストを返す。"""
    return list(THEMES.keys())
