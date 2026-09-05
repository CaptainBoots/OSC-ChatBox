#include "theme.h"
#include <QFontDatabase>

ThemeManager& ThemeManager::instance() {
    static ThemeManager inst;
    return inst;
}

ThemeManager::ThemeManager() : m_currentMode("rich_purple") {
    // Labels
    m_labels = {
        {"dark", "Dark"}, {"rich_purple", "Rich Purple"}, {"dark_sand", "Dark Sand"},
        {"absolute_zero", "Absolute Zero"}, {"light_purple", "Light Purple"}, {"light_sand", "Light Sand"},
        {"mint", "Mint"}, {"dark_mint", "Dark Mint"}, {"dark_red", "Dark Red"}, {"light_red", "Light Red"},
        {"light_blue", "Light Blue"}, {"dark_rainbow", "Dark Rainbow"}, {"light_rainbow", "Light Rainbow"},
        {"pride_flag", "Pride Flag"}, {"trans_flag", "Trans Flag"}, {"nonbinary_flag", "Nonbinary Flag"},
        {"ace_flag", "Ace Flag"}, {"bi_flag", "Bi Flag"}, {"gay_flag", "Gay Flag"}, {"lesbian_flag", "Lesbian Flag"},
        {"pan_flag", "Pan Flag"}, {"genderqueer_flag", "Genderqueer Flag"}, {"aro_flag", "Aro Flag"},
        {"genderfluid_flag", "Genderfluid Flag"}, {"intersex_flag", "Intersex Flag"}, {"demi_flag", "Demi Flag"}
    };

    // Palettes
    m_palettes["dark"] = {
        "#0f0f13", "#17171f", "#2a2a38", "#7c5cfc", "#a78bfa", "#4ade80",
        "#e2e0f0", "#E0E0E0", "#7e7b9a", "#4ade80", "#f87171", "#facc15", "#67e8f9", "#fb923c",
        {}
    };

    m_palettes["rich_purple"] = {
        "#0f0f13", "#1f102a", "#2a2a38", "#9D00FF", "#b44bff", "#4ade80",
        "#e2e0f0", "#E0E0E0", "#7e7b9a", "#00ffcc", "#ff4b72", "#facc15", "#67e8f9", "#fb923c",
        {}
    };

    m_palettes["dark_sand"] = {
        "#1C1D26", "#1f232d", "#353333", "#FFAC8B", "#FFC695", "#FFC695",
        "#F5EFE9", "#E3D8D0", "#AE9281", "#4ade80", "#f87171", "#facc15", "#67e8f9", "#FFAC8B",
        {}
    };

    m_palettes["absolute_zero"] = {
        "#000D21", "#002154", "#003487", "#005CED", "#5496FF", "#2177FF",
        "#EAF3FF", "#D6E8FF", "#A8C4F2", "#4ade80", "#f87171", "#facc15", "#67e8f9", "#fb923c",
        {}
    };

    m_palettes["light_purple"] = {
        "#F6E6FA", "#ffffff", "#DDCAE3", "#9D00FF", "#b44bff", "#000000",
        "#1a1829", "#1a1829", "#1a1829", "#4ade80", "#f87171", "#facc15", "#67e8f9", "#fb923c",
        {}
    };

    m_palettes["light_sand"] = {
        "#fdfbf7", "#f4f1ea", "#e4dfd3", "#2b5c43", "#3d7a5a", "#000000",
        "#1c1b18", "#383630", "#706e64", "#15803d", "#b91c1c", "#b45309", "#0369a1", "#c2410c",
        {}
    };

    m_palettes["mint"] = {
        "#F5FFFA", "#FFFFFF", "#D6F0E4", "#2EC4B6", "#6EE7D8", "#1F2937",
        "#1A2A2A", "#334155", "#64748B", "#22C55E", "#EF4444", "#EAB308", "#06B6D4", "#F97316",
        {}
    };

    m_palettes["dark_mint"] = {
        "#0F1C18", "#163129", "#295247", "#2EC4B6", "#6EE7D8", "#6EE7D8",
        "#E8FFF9", "#D3F5EE", "#8AB5AB", "#4ADE80", "#F87171", "#FACC15", "#67E8F9", "#FB923C",
        {}
    };

    m_palettes["dark_red"] = {
        "#1A0B0B", "#2C1111", "#512121", "#DC2626", "#F87171", "#F87171",
        "#FFF1F1", "#F8DADA", "#B48D8D", "#4ADE80", "#F87171", "#FACC15", "#67E8F9", "#FB923C",
        {}
    };

    m_palettes["light_red"] = {
        "#FFF5F5", "#FFFFFF", "#F4CACA", "#DC2626", "#F87171", "#000000",
        "#2A1111", "#472020", "#735353", "#16A34A", "#DC2626", "#CA8A04", "#0284C7", "#EA580C",
        {}
    };

    m_palettes["light_blue"] = {
        "#F2F9FF", "#FFFFFF", "#D2E8F8", "#3B82F6", "#60A5FA", "#000000",
        "#172033", "#2E4468", "#6A82A8", "#22C55E", "#EF4444", "#EAB308", "#06B6D4", "#F97316",
        {}
    };

    m_palettes["dark_rainbow"] = {
        "#1A1A1A", "#252525", "#444444", "#E40303", "#FF8C00", "#732982",
        "#FFFFFF", "#F0F0F0", "#BBBBBB", "#008026", "#E40303", "#FFED00", "#004DFF", "#FF8C00",
        {"#FF0000", "#FF4400", "#FF8900", "#FFCE00", "#F9FF00", "#ADFF00",
         "#60FF00", "#14FF00", "#00FF38", "#00FF84", "#00FFD1", "#00E8FF",
         "#00AAFF", "#0056FF", "#0002FF", "#4900FF", "#9600FF", "#E200FF",
         "#FF00DD", "#FF0089", "#FF0035"}
    };

    m_palettes["light_rainbow"] = {
        "#FFF5F5", "#FFFFFF", "#F4CACA", "#E40303", "#FF8C00", "#732982",
        "#2A1111", "#472020", "#757575", "#008026", "#E40303", "#FFED00", "#004DFF", "#FF8C00",
        {"#FF0000", "#FF4400", "#FF8900", "#FFCE00", "#F9FF00", "#ADFF00",
         "#60FF00", "#14FF00", "#00FF38", "#00FF84", "#00FFD1", "#00E8FF",
         "#00AAFF", "#0056FF", "#0002FF", "#4900FF", "#9600FF", "#E200FF",
         "#FF00DD", "#FF0089", "#FF0035"}
    };

    m_palettes["pride_flag"] = {
        "#1A1A1A", "#1c1c1c", "#333333", "#FFED00", "#FF8C00", "#FFED00",
        "#FFFFFF", "#F0F0F0", "#CCCCCC", "#008026", "#E40303", "#FFED00", "#004DFF", "#FF8C00",
        {"#E40303", "#FF8C00", "#FFED00", "#008026", "#004DFF", "#750787"}
    };

    m_palettes["trans_flag"] = {
        "#0d1f28", "#1a2e36", "#5BCEFA", "#F5A9B8", "#5BCEFA", "#F5A9B8",
        "#FFFFFF", "#e0f4ff", "#a8d4e8", "#5BCEFA", "#F5A9B8", "#FFFFFF", "#5BCEFA", "#F5A9B8",
        {"#5BCEFA", "#F5A9B8", "#FFFFFF", "#F5A9B8", "#5BCEFA"}
    };

    m_palettes["nonbinary_flag"] = {
        "#1e1230", "#1e1230", "#9C59D1", "#FFF430", "#9C59D1", "#FFF430",
        "#FFFFFF", "#F0F0F0", "#DDDDDD", "#9C59D1", "#FFF430", "#FFF430", "#FFFFFF", "#FFF430",
        {"#FFF430", "#FFFFFF", "#9C59D1", "#2C2C2C"}
    };

    m_palettes["ace_flag"] = {
        "#161616", "#2a002a", "#800080", "#B05ACD", "#CC88EE", "#B05ACD",
        "#FFFFFF", "#F2F2F2", "#CFCFCF", "#B05ACD", "#f87171", "#FFFFFF", "#B05ACD", "#B05ACD",
        {"#161616", "#808080", "#FFFFFF", "#800080"}
    };

    m_palettes["bi_flag"] = {
        "#1a0d1a", "#2b1028", "#9B4F96", "#D60270", "#9B4F96", "#D60270",
        "#FFFFFF", "#F5E6F5", "#C8A0C8", "#9B4F96", "#D60270", "#FFFFFF", "#0038A8", "#D60270",
        {"#D60270", "#D60270", "#9B4F96", "#0038A8", "#0038A8"}
    };

    m_palettes["gay_flag"] = {
        "#00150f", "#002018", "#3D9970", "#3D9970", "#70C9A0", "#3D9970",
        "#FFFFFF", "#E0FFF5", "#7ABBA0", "#3D9970", "#006B54", "#FFFFFF", "#7BADE2", "#3D9970",
        {"#078D70", "#26CEA8", "#98E8C1", "#FFFFFF", "#7BADE2", "#5049CC", "#3D1A8E"}
    };

    m_palettes["lesbian_flag"] = {
        "#1f0d00", "#2e1500", "#D52D00", "#FF9A56", "#FF6D4A", "#FF9A56",
        "#FFFFFF", "#FFE8DC", "#D4907A", "#FF9A56", "#D52D00", "#FF9A56", "#A50062", "#FF9A56",
        {"#D52D00", "#FF9A56", "#FFFFFF", "#D362A4", "#A50062"}
    };

    m_palettes["pan_flag"] = {
        "#0f0f1a", "#1a1a2e", "#FFD800", "#FF218C", "#FFD800", "#FF218C",
        "#FFFFFF", "#F5F5FF", "#BBBBDD", "#21B1FF", "#FF218C", "#FFD800", "#21B1FF", "#FF218C",
        {"#FF218C", "#FF218C", "#FFD800", "#FFD800", "#21B1FF", "#21B1FF"}
    };

    m_palettes["genderqueer_flag"] = {
        "#141020", "#1e1630", "#B57EDC", "#B57EDC", "#CCAAEE", "#B57EDC",
        "#FFFFFF", "#F0EAFF", "#BBA8D8", "#498019", "#B57EDC", "#FFFFFF", "#498019", "#B57EDC",
        {"#B57EDC", "#B57EDC", "#FFFFFF", "#FFFFFF", "#498019", "#498019"}
    };

    m_palettes["aro_flag"] = {
        "#0a120a", "#101e10", "#3DA542", "#3DA542", "#A8D379", "#3DA542",
        "#FFFFFF", "#E8F5E8", "#8CB88C", "#3DA542", "#A8D379", "#FFFFFF", "#3DA542", "#A8D379",
        {"#3DA542", "#A8D379", "#FFFFFF", "#A9A9A9", "#000000"}
    };

    m_palettes["genderfluid_flag"] = {
        "#0d0014", "#170020", "#BE18D6", "#FF76A4", "#BE18D6", "#FF76A4",
        "#FFFFFF", "#F8E8FF", "#C099CC", "#BE18D6", "#FF76A4", "#FFFFFF", "#3300BE", "#FF76A4",
        {"#FF76A4", "#FFFFFF", "#BE18D6", "#000000", "#3300BE"}
    };

    m_palettes["intersex_flag"] = {
        "#1a1400", "#2b2200", "#FFD800", "#FFD800", "#FFE84D", "#FFD800",
        "#FFFFFF", "#FFF8CC", "#CCAA00", "#FFD800", "#7A00C8", "#FFD800", "#7A00C8", "#FFD800",
        {"#FFD800", "#FFD800", "#FFD800", "#7A00C8", "#7A00C8", "#FFD800", "#FFD800", "#FFD800"}
    };

    m_palettes["demi_flag"] = {
        "#121212", "#1e1e1e", "#7A7A7A", "#9966CC", "#BB99EE", "#9966CC",
        "#FFFFFF", "#F0F0F0", "#AAAAAA", "#9966CC", "#7A7A7A", "#FFFFFF", "#9966CC", "#9966CC",
        {"#000000", "#7A7A7A", "#FFFFFF", "#9966CC", "#FFFFFF", "#7A7A7A"}
    };
}

void ThemeManager::setTheme(const QString& mode) {
    if (m_palettes.contains(mode)) {
        m_currentMode = mode;
    } else {
        m_currentMode = "rich_purple";
    }
}

QFont ThemeManager::qtFont(int size, bool bold) {
    QFont font("Consolas", size);
    font.setBold(bold);
    return font;
}

QString ThemeManager::accentButtonQss() const {
    ThemePalette p = currentPalette();
    return QString(
        "QPushButton { background-color: %1; color: %2; "
        "border: none; border-radius: 3px; padding: 6px 14px; font-weight: bold; }"
        "QPushButton:hover { background-color: %3; }"
        "QPushButton:disabled { background-color: %4; color: %5; }"
    ).arg(p.accent, p.bg, p.accent2, p.border, p.subtext);
}

QString ThemeManager::subtleButtonQss() const {
    ThemePalette p = currentPalette();
    return QString(
        "QPushButton { background-color: %1; color: %2; "
        "border: none; border-radius: 3px; padding: 6px 14px; }"
        "QPushButton:hover { background-color: %3; color: %4; }"
    ).arg(p.panel, p.subtext, p.border, p.text);
}

QString ThemeManager::lineEditQss() const {
    ThemePalette p = currentPalette();
    return QString(
        "QLineEdit { background-color: %1; color: %2; "
        "border: 1px solid %3; border-radius: 2px; padding: 3px 6px; "
        "selection-background-color: %4; }"
        "QLineEdit:focus { border: 1px solid %4; }"
    ).arg(p.panel, p.text, p.border, p.accent);
}

QString ThemeManager::qss() const {
    ThemePalette p = currentPalette();
    return QString(
        "QWidget {"
        "    background-color: %1;"
        "    color: %2;"
        "    font-family: 'Consolas';"
        "    border: none;"
        "}"
        "QMainWindow, QDialog { background-color: %1; }"
        "QPushButton {"
        "    background-color: %3; color: %4; border: none;"
        "    border-radius: 3px; padding: 6px 14px; font-weight: bold;"
        "}"
        "QPushButton:hover { background-color: %5; color: %2; }"
        "QPushButton:disabled { color: %6; }"
        "QLineEdit {"
        "    background-color: %3; color: %2; border: 1px solid %5;"
        "    border-radius: 2px; padding: 3px 6px; selection-background-color: %4;"
        "}"
        "QLineEdit:focus { border: 1px solid %4; }"
        "QComboBox {"
        "    background-color: %3; color: %2; border: 1px solid %5;"
        "    border-radius: 2px; padding: 3px 6px;"
        "}"
        "QComboBox QAbstractItemView {"
        "    background-color: %3; color: %2; selection-background-color: %4;"
        "    border: 1px solid %5;"
        "}"
        "QScrollBar:vertical { background: %1; width: 12px; margin: 0; }"
        "QScrollBar::handle:vertical { background: %5; min-height: 24px; border-radius: 4px; }"
        "QScrollBar::handle:vertical:hover { background: %7; }"
        "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }"
        "QToolTip { background-color: %3; color: %2; border: 1px solid %5; }"
    ).arg(p.bg, p.text, p.panel, p.accent, p.border, p.subtext, p.accent2);
}
