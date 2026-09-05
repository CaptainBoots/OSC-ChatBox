#pragma once
#include <QColor>
#include <QString>
#include <QMap>
#include <QStringList>
#include <QFont>

struct ThemePalette {
    QString bg;
    QString panel;
    QString border;
    QString accent;
    QString accent2;
    QString tab;
    QString text;
    QString text2;
    QString subtext;
    QString green;
    QString red;
    QString yellow;
    QString cyan;
    QString orange;
    QStringList stripe_colours;
};

class ThemeManager {
public:
    static ThemeManager& instance();
    void setTheme(const QString& mode);
    QString currentMode() const { return m_currentMode; }
    ThemePalette currentPalette() const { return m_palettes[m_currentMode]; }
    
    QMap<QString, ThemePalette> palettes() const { return m_palettes; }
    QMap<QString, QString> labels() const { return m_labels; }
    
    QFont qtFont(int size, bool bold = false);
    QString accentButtonQss() const;
    QString subtleButtonQss() const;
    QString lineEditQss() const;
    QString qss() const;

private:
    ThemeManager();
    QString m_currentMode;
    QMap<QString, ThemePalette> m_palettes;
    QMap<QString, QString> m_labels;
};
