#pragma once
#include <QDialog>
#include <QStackedWidget>
#include <QLineEdit>
#include <QComboBox>
#include <QPushButton>
#include <QLabel>

class OnboardingWizard : public QDialog {
    Q_OBJECT
public:
    explicit OnboardingWizard(QWidget* parent = nullptr);
    QString toolsDir() const { return m_toolsDir; }
    QString selectedTheme() const { return m_selectedTheme; }
private slots:
    void browseFolder();
    void useDefaultPath();
    void previewTheme(const QString& labelText);
    void goBack();
    void goNext();
private:
    void updateNavButtons();
    
    QStackedWidget* m_stack;
    QLineEdit* m_pathEntry;
    QComboBox* m_themeCombo;
    QPushButton* m_backBtn;
    QPushButton* m_nextBtn;
    
    QString m_toolsDir;
    QString m_selectedTheme;
};
