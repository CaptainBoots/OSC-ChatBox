#pragma once
#include <QDialog>
#include <QLabel>
#include <QPushButton>
#include <QVector>

struct HelpPage {
    QString title;
    QString content;
};

class HelpDialog : public QDialog {
    Q_OBJECT
public:
    explicit HelpDialog(QWidget* parent = nullptr);
private slots:
    void goBack();
    void nextOrFinish();
private:
    void showPage(int idx);
    
    QLabel* m_titleLabel;
    QLabel* m_contentLabel;
    QLabel* m_pageIndicator;
    QPushButton* m_prevBtn;
    QPushButton* m_nextBtn;
    
    int m_currentPage;
    QVector<HelpPage> m_pages;
};
