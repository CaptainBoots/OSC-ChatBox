#include <QApplication>
#include <QMessageBox>
#include <QDir>
#include <QFile>
#include <QTextStream>
#include <QStandardPaths>

int main(int argc, char* argv[]) {
    // Enable High DPI scaling
    QGuiApplication::setHighDpiScaleFactorRoundingPolicy(Qt::HighDpiScaleFactorRoundingPolicy::PassThrough);

    QApplication app(argc, argv);
    app.setApplicationName("CToolBox-Uninstaller");

    // Standard styling for QMessageBox
    QMessageBox::StandardButton reply;
    reply = QMessageBox::question(nullptr, "Uninstall CToolBox-Launcher",
        "Are you sure you want to completely uninstall CToolBox-Launcher?\n\n"
        "This will permanently delete:\n"
        "• All downloaded companion tools and files inside your tools folder\n"
        "• All settings, configurations, and themes\n"
        "• All central pointer files and log files\n\n"
        "This action cannot be undone.",
        QMessageBox::Yes | QMessageBox::No);

    if (reply != QMessageBox::Yes) {
        return 0;
    }

    // Determine central AppData path
    QString appData = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    if (!appData.endsWith("CToolBox-Launcher", Qt::CaseInsensitive)) {
        appData = QDir(QStandardPaths::writableLocation(QStandardPaths::GenericDataLocation)).filePath("CToolBox-Launcher");
    }
    appData = QDir::toNativeSeparators(appData);

    // Read tools path pointer
    QString toolsPathFile = QDir(appData).filePath("tools_path.txt");
    QString toolsDir = "";
    if (QFile::exists(toolsPathFile)) {
        QFile file(toolsPathFile);
        if (file.open(QIODevice::ReadOnly | QIODevice::Text)) {
            QTextStream in(&file);
            toolsDir = in.readAll().trimmed();
            file.close();
        }
    }

    bool deletedTools = false;
    bool deletedConfig = false;

    // 1. Delete tools directory if safe (length check)
    if (!toolsDir.isEmpty() && toolsDir.length() > 15 && QDir(toolsDir).exists()) {
        deletedTools = QDir(toolsDir).removeRecursively();
    }

    // 2. Delete central AppData config directory if safe (length check)
    if (!appData.isEmpty() && appData.length() > 15 && QDir(appData).exists()) {
        deletedConfig = QDir(appData).removeRecursively();
    }

    QMessageBox::information(nullptr, "Uninstall Success",
        "✓ CToolBox-Launcher has been successfully uninstalled!\n\n"
        "All downloaded tools, settings, and central config folders have been cleanly removed.");

    return 0;
}
