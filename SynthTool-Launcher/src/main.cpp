#include <QApplication>
#include <QIcon>
#include <QDir>
#include <QFile>
#include "mainwindow.h"
#include "customwidgets.h"

#ifdef Q_OS_WIN
#include <windows.h>
#include <shellapi.h>
#endif

// Custom Qt message handler to redirect all logs to our ConsoleWindow
void customMessageHandler(QtMsgType type, const QMessageLogContext& context, const QString& msg) {
    QString txt;
    switch (type) {
    case QtDebugMsg:
        txt = QString("[Debug] %1").arg(msg);
        break;
    case QtInfoMsg:
        txt = QString("[Info] %1").arg(msg);
        break;
    case QtWarningMsg:
        txt = QString("[Warning] %1 (Line %2 in %3)").arg(msg).arg(context.line).arg(context.file);
        break;
    case QtCriticalMsg:
        txt = QString("[Critical] %1 (Line %2 in %3)").arg(msg).arg(context.line).arg(context.file);
        break;
    case QtFatalMsg:
        txt = QString("[Fatal] %1 (Line %2 in %3)").arg(msg).arg(context.line).arg(context.file);
        break;
    }
    
    QString formatted = QString("%1\n").arg(txt);
    
    // Append to static console window logs buffer
    ConsoleWindow::appendLog(formatted);
    
    // Also output to original standard error for debugging
    fprintf(stderr, "%s", formatted.toUtf8().constData());
    fflush(stderr);
    
    if (type == QtFatalMsg) {
        abort();
    }
}

int main(int argc, char* argv[]) {
    // Enable High DPI scaling
    QGuiApplication::setHighDpiScaleFactorRoundingPolicy(Qt::HighDpiScaleFactorRoundingPolicy::PassThrough);

    QApplication app(argc, argv);
    app.setApplicationName("SynthTool-Launcher");
    app.setApplicationVersion(ConfigManager::instance().version());

    // Register windows process ID for taskbar grouping
#ifdef Q_OS_WIN
    typedef HRESULT(WINAPI* SetCurrentProcessExplicitAppUserModelIDFunc)(PCWSTR);
    HMODULE shell32 = LoadLibraryW(L"shell32.dll");
    if (shell32) {
        auto setAppId = (SetCurrentProcessExplicitAppUserModelIDFunc)GetProcAddress(shell32, "SetCurrentProcessExplicitAppUserModelID");
        if (setAppId) {
            setAppId(L"CaptainBoots.SynthTool-Launcher.1.0.1");
        }
        FreeLibrary(shell32);
    }
#endif

    // Setup custom message handler
    qInstallMessageHandler(customMessageHandler);

    qDebug() << "SynthTool-Launcher starting up...";
    qDebug() << "Version:" << ConfigManager::instance().version();

    // Set Window Icon
    QString iconPath = QDir(QCoreApplication::applicationDirPath()).filePath("Images/Boot's-ToolBox-256.ico");
    if (!QFile::exists(iconPath)) {
        iconPath = QDir(QCoreApplication::applicationDirPath() + "/../Images/Boot's-ToolBox-256.ico").canonicalPath();
    }
    if (QFile::exists(iconPath)) {
        app.setWindowIcon(QIcon(iconPath));
        qDebug() << "[Process] Loaded application icon from:" << iconPath;
    }

    MainWindow mainWin;
    if (QFile::exists(iconPath)) {
        mainWin.setWindowIcon(QIcon(iconPath));
    }
    mainWin.show();

    return app.exec();
}
