#include "configmanager.h"
#include "theme.h"
#include "onboardingwizard.h"
#include <QStandardPaths>
#include <QDir>
#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QTextStream>
#include <QCoreApplication>
#include <QProcess>
#include <QRegularExpression>
#include <QDebug>

ConfigManager& ConfigManager::instance() {
    static ConfigManager inst;
    return inst;
}

ConfigManager::ConfigManager() : m_updateBranch("main"), m_betaPopupShown(false) {
    // Default tools directory
    QString appData = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    if (!appData.endsWith("SynthTool-Launcher", Qt::CaseInsensitive)) {
        appData = QDir(QStandardPaths::writableLocation(QStandardPaths::GenericDataLocation)).filePath("SynthTool-Launcher");
    }
    m_toolsRootDir = QDir::toNativeSeparators(appData);
}

QString ConfigManager::centralConfigDir() const {
    QString path = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    if (!path.endsWith("SynthTool-Launcher", Qt::CaseInsensitive)) {
        path = QDir(QStandardPaths::writableLocation(QStandardPaths::GenericDataLocation)).filePath("SynthTool-Launcher");
    }
    QDir().mkpath(path);
    return QDir::toNativeSeparators(path);
}

QString ConfigManager::toolsPathPointerFile() const {
    return QDir(centralConfigDir()).filePath("tools_path.txt");
}

QString ConfigManager::installPathPointerFile() const {
    return QDir(centralConfigDir()).filePath("install_path.txt");
}

QString ConfigManager::toolboxConfigDir() const {
    return QDir(m_toolsRootDir).filePath("configs");
}

QString ConfigManager::toolboxConfigFile() const {
    return QDir(toolboxConfigDir()).filePath("toolbox_config.json");
}

void ConfigManager::setToolsRootDir(const QString& dir) {
    m_toolsRootDir = QDir::toNativeSeparators(dir);
    QDir().mkpath(toolboxConfigDir());
}

QString ConfigManager::activePython() const {
    if (!m_pythonInterpreter.isEmpty() && QFile::exists(m_pythonInterpreter)) {
        return m_pythonInterpreter;
    }

    // Search system path on Windows
#ifdef Q_OS_WIN
    QStringList searchCmds = {"pythonw.exe", "python.exe", "python3.exe"};
    for (const auto& cmd : searchCmds) {
        QString fullPath = QStandardPaths::findExecutable(cmd);
        if (!fullPath.isEmpty()) {
            return fullPath;
        }
    }
    return "pythonw.exe"; // Fallback to let OS resolve it
#else
    QString path = QStandardPaths::findExecutable("python3");
    if (path.isEmpty()) {
        path = QStandardPaths::findExecutable("python");
    }
    return path.isEmpty() ? "python3" : path;
#endif
}

void ConfigManager::load() {
    // Save install path pointer
    QString appDir = QCoreApplication::applicationDirPath();
    QFile instFile(installPathPointerFile());
    if (instFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QTextStream out(&instFile);
        out << QDir::toNativeSeparators(appDir);
        instFile.close();
    }

    bool configLoaded = false;
    QString toolsRoot = "";

    // 1. Check tools path pointer
    QFile ptrFile(toolsPathPointerFile());
    if (ptrFile.exists() && ptrFile.open(QIODevice::ReadOnly | QIODevice::Text)) {
        QTextStream in(&ptrFile);
        toolsRoot = in.readAll().trimmed();
        ptrFile.close();
    }

    QJsonObject config;
    if (!toolsRoot.isEmpty() && QDir(toolsRoot).exists()) {
        setToolsRootDir(toolsRoot);
        QFile cfgFile(toolboxConfigFile());
        if (cfgFile.exists() && cfgFile.open(QIODevice::ReadOnly)) {
            QJsonDocument doc = QJsonDocument::fromJson(cfgFile.readAll());
            if (doc.isObject()) {
                config = doc.object();
                configLoaded = true;
            }
            cfgFile.close();
        }
    }

    // 2. Fallback to default location
    if (!configLoaded) {
        QString fallbackConfigFile = QDir(centralConfigDir()).filePath("toolbox_config.json");
        QFile fFile(fallbackConfigFile);
        if (fFile.exists() && fFile.open(QIODevice::ReadOnly)) {
            QJsonDocument doc = QJsonDocument::fromJson(fFile.readAll());
            if (doc.isObject()) {
                config = doc.object();
                configLoaded = true;
                toolsRoot = config["tools_root_dir"].toString();
                if (!toolsRoot.isEmpty()) {
                    setToolsRootDir(toolsRoot);
                }
            }
            fFile.close();
        }
    }

    // 3. First-time run! Run wizard
    if (!configLoaded) {
        OnboardingWizard wizard;
        if (wizard.exec() == QDialog::Accepted) {
            toolsRoot = wizard.toolsDir();
            QString theme = wizard.selectedTheme();
            setToolsRootDir(toolsRoot);
            ThemeManager::instance().setTheme(theme);
        } else {
            toolsRoot = centralConfigDir();
            setToolsRootDir(toolsRoot);
            ThemeManager::instance().setTheme("rich_purple");
        }

        m_updateBranch = "main";
        m_betaPopupShown = false;
        m_pythonInterpreter = "";
        m_managedScripts = discoverManagedScripts();
        
        save();
        configLoaded = true;
    }

    // Save tools path pointer
    if (ptrFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QTextStream out(&ptrFile);
        out << QDir::toNativeSeparators(m_toolsRootDir);
        ptrFile.close();
    }

    if (configLoaded && !config.isEmpty()) {
        m_updateBranch = config.value("update_branch").toString("main");
        m_betaPopupShown = config.value("beta_popup_shown").toBool(false);
        m_pythonInterpreter = config.value("python_interpreter").toString();
        ThemeManager::instance().setTheme(config.value("theme_mode").toString("rich_purple"));
        
        // Load cached labels
        m_cachedLabels.clear();
        QJsonObject cachedObj = config.value("cached_labels").toObject();
        for (auto it = cachedObj.begin(); it != cachedObj.end(); ++it) {
            m_cachedLabels[it.key()] = it.value().toString();
        }

        // Load managed scripts
        m_managedScripts.clear();
        QJsonArray scriptsArr = config.value("managed_scripts").toArray();
        for (const auto& val : scriptsArr) {
            QJsonObject sObj = val.toObject();
            ManagedScript s;
            s.filename = sObj["filename"].toString();
            s.label = sObj["label"].toString();
            s.custom = sObj["custom"].toBool(false);
            m_managedScripts.append(s);
        }

        // If scripts list is empty or doesn't have local scans, merge with discovered scripts
        if (m_managedScripts.isEmpty()) {
            m_managedScripts = discoverManagedScripts();
            save();
        }
    }
}

void ConfigManager::save() {
    QJsonObject config;
    config["version"] = version();
    config["update_branch"] = m_updateBranch;
    config["beta_popup_shown"] = m_betaPopupShown;
    config["python_interpreter"] = m_pythonInterpreter;
    config["theme_mode"] = ThemeManager::instance().currentMode();
    config["tools_root_dir"] = m_toolsRootDir;

    // Save cached labels
    QJsonObject cachedObj;
    for (auto it = m_cachedLabels.begin(); it != m_cachedLabels.end(); ++it) {
        cachedObj[it.key()] = it.value();
    }
    config["cached_labels"] = cachedObj;

    // Save managed scripts
    QJsonArray scriptsArr;
    for (const auto& s : m_managedScripts) {
        QJsonObject sObj;
        sObj["filename"] = s.filename;
        sObj["label"] = s.label;
        sObj["custom"] = s.custom;
        scriptsArr.append(sObj);
    }
    config["managed_scripts"] = scriptsArr;

    QDir().mkpath(toolboxConfigDir());
    QFile cfgFile(toolboxConfigFile());
    if (cfgFile.open(QIODevice::WriteOnly)) {
        QJsonDocument doc(config);
        cfgFile.write(doc.toJson(QJsonDocument::Indented));
        cfgFile.close();
    }
}

QVector<ManagedScript> ConfigManager::discoverManagedScripts() {
    QVector<ManagedScript> detected;

    // 1. Always include LibreHardwareMonitor as a static default helper tool
    ManagedScript lhm;
    lhm.filename = "LibreHardwareMonitor/LibreHardwareMonitor.exe";
    lhm.label = "Libre Hardware Monitor";
    lhm.custom = false;
    detected.append(lhm);

    // Keep track of folders discovered
    QStringList seenFolders;

    // 2. Scan local tools directory
    QDir dir(m_toolsRootDir);
    if (dir.exists()) {
        QStringList entries = dir.entryList(QDir::Dirs | QDir::NoDotAndDotDot);
        for (const QString& item : entries) {
            if (item == "LibreHardwareMonitor" || item == "configs" || item == "ToolBox Backup") {
                continue;
            }

            QDir subDir(dir.filePath(item));
            QFile mainPy(subDir.filePath("main.py"));
            if (mainPy.exists() && mainPy.open(QIODevice::ReadOnly | QIODevice::Text)) {
                QTextStream in(&mainPy);
                QString content = in.readAll();
                mainPy.close();

                // Extract NAME = "..."
                QString label = item;
                QRegularExpression re("NAME\\s*=\\s*['\"]([^'\"]+)['\"]");
                QRegularExpressionMatch match = re.match(content);
                if (match.hasMatch()) {
                    label = match.captured(1);
                }

                ManagedScript s;
                s.filename = item + "/main.py";
                s.label = label;
                s.custom = false;
                detected.append(s);
                seenFolders.append(item);
            }
        }
    }

    // Incorporate custom scripts already present in current config if any
    for (const auto& s : m_managedScripts) {
        if (s.custom) {
            detected.append(s);
        }
    }

    return detected;
}
