#pragma once
#include <QString>
#include <QVector>
#include <QMap>

struct ManagedScript {
    QString filename;
    QString label;
    bool custom = false;
};

class ConfigManager {
public:
    static ConfigManager& instance();
    
    void load();
    void save();
    
    QString version() const { return APP_VERSION; }
    
    QString updateBranch() const { return m_updateBranch; }
    void setUpdateBranch(const QString& branch) { m_updateBranch = branch; }
    
    bool betaPopupShown() const { return m_betaPopupShown; }
    void setBetaPopupShown(bool shown) { m_betaPopupShown = shown; }
    
    QString pythonInterpreter() const { return m_pythonInterpreter; }
    void setPythonInterpreter(const QString& path) { m_pythonInterpreter = path; }
    
    QString toolsRootDir() const { return m_toolsRootDir; }
    void setToolsRootDir(const QString& dir);
    
    QVector<ManagedScript> managedScripts() const { return m_managedScripts; }
    void setManagedScripts(const QVector<ManagedScript>& scripts) { m_managedScripts = scripts; }
    void addManagedScript(const ManagedScript& s) { m_managedScripts.append(s); }
    void removeManagedScript(int idx) { if (idx >= 0 && idx < m_managedScripts.size()) m_managedScripts.removeAt(idx); }
    
    QMap<QString, QString> cachedLabels() const { return m_cachedLabels; }
    void setCachedLabel(const QString& filename, const QString& label) { m_cachedLabels[filename] = label; }
    
    QString activePython() const;
    QString centralConfigDir() const;
    QString toolsPathPointerFile() const;
    QString installPathPointerFile() const;
    QString toolboxConfigFile() const;
    QString toolboxConfigDir() const;
    
    QVector<ManagedScript> discoverManagedScripts();
    
private:
    ConfigManager();
    
    QString m_updateBranch;
    bool m_betaPopupShown;
    QString m_pythonInterpreter;
    QString m_toolsRootDir;
    QVector<ManagedScript> m_managedScripts;
    QMap<QString, QString> m_cachedLabels;
};
