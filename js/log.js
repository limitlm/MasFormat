// 日志模块
const LogModule = {
    /**
     * 从 PluginStorage 加载日志
     */
    loadLogs() {
        try {
            if (window.Application && window.Application.PluginStorage) {
                const logsJson = window.Application.PluginStorage.getItem('LogModuleLogs');
                return logsJson ? JSON.parse(logsJson) : [];
            }
        } catch (error) {
            console.error('加载日志失败:', error);
        }
        return [];
    },
    
    /**
     * 保存日志到 PluginStorage
     * @param {Array} logs 日志数组
     */
    saveLogs(logs) {
        try {
            if (window.Application && window.Application.PluginStorage) {
                window.Application.PluginStorage.setItem('LogModuleLogs', JSON.stringify(logs));
            }
        } catch (error) {
            console.error('保存日志失败:', error);
        }
    },
    
    /**
     * 记录日志
     * @param {string} message 日志消息
     * @param {string} level 日志级别
     */
    addLog(message, level = 'info') {
        const now = new Date();
        const timeString = now.toLocaleString('zh-CN', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });
        
        const logEntry = {
            time: timeString,
            level: level.toUpperCase(),
            message: message
        };
        
        // 加载现有日志
        const logs = this.loadLogs();
        logs.push(logEntry);
        
        // 保存日志
        this.saveLogs(logs);
        
        // 通知所有监听器
        this.notifyListeners(logEntry);
        
        return logEntry;
    },
    
    /**
     * 清空日志
     */
    clearLogs() {
        // 清空 PluginStorage 中的日志
        this.saveLogs([]);
        
        // 通知所有监听器
        this.notifyListeners({ type: 'clear' });
    },
    
    /**
     * 获取所有日志
     */
    getLogs() {
        return this.loadLogs();
    },
    
    // 监听器管理
    listeners: [],
    
    /**
     * 添加监听器
     */
    addListener(listener) {
        this.listeners.push(listener);
    },
    
    /**
     * 通知所有监听器
     */
    notifyListeners(data) {
        this.listeners.forEach(listener => {
            if (typeof listener === 'function') {
                listener(data);
            }
        });
    }
};

// 导出模块
window.LogModule = LogModule;