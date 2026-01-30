// 日志模块
const LogModule = {
    // 广播通道实例
    broadcastChannel: null,
    
    /**
     * 初始化广播通道
     */
    initBroadcastChannel() {
        try {
            if (window.BroadcastChannel) {
                this.broadcastChannel = new BroadcastChannel('mas-format-logs-channel');
                console.log('BroadcastChannel初始化成功');
            } else {
                console.warn('当前环境不支持BroadcastChannel');
            }
        } catch (error) {
            console.error('初始化BroadcastChannel失败:', error);
        }
    },
    
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
            message: message,
            timestamp: now.getTime()
        };
        
        // 加载现有日志
        const logs = this.loadLogs();
        logs.push(logEntry);
        
        // 保存日志
        this.saveLogs(logs);
        
        // 通知所有监听器
        this.notifyListeners(logEntry);
        
        // 广播日志到其他页面
        this.broadcastLog(logEntry);
        
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
        
        // 广播清空命令到其他页面
        this.broadcastLog({ type: 'clear' });
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
                try {
                    listener(data);
                } catch (error) {
                    console.error('通知监听器失败:', error);
                }
            }
        });
    },
    
    /**
     * 广播日志到其他页面
     * @param {Object} data 日志数据
     */
    broadcastLog(data) {
        try {
            if (this.broadcastChannel) {
                // 验证数据格式
                const validData = this.validateLogData(data);
                this.broadcastChannel.postMessage(validData);
            }
        } catch (error) {
            console.error('广播日志失败:', error);
        }
    },
    
    /**
     * 验证日志数据格式
     * @param {Object} data 日志数据
     * @returns {Object} 验证后的日志数据
     */
    validateLogData(data) {
        if (typeof data !== 'object' || data === null) {
            return { type: 'error', message: 'Invalid log data' };
        }
        
        if (data.type === 'clear') {
            return { type: 'clear', timestamp: Date.now() };
        }
        
        return {
            time: data.time || new Date().toLocaleString('zh-CN'),
            level: data.level || 'INFO',
            message: data.message || '',
            timestamp: data.timestamp || Date.now(),
            type: 'log'
        };
    },
    
    /**
     * 关闭广播通道
     */
    closeBroadcastChannel() {
        try {
            if (this.broadcastChannel) {
                this.broadcastChannel.close();
                this.broadcastChannel = null;
                console.log('BroadcastChannel已关闭');
            }
        } catch (error) {
            console.error('关闭BroadcastChannel失败:', error);
        }
    }
};

// 导出模块
window.LogModule = LogModule;

// 初始化广播通道
window.LogModule.initBroadcastChannel();

// 页面卸载时关闭广播通道
window.addEventListener('beforeunload', function() {
    window.LogModule.closeBroadcastChannel();
});