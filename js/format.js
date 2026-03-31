// 格式化相关功能实现

/**
 * 页面格式化功能
 * @description 设置文档为A4规格、2.5厘米页边距，配置页眉页脚和页码
 * @returns {boolean} 操作是否成功
 */
function pageFormat() {
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    window.LogModule.addLog("开始执行页面格式化", "warning");
    
    // 常量定义
    const CM_TO_POINT = 28.3465;
    const PAPER_A4 = 7;
    const ALIGN_CENTER = 1;
    const ALIGN_RIGHT = 2;
    const PAGE_NUMBER_ARABIC = 0;

    // 步骤1：保存所有节的原有页面方向
    const originalOrientations = [];
    for (let i = 1; i <= doc.Sections.Count; i++) {
      const ori = doc.Sections.Item(i).PageSetup.Orientation;
      originalOrientations.push(ori);
    }
    window.LogModule.addLog(`已保存 ${doc.Sections.Count} 个节的页面方向`, "info");

    // 步骤2：逐个节设置页面基础属性（避免文档级别干扰）
    for (let i = 1; i <= doc.Sections.Count; i++) {
      const section = doc.Sections.Item(i);
      section.PageSetup.PaperSize = PAPER_A4;
      section.PageSetup.LeftMargin = 2.5 * CM_TO_POINT;
      section.PageSetup.RightMargin = 2.5 * CM_TO_POINT;
      section.PageSetup.TopMargin = 2.5 * CM_TO_POINT;
      section.PageSetup.BottomMargin = 2.5 * CM_TO_POINT;
      section.PageSetup.HeaderDistance = 1.5 * CM_TO_POINT;
      section.PageSetup.FooterDistance = 1.75 * CM_TO_POINT;
    }
    window.LogModule.addLog("页面基础设置完成：A4规格，2.5厘米页边距", "info");

    // 步骤3：恢复所有节的原有页面方向
    for (let i = 1; i <= doc.Sections.Count; i++) {
      doc.Sections.Item(i).PageSetup.Orientation = originalOrientations[i - 1];
    }
    window.LogModule.addLog("页面方向恢复完成", "info");

    // 步骤4：配置第一节页脚（公司名称+页码）
    const section1 = doc.Sections.Item(1);
    const footer = section1.Footers.Item(1);
    
    // 清除原有页脚内容
    footer.Range.Delete();
    
    // 设置首页不同
    section1.PageSetup.DifferentFirstPageHeaderFooter = true;
    section1.PageSetup.PageNumberStyle = PAGE_NUMBER_ARABIC;
    
    // 添加公司名称
    const footerRange = footer.Range;
    footerRange.Text = "重庆梅安森科技股份有限公司 编制";
    footerRange.ParagraphFormat.Alignment = ALIGN_RIGHT;
    
    // 插入换行符和页码
    footerRange.Collapse(1);
    footerRange.Text = "\n";
    footerRange.MoveEnd(1, -1);
    const pageField = footerRange.Fields.Add(footerRange, -1, "PAGE", false);
    pageField.Code.ParagraphFormat.Alignment = ALIGN_CENTER;
    
    window.LogModule.addLog("第一节页脚配置完成", "info");

    // 步骤5：配置其他节（页眉页脚同前节，页码连续）
    if (doc.Sections.Count > 1) {
      for (let i = 2; i <= doc.Sections.Count; i++) {
        try {
          const section = doc.Sections.Item(i);
          section.PageSetup.DifferentFirstPageHeaderFooter = false;
          
          try {
            section.Headers.Item(1).LinkToPrevious = true;
            section.Footers.Item(1).LinkToPrevious = true;
          } catch (e) {
            window.LogModule.addLog(`警告：第${i}节页眉页脚同前节失败 - ${e.description}`, "warning");
          }
          
          section.Footers.Item(1).PageNumbers.RestartNumberingAtSection = false;
        } catch (e) {
          window.LogModule.addLog(`错误：处理第${i}节失败 - ${e.description}`, "error");
        }
      }
      window.LogModule.addLog("其他节配置完成", "info");
    }

    // 步骤6：设置页码起始值
    section1.Footers.Item(1).PageNumbers.RestartNumberingAtSection = true;
    section1.Footers.Item(1).PageNumbers.StartingNumber = 0;
    window.LogModule.addLog("页码起始设置完成", "info");

    // 完成
    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(`页面格式化完成，耗时：${duration}秒`, "success");
    return true;
  } catch (error) {
    window.LogModule.addLog(`页面格式化失败: ${error.message}`, "error");
    return false;
  }
}

/**
 * 标题格式化功能
 * @param {Object} options - 配置选项
 * @param {boolean} options.removeNumbering - 是否移除标题编号
 * @param {boolean} options.refreshStyles - 是否刷新标题样式
 * @param {boolean} options.showConfirm - 是否显示确认弹窗
 * @returns {boolean} 操作是否成功
 */
function titleFormat(options = {}) {
  const {
    removeNumbering = true,
    refreshStyles = true,
    showConfirm = true
  } = options;

  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  let originalSelection;
  try {
    // 显示确认弹窗
    if (showConfirm) {
      const confirmed = window.confirm("请确认已经应用自定义样式集，\n点击确认将自动执行标题格式化");
      if (!confirmed) {
        window.LogModule.addLog("用户取消了标题格式化操作", "info");
        return false;
      }
    }
    
    // 开始记录执行时间
    const startTime = performance.now();
    let processedCount = 0;
    
    // 先保存当前选择范围
    originalSelection = window.Application.Selection.Range;
    window.LogModule.addLog("开始刷新标题格式", "warning");
    
    // 处理标题样式
    if (refreshStyles) {
      for (let i = 1; i <= 9; i++) {
        const styleName = `标题 ${i}`;
        
        // 尝试选择所有当前标题样式实例
        doc.SelectStyleInstance(styleName);
        
        // 检查是否成功选择了当前标题样式实例
        const currentSelection = window.Application.Selection.Range;
        if (currentSelection.Start !== currentSelection.End) {
          // 只有在找到当前标题样式实例时才设置样式
          window.Application.Selection.Style = styleName;
          window.LogModule.addLog(`${styleName} 格式化完成`, "info");
        }
      }
      

    }

    // 清理标题多余字符
    if (removeNumbering) {
      window.LogModule.addLog("开始清理标题多余字符", "warning");
      
      // 遍历所有标题样式（标题1-9）
      for (let i = 1; i <= 9; i++) {
        const styleName = `标题 ${i}`;
        
        // 使用Range.Find来查找所有使用该标题样式的段落
        const findRange = doc.Content;
        const find = findRange.Find;
        
        // 设置查找参数
        find.ClearFormatting();
        find.Style = styleName;
        find.Forward = true;
        find.Wrap = 1; // wdFindStop
        find.Format = true;
        
        // 开始查找
        while (find.Execute()) {
          // 获取找到的段落
          const para = findRange.Paragraphs.Item(1);
          const range = para.Range;
          const text = range.Text;
          
          // 检查是否以数字或空格开头
          if (/^\d|^\s/.test(text)) {
            // 获取段落标记前的内容长度
            const contentLength = range.Text.length - 1;
            
            // 如果有内容需要处理
            if (contentLength > 0) {
              // 设置范围只包含内容，不包含段落标记
              range.SetRange(range.Start, range.Start + contentLength);
              const content = range.Text;
              
              // 处理文本
              const processedContent = content.replace(/^[\d\.、\s]+/, "");
              
              // 检查是否有变化
              if (processedContent !== content) {
                range.Text = processedContent;
                processedCount++;
                // 输出修改日志
                window.LogModule.addLog(`修改异常标题: "${content}"->"${processedContent}"`, "info");
              }
            }
          }
        }
      }
      
      // 显示处理结果
      window.LogModule.addLog(`处理完成！共处理了 ${processedCount} 个标题。`, "info");
    }

    // 恢复原始选择
    if (originalSelection) {
      try {
        originalSelection.Select();
      } catch (e) {
        // 忽略恢复选择时的错误
      }
    }

    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      `标题格式化完成，耗时：${duration}秒`,
      "success"
    );
    return true;
  } catch (error) {
    // 恢复原始选择
    if (originalSelection) {
      try {
        originalSelection.Select();
      } catch (e) {
        // 忽略恢复选择时的错误
      }
    }
    
    window.LogModule.addLog(`标题格式化失败: ${error.message}`, "error");
    return false;
  }
}

/**
 * 表格格式化功能
 * @returns {boolean} 操作是否成功
 */
function tableFormat() {
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    // 实现表格格式化逻辑
    // 这里可以添加具体的表格格式化代码
    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "表格格式化完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("表格格式化失败: " + error.message, "error");
    return false;
  }
}

/**
 * 图片格式化功能
 * @returns {boolean} 操作是否成功
 */
function imageFormat() {
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    // 实现图片格式化逻辑
    // 这里可以添加具体的图片格式化代码
    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "图片格式化完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("图片格式化失败: " + error.message, "error");
    return false;
  }
}

/**
 * 正文格式化功能
 * @returns {boolean} 操作是否成功
 */
function bodyFormat() {
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    // 实现正文格式化逻辑
    // 这里可以添加具体的正文格式化代码
    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "正文格式化完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("正文格式化失败: " + error.message, "error");
    return false;
  }
}

/**
 * 更新目录域功能
 * @returns {boolean} 操作是否成功
 */
function updateTOC() {
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    // 实现更新目录域逻辑
    // 这里可以添加具体的更新目录域代码
    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "目录域更新完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("目录域更新失败: " + error.message, "error");
    return false;
  }
}

/**
 * 全部执行格式化功能
 * @returns {boolean} 操作是否成功
 */
function executeAllFormats() {
  const startTime = performance.now();
  try {
    // 依次执行所有格式化功能
    window.LogModule.addLog("开始执行全部格式化操作", "warning");
    pageFormat();
    titleFormat();
    tableFormat();
    imageFormat();
    bodyFormat();
    updateTOC();

    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "全部格式化操作完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("全部执行失败: " + error.message, "error");
    return false;
  }
}

// 导出格式化功能供其他模块使用
const FormatUtils = {
  pageFormat,
  titleFormat,
  tableFormat,
  imageFormat,
  bodyFormat,
  updateTOC,
  executeAllFormats,
};
