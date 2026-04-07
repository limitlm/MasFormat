// 格式化相关功能实现

/**
 * 查找目录信息
 * @param {Object} doc - Word文档对象
 * @returns {Object|null} 包含目录节索引和Range的对象，如果未找到则返回null
 */
function findTOCInfo(doc) {
  try {
    // 查找TOC域（Table of Contents）
    if (doc.TablesOfContents && doc.TablesOfContents.Count > 0) {
      try {
        const toc = doc.TablesOfContents.Item(1);
        const tocRange = toc.Range;
        const sectionNum = tocRange.Information(2); // wdActiveEndSectionNumber
        window.LogModule.addLog(`TOC域法成功：找到目录在第${sectionNum}节`, "info");
        return {
          sectionIndex: sectionNum,
          range: tocRange
        };
      } catch (e) {
        window.LogModule.addLog(`TOC域检测失败: ${e.message}`, "warning");
      }
    } else {
      window.LogModule.addLog("未找到TOC域", "warning");
    }

    return null;
  } catch (error) {
    window.LogModule.addLog(`查找目录信息失败: ${error.message}`, "warning");
    return null;
  }
}

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
    window.LogModule.addLog("页面基础设置完成：A4规格，2.5厘米页边距，1.5厘米页眉距离，1.75厘米页脚距离", "info");

    // 步骤3：恢复所有节的原有页面方向
    for (let i = 1; i <= doc.Sections.Count; i++) {
      doc.Sections.Item(i).PageSetup.Orientation = originalOrientations[i - 1];
    }
    window.LogModule.addLog("页面方向恢复完成", "info");

    // 步骤4：查找目录信息
    let tocInfo = findTOCInfo(doc);
    let hasTOC = !!tocInfo;
    let startSectionIndex = 1;
    
    if (hasTOC) {
      // 步骤5：在目录所在页的前一页末尾插入分节符（下一页）
      // 这样可以确保目录所在页的整页内容都在新节里
      try {
        let insertRange;
        try {
          // 找到目录所在页的页码
          const tocPageNum = tocInfo.range.Information(1); // wdActiveEndPageNumber
          window.LogModule.addLog(`目录位于第${tocPageNum}页`, "info");
          
          if (tocPageNum > 1) {
            // 目录不在第1页，跳转到前一页的末尾
            insertRange = doc.Range(0, 0);
            insertRange.GoTo(1, 1, tocPageNum - 1); // 跳转到前一页
            insertRange.GoTo(3, 1); // wdGoToLine, wdGoToLast - 跳转到当前页的最后一行
            window.LogModule.addLog(`定位到第${tocPageNum - 1}页末尾`, "info");
          } else {
            // 目录在第1页，直接在文档开头插入分节符
            insertRange = doc.Range(0, 0);
            window.LogModule.addLog("目录在第1页，在文档开头插入分节符", "info");
          }
        } catch (e) {
          window.LogModule.addLog(`定位失败: ${e.message}，使用目录起始位置作为后备`, "warning");
          // 后备方案：直接使用tocRange的起始位置
          insertRange = doc.Range(tocInfo.range.Start, tocInfo.range.Start);
        }
        
        // 在定位的位置插入分节符
        insertRange.Collapse(0); // wdCollapseEnd - 在范围末尾插入
        insertRange.InsertBreak(2); // wdSectionBreakNextPage
        window.LogModule.addLog("已在目录页前一页末尾插入分节符", "info");
        
        // 插入分节符后，目录所在的节索引会增加1
        startSectionIndex = tocInfo.sectionIndex + 1;
        window.LogModule.addLog(`目录现在位于第${startSectionIndex}节`, "info");
      } catch (e) {
        window.LogModule.addLog(`插入分节符失败: ${e.message}`, "warning");
        startSectionIndex = tocInfo.sectionIndex;
      }
    } else {
      window.LogModule.addLog("未能找到目录，将从第1页开始配置连续页码", "info");
    }
    
    if (doc.Sections.Count < startSectionIndex) {
      window.LogModule.addLog("文档节数不足，无法找到起始节", "warning");
    } else {
      // 步骤6：配置起始节
      const startSection = doc.Sections.Item(startSectionIndex);
      const startFooter = startSection.Footers.Item(1);
      
      // 清除原有页脚内容
      startFooter.Range.Delete();
      
      // 设置不使用首页不同
      startSection.PageSetup.DifferentFirstPageHeaderFooter = false;
      startSection.PageSetup.PageNumberStyle = PAGE_NUMBER_ARABIC;
      
      // 如果有目录，不链接到前节；如果没有目录，保持链接
      if (hasTOC && startSectionIndex > 1) {
        try {
          startSection.Headers.Item(1).LinkToPrevious = false;
          startSection.Footers.Item(1).LinkToPrevious = false;
        } catch (e) {
          window.LogModule.addLog(`警告：起始节页眉页脚断开同前节链接失败 - ${e.description}`, "warning");
        }
      }
      
      // 添加公司名称
      const startFooterRange = startFooter.Range;
      startFooterRange.Text = "重庆梅安森科技股份有限公司 编制";
      startFooterRange.ParagraphFormat.Alignment = ALIGN_RIGHT;
      
      // 插入换行符和页码
      startFooterRange.Collapse(1);
      startFooterRange.Text = "\n";
      startFooterRange.MoveEnd(1, -1);
      const startPageField = startFooterRange.Fields.Add(startFooterRange, -1, "PAGE", false);
      startPageField.Code.ParagraphFormat.Alignment = ALIGN_CENTER;
      
      // 有目录时从1开始，无目录时保持连续
      if (hasTOC) {
        startFooter.PageNumbers.RestartNumberingAtSection = true;
        startFooter.PageNumbers.StartingNumber = 1;
      } else {
        startFooter.PageNumbers.RestartNumberingAtSection = false;
      }
      
      window.LogModule.addLog(`起始节（第${startSectionIndex}节）页眉页脚配置完成`, "info");
      
      // 步骤7：配置后续节
      if (doc.Sections.Count > startSectionIndex) {
        for (let i = startSectionIndex + 1; i <= doc.Sections.Count; i++) {
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
        window.LogModule.addLog("后续节配置完成", "info");
      }
    }

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
