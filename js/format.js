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
        window.LogModule.addLog(`TOC域法成功：找到目录在第${sectionNum}节`, "warning");
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
  // ====== 配置选项 ======
  const CM_TO_POINT = 28.3465;        // 厘米转磅的换算系数
  const PAPER_A4 = 7;                  // Word API 中 A4 纸张对应的常量值
  const ALIGN_CENTER = 1;              // 居中对齐
  const ALIGN_RIGHT = 2;               // 右对齐
  const PAGE_NUMBER_ARABIC = 0;        // 阿拉伯数字页码格式
  // =======================

  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    window.LogModule.addLog("开始执行页面格式化", "warning");

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
    window.LogModule.addLog("页面基础设置完成：A4规格，2.5厘米页边距，1.5厘米页眉距离，1.75厘米页脚距离", "success");

    // 步骤3：恢复所有节的原有页面方向
    for (let i = 1; i <= doc.Sections.Count; i++) {
      doc.Sections.Item(i).PageSetup.Orientation = originalOrientations[i - 1];
    }
    window.LogModule.addLog("页面方向恢复完成", "success");

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
        window.LogModule.addLog("已在目录页前一页末尾插入分节符", "success");
        
        // 插入分节符后，目录所在的节索引会增加1
        startSectionIndex = tocInfo.sectionIndex + 1;
        window.LogModule.addLog(`目录现在位于第${startSectionIndex}节`, "warning");
      } catch (e) {
        window.LogModule.addLog(`插入分节符失败: ${e.message}`, "warning");
        startSectionIndex = tocInfo.sectionIndex;
      }
    } else {
      window.LogModule.addLog("未能找到目录，将从第1页开始配置连续页码", "warning");
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
      
      window.LogModule.addLog(`起始节（第${startSectionIndex}节）页眉页脚配置完成`, "success");
      
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
        window.LogModule.addLog("后续节配置完成", "success");
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
 * @description 用于格式化文档中的标题样式，包括9个标题样式，以及清理标题中的多余字符。
 * @returns {boolean} 操作是否成功
 */
function titleFormat() {
  // ====== 功能开关 ======
  // 是否移除标题编号（如："1. 标题" -> "标题"）
  const removeNumbering = true;
  // 是否刷新标题样式（重新应用标题1-9样式）
  const refreshStyles = true;
  // 是否显示确认弹窗
  const showConfirm = true;
  // =======================

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
        
        // 重置选择范围，防止上一次选择影响当前操作
        window.Application.Selection.Collapse();
        
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
      window.LogModule.addLog(`处理完成！共处理了 ${processedCount} 个标题。`, "success");
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
    window.LogModule.addLog(`标题格式化完成，耗时：${duration}秒`, "success");
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
 * @description 统一字体、首行加粗、单元格居中、表格自适应窗口宽度、识别价格列
 * @returns {boolean} 操作是否成功
 */
function tableFormat() {
  // ====== 配置选项 ======
  const CONFIG = {
    FONT_NAME: "宋体",           // 表格字体名称
    FONT_SIZE: 10.5,            // 表格字体大小
    DECIMAL_PLACES: 2,          // 价格列保留小数位数
    PRICE_KEYWORDS: ["单价","总价","元"], // 识别价格列的关键词
  };
  // Word API 常量定义
  const WD_ALIGN_PARAGRAPH_CENTER = 1;       // 段落居中对齐
  const WD_CELL_ALIGN_VERTICAL_CENTER = 1;   // 单元格垂直居中
  const WD_AUTO_FIT_WINDOW = 2;              // 表格自适应窗口宽度
  // =======================

  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    window.LogModule.addLog("开始执行表格格式化", "warning");

    // 设置表格基本格式（字体、对齐方式、自适应）
    function setTableBasicFormat(table, tableIndex) {
      if (!table) return;
      
      window.LogModule.addLog(`开始处理第${tableIndex}个表格基本格式`, "warning");
      
      const range = table.Range;
      
      // 1. 统一设置字体和字号
      const originalFontName = range.Font.Name;
      const originalFontSize = range.Font.Size;
      range.Font.Name = CONFIG.FONT_NAME;
      range.Font.Size = CONFIG.FONT_SIZE;
      window.LogModule.addLog(`  字体设置: "${originalFontName}"->"${CONFIG.FONT_NAME}", ${originalFontSize}pt->${CONFIG.FONT_SIZE}pt`, "info");
      
      // 2. 设置所有单元格的对齐方式
      range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER;
      range.Cells.VerticalAlignment = WD_CELL_ALIGN_VERTICAL_CENTER;
      window.LogModule.addLog(`  对齐方式: 水平居中、垂直居中`, "info");
      
      // 3. 设置首行加粗 - 使用多策略方案处理各种情况
      window.LogModule.addLog(`  开始设置首行加粗...`, "info");
      let success = false;
      
      // 策略1：直接使用Row.Range（简单情况）
      try {
        if (table.Rows && table.Rows.Count > 0) {
          const firstRow = table.Rows.Item(1);
          if (firstRow && firstRow.Range) {
            firstRow.Range.Font.Bold = true;
            window.LogModule.addLog(`  首行加粗完成（策略1：直接使用Row.Range）`, "info");
            success = true;
          }
        }
      } catch (e) {
        window.LogModule.addLog(`  策略1失败: ${e.message || e}`, "warning");
      }
      
      // 策略2：逐个单元格设置（处理合并单元格）
      if (!success) {
        try {
          let cellCount = 0;
          for (let col = 1; col <= table.Columns.Count + 20; col++) {
            try {
              const cell = table.Cell(1, col);
              if (cell && cell.Range) {
                cell.Range.Font.Bold = true;
                cellCount++;
              }
            } catch (e) {
              break;
            }
          }
          if (cellCount > 0) {
            window.LogModule.addLog(`  首行加粗完成（策略2：逐个单元格设置，共${cellCount}个）`, "info");
            success = true;
          } else {
            window.LogModule.addLog(`  首行加粗失败：没有找到可设置的单元格`, "error");
          }
        } catch (e) {
          window.LogModule.addLog(`  策略2失败: ${e.message || e}`, "error");
        }
      }
      
      // 4. 设置表格自适应窗口
      table.AutoFitBehavior(WD_AUTO_FIT_WINDOW);
      window.LogModule.addLog(`  表格自适应窗口设置完成`, "info");
      window.LogModule.addLog(`第${tableIndex}个表格基本格式处理完成`, "success");
    }

    // 识别并处理价格列
    function processPriceColumns(table, tableIndex) {
      if (!table || table.Rows.Count === 0 || table.Columns.Count === 0) return;

      window.LogModule.addLog(`开始处理第${tableIndex}个表格价格列`, "warning");

      const priceColumns = [];
      
      const foundColumns = new Set();
      
      for (let row = 1; row <= 2 && row <= table.Rows.Count; row++) {
        for (let col = 1; col <= table.Columns.Count; col++) {
          try {
            const cell = table.Cell(row, col);
            const cellText = cell.Range.Text.replace(/\r\a/g, "");
            
            const isPriceColumn = CONFIG.PRICE_KEYWORDS.some(
              (keyword) => cellText.indexOf(keyword) > -1
            );
            
            if (isPriceColumn && !foundColumns.has(col)) {
              foundColumns.add(col);
              priceColumns.push(col);
              window.LogModule.addLog(`  识别到价格列: 第${col}列（表头行${row}: "${cellText}"）`, "info");
            }
            
            const mergeArea = cell.MergeArea;
            if (mergeArea && mergeArea.Columns && mergeArea.Columns.Count > 1) {
              col += mergeArea.Columns.Count - 1;
            }
          } catch (e) {
            continue;
          }
        }
      }
      
      if (priceColumns.length === 0) {
        window.LogModule.addLog(`  未识别到价格列`, "warning");
        return;
      }

      let modifiedCellCount = 0;
      
      for (let row = 2; row <= table.Rows.Count; row++) {
        for (let i = 0; i < priceColumns.length; i++) {
          const col = priceColumns[i];
          try {
            const cell = table.Cell(row, col);
            const cellText = cell.Range.Text.replace(/\r\a/g, "");
            
            if (cellText.match(/\d/)) {
              const numValue = parseFloat(cellText.replace(/[^\d.-]/g, ""));
              if (!isNaN(numValue)) {
                const newValue = numValue.toFixed(CONFIG.DECIMAL_PLACES);
                if (cellText !== newValue) {
                  cell.Range.Text = newValue;
                  window.LogModule.addLog(`  修改单元格[${row},${col}]: "${cellText}"->"${newValue}"`, "info");
                  modifiedCellCount++;
                }
              }
            }
          } catch (e) {
            continue;
          }
        }
      }
      
      window.LogModule.addLog(`第${tableIndex}个表格价格列处理完成，共修改 ${modifiedCellCount} 个单元格`, "success");
    }

    const tables = doc.Tables;
    if (tables.Count === 0) {
      window.LogModule.addLog("文档中未找到表格！", "warning");
      const endTime = performance.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);
      window.LogModule.addLog("表格格式化完成，耗时：" + duration + "秒", "success");
      return true;
    }

    window.LogModule.addLog(`文档中共找到 ${tables.Count} 个表格`, "warning");

    for (let i = 1; i <= tables.Count; i++) {
      const table = tables.Item(i);
      window.LogModule.addLog(`正在处理第${i}个表格（共${tables.Count}个）`, "warning");
      setTableBasicFormat(table, i);
      processPriceColumns(table, i);
    }

    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(`成功格式化了 ${tables.Count} 个表格！`, "success");
    window.LogModule.addLog("表格格式化完成，耗时：" + duration + "秒", "success");
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
  // ====== 配置选项 ======
  const PICTURE_STYLE_NAME = "图片";
  const MSO_PICTURE = 13;
  const MIN_HEIGHT_POINTS = 50; // 排除大约2cm以内的签名图片
  const MAX_HEIGHT_CM = 21;
  const MAX_WIDTH_CM = 16;
  const CM_TO_POINT = 28.3465; // 厘米转磅的换算系数
  // =======================
  
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    window.LogModule.addLog("开始执行图片格式化", "warning");

    // 记录统计数据
    let convertedCount = 0,
        styledCount = 0,
        resizedCount = 0;

    // 验证样式存在
    let pictureStyleExists = false;
    
    try {
      const style = doc.Styles.Item(PICTURE_STYLE_NAME);
      if (style && style.NameLocal === PICTURE_STYLE_NAME) {
        pictureStyleExists = true;
        window.LogModule.addLog(`"图片"样式验证通过`, "info");
      }
    } catch (e) {
      window.LogModule.addLog(`检查样式时发生错误: ${e.message}`, "warning");
    }
    
    if (!pictureStyleExists) {
      window.LogModule.addLog("文档中不存在'图片'样式，请手动创建！", "warning");
      const endTime = performance.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);
      window.LogModule.addLog("图片格式化完成，耗时：" + duration + "秒", "success");
      return false;
    }

    // 转换浮动图片为嵌入式
    window.LogModule.addLog(`开始转换浮动图片，共 ${doc.Shapes.Count} 个形状`, "warning");
    for (let i = doc.Shapes.Count; i >= 1; i--) {
      try {
        const shape = doc.Shapes.Item(i);
        if (shape.Type === MSO_PICTURE && shape.Height >= MIN_HEIGHT_POINTS) {
          shape.ConvertToInlineShape();
          convertedCount++;
        } else if (shape.Type === MSO_PICTURE && shape.Height < MIN_HEIGHT_POINTS) {
          const pageNum = shape.Anchor.Information(1); // wdActiveEndPageNumber
          window.LogModule.addLog(`疑似签名图片被排除，位于第${pageNum}页`, "info");
        }
      } catch (e) {
        window.LogModule.addLog(`转换第${i}个形状失败: ${e.message}`, "warning");
        continue;
      }
    }
    window.LogModule.addLog(`浮动图片转换完成，共转换 ${convertedCount} 张`, "success");

    // 应用图片样式
    window.LogModule.addLog(`开始验证图片尺寸并应用图片样式，共 ${doc.InlineShapes.Count} 张嵌入式图片`, "warning");
    const maxLandscapeLongSide = MAX_WIDTH_CM * CM_TO_POINT;     // 横向图片长边限制
    const maxPortraitLongSide = MAX_HEIGHT_CM * CM_TO_POINT;    // 竖向图片长边限制
    const EPSILON = 0.1;  // 允许的微小误差阈值（磅），避免精度问题
    
    for (let i = 1; i <= doc.InlineShapes.Count; i++) {
      try {
        const inlineShape = doc.InlineShapes.Item(i);
        const pageNum = inlineShape.Range.Information(1); // wdActiveEndPageNumber
        
        // 转换为Shape获取旋转角度
        const shape = inlineShape.ConvertToShape();
        const rotation = shape.Rotation;
        const shapeWidth = shape.Width;
        const shapeHeight = shape.Height;
        
        // 判断视觉方向
        let isLandscape;
        if (rotation === 90 || rotation === 270) {
          // 旋转90/270度，视觉方向与原始相反
          isLandscape = shapeHeight > shapeWidth;
        } else {
          // 0/180度，视觉方向与原始相同
          isLandscape = shapeWidth > shapeHeight;
        }
        const directionText = isLandscape ? "横向" : "竖向";
        const maxLongSide = isLandscape ? maxLandscapeLongSide : maxPortraitLongSide;
        
        // 找出当前长边
        const currentLongSide = Math.max(shapeWidth, shapeHeight);
        
        // 检查是否需要调整（增加误差阈值）
        if (currentLongSide > maxLongSide + EPSILON) {
          shape.LockAspectRatio = -1; // msoTrue
          
          // 直接设置长边为限制值，让Word自动处理短边
          if (shapeWidth >= shapeHeight) {
            // 原始宽度是长边
            shape.Width = maxLongSide;
          } else {
            // 原始高度是长边
            shape.Height = maxLongSide;
          }
          
          const origLongCm = (currentLongSide / CM_TO_POINT).toFixed(2);
          const newLongCm = (maxLongSide / CM_TO_POINT).toFixed(2);
          window.LogModule.addLog(`第${pageNum}页第${i}张${directionText}图片，长边${origLongCm}cm→${newLongCm}cm`, "info");
          resizedCount++;
        }
        
        // 转换回嵌入式图片并应用样式
        const newInlineShape = shape.ConvertToInlineShape();
        
        // 检查当前样式，避免重复应用（优化性能）
        const currentStyleName = newInlineShape.Range.Style.NameLocal;
        if (currentStyleName !== PICTURE_STYLE_NAME) {
          newInlineShape.Range.Style = PICTURE_STYLE_NAME;
          styledCount++;
        }
      } catch (e) {
        window.LogModule.addLog(`应用样式到第${i}张图片失败: ${e.message}`, "warning");
        continue;
      }
    }
    window.LogModule.addLog(`共调整尺寸 ${resizedCount} 张，共应用样式 ${styledCount} 张`, "success");

    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog("图片格式化完成，耗时：" + duration + "秒", "success");
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
    window.LogModule.addLog("正文格式化完成，耗时：" + duration + "秒", "success");
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
    window.LogModule.addLog("目录域更新完成，耗时：" + duration + "秒", "success");
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
    window.LogModule.addLog("全部格式化操作完成，耗时：" + duration + "秒", "success");
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
