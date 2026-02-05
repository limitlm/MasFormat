// 格式化相关功能实现

/**
 * 页面格式化功能
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
    window.LogModule.addLog("开始执行页面格式化", "info");
    // 实现页面格式化逻辑
    // 1. 页面设置：A4规格，2.5厘米页边距（全局设置）
    doc.PageSetup.PaperSize = 7;
    doc.PageSetup.LeftMargin = 2.5 * 28.3465;
    doc.PageSetup.RightMargin = 2.5 * 28.3465;
    doc.PageSetup.TopMargin = 2.5 * 28.3465;
    doc.PageSetup.BottomMargin = 2.5 * 28.3465;
    window.LogModule.addLog("页面设置完成：A4规格，2.5厘米页边距", "info");

    // 2. 页眉页脚设置（全局设置）
    doc.PageSetup.HeaderDistance = 1.5 * 28.3465;
    doc.PageSetup.FooterDistance = 1.75 * 28.3465;
    window.LogModule.addLog(
      "页眉页脚设置完成：页眉1.5厘米，页脚1.75厘米",
      "info",
    );

    // 获取第一节
    var section1 = doc.Sections.Item(1);

    // 获取页脚对象
    var footer = section1.Footers.Item(1);

    // 清除原有页脚内容
    footer.Range.Delete();
    window.LogModule.addLog("清除原有页脚内容完成", "info");

    // 3. 设置页码：在第二页添加阿拉伯数字页码，首页不显示
    // 设置首页不同（仅第一节）
    section1.PageSetup.DifferentFirstPageHeaderFooter = true;
    window.LogModule.addLog("设置首页不同页脚完成", "info");

    // 设置页码样式为阿拉伯数字
    section1.PageSetup.PageNumberStyle = 0;
    window.LogModule.addLog("设置页码样式阿拉伯数字完成", "info");

    // 4. 在主页脚中添加页码和公司名称
    // 创建页脚范围
    var footerRange = footer.Range;

    // 先添加公司名称
    footerRange.Text = "重庆梅安森科技股份有限公司 编制";
    window.LogModule.addLog("添加公司名称到页脚完成", "info");

    // 设置右对齐
    footerRange.ParagraphFormat.Alignment = 2;
    window.LogModule.addLog("设置公司名称右对齐完成", "info");

    // 移动到行首并插入换行符
    footerRange.Collapse(1);
    footerRange.Text = "\n";
    footerRange.MoveEnd(1, -1); // 移动到换行符前
    window.LogModule.addLog("插入换行符完成", "info");

    // 在换行符后插入页码字段
    var pageField = footerRange.Fields.Add(footerRange, -1, "PAGE", false);
    window.LogModule.addLog("插入页码字段完成", "info");

    // 设置页码居中
    pageField.Code.ParagraphFormat.Alignment = 1;
    window.LogModule.addLog("设置页码居中完成", "info");

    // 5. 设置节设置：确保页眉页脚同前节，页码连续
    window.LogModule.addLog("开始设置其他节属性", "info");

    // 遍历所有节（从第二节开始）
    for (var i = 2; i <= doc.Sections.Count; i++) {
      try {
        var section = doc.Sections.Item(i);

        // 对于其他节，取消首页不同设置，确保每节第一页也有页码
        section.PageSetup.DifferentFirstPageHeaderFooter = false;

        // 设置页眉页脚同前节
        try {
          section.Headers.Item(1).LinkToPrevious = true;
          section.Footers.Item(1).LinkToPrevious = true;
        } catch (headerFooterError) {
          window.LogModule.addLog(
            "警告：设置第" +
              i +
              "节页眉页脚同前节失败 - " +
              headerFooterError.description,
            "warning",
          );
        }

        // 确保页码连续
        section.Footers.Item(1).PageNumbers.RestartNumberingAtSection = false;
        window.LogModule.addLog("设置第" + i + "节属性完成", "info");
      } catch (sectionError) {
        window.LogModule.addLog(
          "错误：处理第" + i + "节失败 - " + sectionError.description,
          "error",
        );
      }
    }

    window.LogModule.addLog("所有节属性设置完成", "info");

    // 6. 确保页码从1开始，首页设置为0
    section1.Footers.Item(1).PageNumbers.RestartNumberingAtSection = true;
    section1.Footers.Item(1).PageNumbers.StartingNumber = 0;
    window.LogModule.addLog("页码起始设置完成", "info");

    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "页面格式化完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("页面格式化失败: " + error.message, "error");
    return false;
  }
}

/**
 * 标题格式化功能
 * @returns {boolean} 操作是否成功
 */
function titleFormat() {
  const startTime = performance.now();
  const doc = window.Application.ActiveDocument;
  if (!doc) {
    window.LogModule.addLog("当前没有打开任何文档", "warning");
    return false;
  }

  try {
    // 实现标题格式化逻辑
    // 这里可以添加具体的标题格式化代码
    const endTime = performance.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    window.LogModule.addLog(
      "标题格式化完成，耗时：" + duration + "秒",
      "success",
    );
    return true;
  } catch (error) {
    window.LogModule.addLog("标题格式化失败: " + error.message, "error");
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
function bodyTextFormat() {
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
    window.LogModule.addLog("开始执行全部格式化操作", "info");
    pageFormat();
    titleFormat();
    tableFormat();
    imageFormat();
    bodyTextFormat();
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
  bodyTextFormat,
  updateTOC,
  executeAllFormats,
};
