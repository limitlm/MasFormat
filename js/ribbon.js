// 引入格式化工具函数
// 注意：在实际运行环境中，需要确保format.js已被正确加载

//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
  if (typeof window.Application.ribbonUI != "object") {
    window.Application.ribbonUI = ribbonUI;
  }

  if (typeof window.Application.Enum != "object") {
    // 如果没有内置枚举值
    window.Application.Enum = WPS_Enum;
  }

  window.Application.PluginStorage.setItem("EnableFlag", false); //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
  window.Application.PluginStorage.setItem("ApiEventFlag", false); //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
  return true;
}

var WebNotifycount = 0;
function OnAction(control) {
  const eleId = control.Id;
  switch (eleId) {
    case "btnShowMsg":
      {
        const doc = window.Application.ActiveDocument;
        if (!doc) {
          alert("当前没有打开任何文档");
          return;
        }
        alert(doc.Name);
      }
      break;
    case "btnIsEnbable": {
      let bFlag = window.Application.PluginStorage.getItem("EnableFlag");
      window.Application.PluginStorage.setItem("EnableFlag", !bFlag);

      //通知wps刷新以下几个按饰的状态
      window.Application.ribbonUI.InvalidateControl("btnIsEnbable");
      window.Application.ribbonUI.InvalidateControl("btnShowDialog");
      window.Application.ribbonUI.InvalidateControl("btnShowTaskPane");
      //window.Application.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
      break;
    }
    case "btnShowDialog":
      window.Application.ShowDialog(
        GetUrlPath() + "/ui/dialog.html",
        "这是一个对话框网页",
        400 * window.devicePixelRatio,
        400 * window.devicePixelRatio,
        false,
      );
      break;
    case "btnShowTaskPane":
      {
        let tsId = window.Application.PluginStorage.getItem("taskpane_id");
        if (!tsId) {
          let tskpane = window.Application.CreateTaskPane(
            GetUrlPath() + "/ui/taskpane.html",
          );
          let id = tskpane.ID;
          window.Application.PluginStorage.setItem("taskpane_id", id);
          tskpane.Visible = true;
        } else {
          let tskpane = window.Application.GetTaskPane(tsId);
          tskpane.Visible = !tskpane.Visible;
        }
      }
      break;
    case "btnApiEvent":
      {
        let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag");
        let bRegister = bFlag ? false : true;
        window.Application.PluginStorage.setItem("ApiEventFlag", bRegister);
        if (bRegister) {
          window.Application.ApiEvent.AddApiEventListener(
            "DocumentNew",
            OnNewDocumentApiEvent,
          );
        } else {
          window.Application.ApiEvent.RemoveApiEventListener(
            "DocumentNew",
            OnNewDocumentApiEvent,
          );
        }

        window.Application.ribbonUI.InvalidateControl("btnApiEvent");
      }
      break;
    case "btnWebNotify":
      {
        let currentTime = new Date();
        let timeStr =
          currentTime.getHours() +
          ":" +
          currentTime.getMinutes() +
          ":" +
          currentTime.getSeconds();
        window.Application.OAAssist.WebNotify(
          "这行内容由wps加载项主动送达给业务系统，可以任意自定义, 比如时间值:" +
            timeStr +
            "，次数：" +
            ++WebNotifycount,
          true,
        );
      }
      break;
    case "btnOpenLog":
      {
        // 打开日志窗口，加载logViewer.html页面
        window.Application.ShowDialog(
          GetUrlPath() + "/ui/logViewer.html",
          "Mas标书格式化 - 执行日志",
          500 * window.devicePixelRatio,
          635 * window.devicePixelRatio,
          false,
        );
      }
      break;
    case "btnUsageGuide":
      {
        // 打开使用说明窗口，加载index.html页面
        window.Application.ShowDialog(
          GetUrlPath() + "/index.html",
          "Mas标书格式化 - 使用说明",
          800 * window.devicePixelRatio,
          600 * window.devicePixelRatio,
          false,
        );
      }
      break;
    case "btnPageFormat":
      {
        // 调用页面格式化功能
        if (typeof pageFormat === "function") {
          pageFormat();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.pageFormat === "function") {
          FormatUtils.pageFormat();
        } else {
          alert("页面格式化功能未加载");
        }
      }
      break;
    case "btnTitleFormat":
      {
        // 调用标题格式化功能
        if (typeof titleFormat === "function") {
          titleFormat();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.titleFormat === "function") {
          FormatUtils.titleFormat();
        } else {
          alert("标题格式化功能未加载");
        }
      }
      break;
    case "btnTableFormat":
      {
        // 调用表格格式化功能
        if (typeof tableFormat === "function") {
          tableFormat();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.tableFormat === "function") {
          FormatUtils.tableFormat();
        } else {
          alert("表格格式化功能未加载");
        }
      }
      break;
    case "btnImageFormat":
      {
        // 调用图片格式化功能
        if (typeof imageFormat === "function") {
          imageFormat();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.imageFormat === "function") {
          FormatUtils.imageFormat();
        } else {
          alert("图片格式化功能未加载");
        }
      }
      break;
    case "btnBodyTextFormat":
      {
        // 调用正文格式化功能
        if (typeof bodyTextFormat === "function") {
          bodyTextFormat();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.bodyTextFormat === "function") {
          FormatUtils.bodyTextFormat();
        } else {
          alert("正文格式化功能未加载");
        }
      }
      break;
    case "btnUpdateTOC":
      {
        // 调用更新目录域功能
        if (typeof updateTOC === "function") {
          updateTOC();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.updateTOC === "function") {
          FormatUtils.updateTOC();
        } else {
          alert("更新目录域功能未加载");
        }
      }
      break;
    case "btnExecuteAll":
      {
        // 调用全部执行格式化功能
        if (typeof executeAllFormats === "function") {
          executeAllFormats();
        } else if (typeof FormatUtils !== "undefined" && typeof FormatUtils.executeAllFormats === "function") {
          FormatUtils.executeAllFormats();
        } else {
          alert("全部执行功能未加载");
        }
      }
      break;
    default:
      break;
  }
  return true;
}

function GetImage(control) {
  const eleId = control.Id;
  switch (eleId) {
    case "btnShowMsg":
      return "images/1.svg";
    case "btnShowDialog":
      return "images/2.svg";
    case "btnShowTaskPane":
      return "images/3.svg";
    default:
  }
  return "images/newFromTemp.svg";
}

function OnGetEnabled(control) {
  const eleId = control.Id;
  switch (eleId) {
    case "btnShowMsg":
      return true;
      break;
    case "btnShowDialog": {
      let bFlag = window.Application.PluginStorage.getItem("EnableFlag");
      return bFlag;
      break;
    }
    case "btnShowTaskPane": {
      let bFlag = window.Application.PluginStorage.getItem("EnableFlag");
      return bFlag;
      break;
    }
    default:
      break;
  }
  return true;
}

function OnGetVisible(control) {
  return true;
}

function OnGetLabel(control) {
  const eleId = control.Id;
  switch (eleId) {
    case "btnIsEnbable": {
      let bFlag = window.Application.PluginStorage.getItem("EnableFlag");
      return bFlag ? "按钮Disable" : "按钮Enable";
      break;
    }
    case "btnApiEvent": {
      let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag");
      return bFlag ? "清除新建文件事件" : "注册新建文件事件";
      break;
    }
  }
  return "";
}

function OnNewDocumentApiEvent(doc) {
  alert("新建文件事件响应，取文件名: " + doc.Name);
}
