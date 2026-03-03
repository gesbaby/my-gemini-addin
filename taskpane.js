Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("app-body").innerHTML = "Gemini 插件已在线激活！";
  }
});