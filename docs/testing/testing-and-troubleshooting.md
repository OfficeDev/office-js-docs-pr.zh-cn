---
title: 排查 Office 加载项中的用户错误
description: 了解如何解决 Office 外接程序中的用户错误。
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 1dbc8cc18e0c9b12ccff605b655dd7c8629fb9cf
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810847"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>排查 Office 加载项中的用户错误

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in. 

还可以使用 [Fiddler](https://www.telerik.com/fiddler) 发现和调试加载项问题。

## <a name="common-errors-and-troubleshooting-steps"></a>常见错误和故障排除步骤

下表列出了用户可能遇到的常见错误消息以及用户可以采取以解决这些错误的步骤。



|**错误消息**|**解决方案**|
|:-----|:-----|
|应用错误：无法访问目录|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|确认已安装最新的 Office 更新，或下载 [Office 2013 更新](https://support.microsoft.com/kb/2986156/)。|
|错误：对象不 支持此属性或方法 "defineProperty"|确认 Internet Explorer 不是在兼容模式下运行。 转到“工具”>“兼容性视图设置”****。|
|Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>安装加载项时，状态栏中会显示“加载加载项时出错”

1. 关闭 Office。
2. 验证清单是否有效
3. 重启加载项
4. 再次安装加载项。

你还可以向我们提供反馈：如果使用 Windows 版 Excel 或 Mac 版 Excel，可以直接从 Excel 向 Office 扩展性团队报告反馈。 若要执行此操作，请选择“**文件**” | “**反馈**” | “**发送哭脸**”。 发送哭脸将提供必要的日志，以帮助我们了解该问题。

## <a name="outlook-add-in-doesnt-work-correctly"></a>Outlook 外接程序不能正常工作

如果在 Windows 上运行并[使用 Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) 的 Outlook 加载项不能正常工作，请尝试在 Internet Explorer 中启用脚本调试。 


- 转到 Tools > **Internet Options**  >  **Advanced**"。
    
- 在“浏览”**** 下，取消选中“禁用脚本调试 (Internet Explorer)”**** 和“禁用脚本调试 (其他)”****。
    
We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.


## <a name="add-in-doesnt-activate-in-office-2013"></a>外接程序在 Office 2013 中无法激活

如果在用户执行下列步骤时外接程序无法激活：


1. 使用 Microsoft 帐户在 Office 2013 中登录。
    
2. 为其 Microsoft 帐户启用两步验证。
    
3. 尝试插入外接程序时在收到提示的时候验证其身份。
    
确认是否已安装最新的 Office 更新程序，或下载 [Office 2013 更新程序](https://support.microsoft.com/kb/2986156/)。


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>外接程序无法在任务窗格中加载，或外接程序清单存在其他问题

请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。


## <a name="add-in-dialog-box-cannot-be-displayed"></a>无法显示外接程序对话框

When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![对话框错误消息的屏幕截图](http://i.imgur.com/3mqmlgE.png)

|**受影响的浏览器**|**受影响的平台**|
|:--------------------|:---------------------|
|Internet Explorer、Microsoft Edge|Office 网页版|

To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.

> [!IMPORTANT]
> 请勿将不信任的加载项的 URL 添加到受信任网站列表中。

要将 URL 添加到受信任站点的列表中，请执行以下操作：

1. 在“**控制面板**”中，转到“**Internet 选项**” > “**安全性**”。
2. 选择“**受信任站点**”区域，并选择“**网站**”。
3. 输入错误消息中显示的 URL，然后选择“**添加**”。
4. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>对加载项命令（包括功能区按钮和菜单项）的更改未生效

如果在清单中进行的更改（如功能区按钮图标的文件名或菜单项的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。 

#### <a name="for-windows"></a>对于 Windows：
删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。

#### <a name="for-mac"></a>对于 Mac：

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>对于 iOS：
Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>对静态文件（例如 JavaScript、HTML 和 CSS）的更改未生效

浏览器可能正在缓存这些文件。 若要阻止此操作，请在开发时关闭客户端缓存。 详细信息取决于你所使用的服务器类型。 在大多数情况下，它涉及将某些标头添加到 HTTP 响应。 我们建议使用以下设置：

- Cache-Control：“private、no-cache、no-store”
- Pragma：“no-cache”
- 过期：“-1”

有关在 Node.JS Express 服务器中执行此操作的示例，请参阅[此 app.js 文件](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)。 有关 ASP.NET 项目中的示例，请参阅[此 cshtml 文件](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)。

如果加载项托管在 Internet Information Server (IIS) 中，则还可以将以下内容添加到 web.config 中。

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

如果这些步骤一开始似乎不起作用，则可能需要清除浏览器的缓存。 请通过浏览器的 UI 执行此操作。 有时，当你尝试在边缘 UI 中清除边缘缓存时，无法成功清除它。 如果出现这种情况，请在 Windows 命令提示符中运行以下命令。

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a>另请参阅

- [在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md) 
- [将 Office 外接程序旁加载到 iPad 和 Mac 上](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [在 iPad 和 Mac 上调试 Office 外接程序](debug-office-add-ins-on-ipad-and-mac.md)  
- [适用于 Visual Studio Code 的 Microsoft Office 外接程序调试器扩展](./debug-with-vs-extension.md)
- [验证 Office 加载项的清单](troubleshoot-manifest.md)
- [使用运行时日志记录功能调试加载项](runtime-logging.md)
