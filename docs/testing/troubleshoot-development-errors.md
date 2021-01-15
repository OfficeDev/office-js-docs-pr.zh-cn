---
title: Office 加载项开发错误疑难解答
description: 了解如何解决 Office 外接程序中的开发错误。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 48216230db4bf90ca53ef10d98786877bd3905c2
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771422"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Office 加载项开发错误疑难解答

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>外接程序无法在任务窗格中加载，或外接程序清单存在其他问题

请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>对加载项命令（包括功能区按钮和菜单项）的更改未生效

如果在清单中进行的更改（如功能区按钮图标的文件名或菜单项的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。 

#### <a name="for-windows"></a>对于 Windows：

删除文件夹的内容， `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 并删除文件夹的内容（ `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` 如果存在）。

#### <a name="for-mac"></a>对于 Mac：

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>对于 iOS：
在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>对属性值所做的更改不会发生，并且没有错误消息

查看属性的参考文档，以查看其是否为只读。 此外，Office JS [的 TypeScript 定义](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) 指定哪些对象属性是只读的。 如果尝试设置只读属性，写入操作将失败，无提示，不会引发任何错误。 以下示例错误地尝试将只读属性设置为 [Chart.id。](/javascript/api/excel/excel.chart#id)另请参阅 [某些属性无法直接设置](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>收到错误："此外接程序不再可用"

以下是导致此错误的一些原因。 如果发现其他原因，请使用页面底部的反馈工具告诉我们。

- 如果使用 Visual Studio，则旁加载可能有问题。 关闭 Office 主机的所有实例Visual Studio。 重新启动Visual Studio并再次尝试按 F5。
- 外接程序的清单已从其部署位置（如集中部署、SharePoint 目录或网络共享）中删除。
- 清单 [中 ID](../reference/manifest/id.md) 元素的值已直接在已部署的副本中更改。 如果出于任何原因需要更改此 ID，请首先从 Office 主机中删除外接程序，然后将原始清单替换为已更改的清单。 许多用户需要清除 Office 缓存以删除原始缓存的所有跟踪。 请参阅本文前面部分对外接程序命令 [（包括功能区按钮](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) 和菜单项）的更改不会生效。
- 加载项清单的清单在清单的"资源"部分的任何位置未定义，或者其使用位置和定义位置之间的拼写不匹配。 `resid` [](../reference/manifest/resources.md) `resid` `<Resources>`
- 清单 `resid` 中的某位置有一个超过 32 个字符的属性。 属性和节中相应资源的属性不能超过 `resid` `id` `<Resources>` 32 个字符。
- 加载项具有自定义外接程序命令，但您尝试在不支持这些命令的平台上运行该命令。 有关详细信息，请参阅 [外接程序命令要求集](../reference/requirement-sets/add-in-commands-requirement-sets.md)。

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>加载项在 Edge 上不起作用，但它适用于其他浏览器

请参阅 [Microsoft Edge 问题疑难解答](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel 加载项引发错误，但不一致

有关 [可能的原因，请参阅 Excel](../excel/excel-add-ins-troubleshooting.md) 加载项疑难解答。

## <a name="see-also"></a>另请参阅

- [在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)
- [将 Office 外接程序旁加载到 iPad 和 Mac 上](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [在 iPad 和 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)  
- [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)
- [验证 Office 加载项的清单](troubleshoot-manifest.md)
- [使用运行时日志记录功能调试加载项](runtime-logging.md)
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)
