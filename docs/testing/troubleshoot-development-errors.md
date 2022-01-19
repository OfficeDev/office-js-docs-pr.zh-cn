---
title: 排查 Office 加载项中的开发错误
description: 了解如何解决加载项中的Office错误。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5c8c17077295313b4f10874a851c4d9d6dbef62b
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074313"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>排查 Office 加载项中的开发错误

下面列出了在开发加载项时可能会遇到的Office问题。

> [!TIP]
> 清除Office缓存通常会修复与过时代码相关的问题。 这可确保使用当前文件名、菜单文本和其他命令元素上载最新的清单。 若要了解更多信息，请参阅[清除Office缓存。](clear-cache.md)

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>外接程序无法在任务窗格中加载，或外接程序清单存在其他问题

请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>对加载项命令（包括功能区按钮和菜单项）的更改未生效

清除缓存有助于确保使用外接程序清单的最新版本。 若要清除Office缓存，请按照清除缓存Office[中的说明操作](clear-cache.md)。 如果你使用的是 Office web 版，请通过浏览器的 UI 清除浏览器的缓存。

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>对静态文件（例如 JavaScript、HTML 和 CSS）的更改未生效

浏览器可能正在缓存这些文件。 若要阻止此操作，请在开发时关闭客户端缓存。 详细信息取决于你所使用的服务器类型。 在大多数情况下，它涉及将某些标头添加到 HTTP 响应。 建议设置以下集合。

- Cache-Control：“private、no-cache、no-store”
- Pragma：“no-cache”
- 过期：“-1”

有关在 Node.JS Express 服务器中执行此操作的示例，请参阅[此 app.js 文件](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js)。 有关 ASP.NET 项目中的示例，请参阅[此 cshtml 文件](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)。

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

检查属性的参考文档，以查看其是否为只读。 此外，Office JS 的[TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)定义指定哪些对象属性是只读的。 如果您尝试设置只读属性，写入操作将失败，无提示，不会引发错误。 以下示例错误地尝试将只读属性设置为 [Chart.id](/javascript/api/excel/excel.chart#id)。另请参阅 [某些属性不能直接设置](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>收到错误："此外接程序不再可用"

以下是导致此错误的一些原因。 如果发现其他原因，请使用页面底部的反馈工具告诉我们。

- 如果使用 Visual Studio，则旁加载可能有问题。 关闭主机和Office的所有Visual Studio。 重新启动Visual Studio并再次尝试按 F5。
- 外接程序的清单已从其部署位置（如集中部署、SharePoint目录或网络共享）中删除。
- 清单 [中 ID](../reference/manifest/id.md) 元素的值已在已部署的副本中直接更改。 如果出于任何原因需要更改此 ID，请首先从 Office 主机中删除外接程序，然后将原始清单替换为已更改的清单。 许多用户需要清除Office缓存以删除原始缓存的所有跟踪。 有关[为操作系统清除Office](clear-cache.md)的说明，请参阅清除缓存缓存文章。
- 加载项的清单有 一个 未在清单的"资源"部分的任何位置定义的 ，或者其使用位置和在部分中定义位置的拼写不匹配。 `resid` [](../reference/manifest/resources.md) `resid` `<Resources>`
- 清单 `resid` 中的某位置有一个超过 32 个字符的属性。 属性和节中相应资源的属性不能超过 `resid` `id` `<Resources>` 32 个字符。
- 加载项具有自定义加载项命令，但尝试在不支持命令的平台上运行。 有关详细信息，请参阅加载项 [命令要求集](../reference/requirement-sets/add-in-commands-requirement-sets.md)。

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>外接程序在 Edge 上不起作用，但它适用于其他浏览器

请参阅[疑难解答Microsoft Edge问题](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel加载项抛出错误，但不一致

请参阅[Excel加载项疑难](../excel/excel-add-ins-troubleshooting.md)解答了解可能的原因。

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>清单架构验证错误Visual Studio项目中

如果使用的是需要更改清单文件的较新功能，则可能会收到清单Visual Studio。 例如，添加 元素 `<Runtimes>` 来实现共享的 JavaScript 运行时时，你可能会看到以下验证错误。

**命名空间 中的元素"Host"' 在命名空间 ' 中具有无效 http://schemas.microsoft.com/office/taskpaneappversionoverrides 的子元素 http://schemas.microsoft.com/office/taskpaneappversionoverrides "Runtimes"**

如果发生这种情况，你可以将 XSD 文件更新Visual Studio最新版本。 最新架构版本位于 [[MS-OWEMXML]：附录 A：完整 XML 架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。

### <a name="locate-the-xsd-files"></a>找到 XSD 文件

1. 在 Visual Studio 中打开项目。
1. 在 **"解决方案资源管理器**"中，打开manifest.xml文件。 清单通常位于解决方案下的第一个项目中。
1. 选择 **"查看**  >  **属性窗口**" (F4) 。
1. 在" **属性窗口**"中，选择省略号" (...) "以打开 **XML 架构** 编辑器。 你可以在此处找到项目使用的所有架构文件的确切文件夹位置。

### <a name="update-the-xsd-files"></a>更新 XSD 文件

1. 在文本编辑器中打开要更新的 XSD 文件。 验证错误中的架构名称将关联到 XSD 文件名。 例如，打开 **TaskPaneAppVersionOverridesV1_0.xsd**。
1. 找到更新的架构 [，位置为 [MS-一级：XML 架构]：附录 A：完整 XML 架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。 例如，TaskPaneAppVersionOverridesV1_0位于 [taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)中。
1. 将文本复制到文本编辑器中。
1. 保存更新后的 XSD 文件。
1. 重新启动Visual Studio以选取新的 XSD 文件更改。

您可以对过期的其他任何架构重复上述过程。

## <a name="see-also"></a>另请参阅

- [在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)
- [将 Office 外接程序旁加载到 iPad 和 Mac 上](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [在 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)  
- [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)
- [验证 Office 加载项的清单](troubleshoot-manifest.md)
- [使用运行时日志记录功能调试加载项](runtime-logging.md)
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)
- [Microsoft Q&a (office-js-dev) ](/answers/topics/office-js-dev.html)
