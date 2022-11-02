---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 09/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: a75cab613605760e774f8b2a163172e4ec6cb5bd
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810153"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office 外接程序是在 Office web 版 中运行时使用 iFrame 显示的 Web 应用程序。 在 Office 桌面和移动客户端中，Office 外接程序使用嵌入式浏览器控件 (也称为 Web 视图) 。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入式浏览器和引擎均由安装在用户计算机上的浏览器提供。

要使用的浏览器取决于：

- 计算机的操作系统。
- 加载项是在 Office web 版、从 Microsoft 365 订阅下载的 Office 中运行的，还是在永久 Office 2013 或更高版本中运行。
- 在 Windows 上的 Office 永久版本中，加载项是在“零售”还是“批量许可”变体中运行。

> [!NOTE]
> 本文假定加载项在 *不受* [Windows 信息保护保护 (WIP)](/windows/uwp/enterprise/wip-hub)保护的文档中运行。 对于受 WIP 保护的文档，本文中的信息有一些例外。 有关详细信息，请参阅 [WIP 保护的文档](#wip-protected-documents)。

> [!IMPORTANT]
> **Internet Explorer 仍在 Office 加载项中使用**
>
> 平台和 Office 版本的一些组合（包括通过 Office 2019 的批量许可永久版本）仍使用 Internet Explorer 11 附带的 Webview 控件来托管加载项，如本文所述。 我们建议 (但不需要) ，至少在 Internet Explorer Web 视图中启动加载项时，通过向外接程序的用户提供正常失败消息，继续支持这些组合。 请记住以下附加要点：
>
> - Office web 版不再在 Internet Explorer 中打开。 因此，[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 不再使用 Internet Explorer 作为浏览器在 Office web 版 中测试加载项。
> - AppSource 仍会测试使用 Internet Explorer 的平台和 Office *桌面* 版本的组合。 但是，仅当加载项不支持 Internet Explorer 时，它才会发出警告;AppSource 不会拒绝加载项。
> - [Script Lab工具](../overview/explore-with-script-lab.md)不再支持 Internet Explorer。
>
> 有关在外接程序上支持 Internet Explorer 和配置正常失败消息的详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md)。

以下部分指定用于各种平台和操作系统的浏览器。

## <a name="non-windows-platforms"></a>非 Windows 平台

对于这些平台，平台将单独确定所使用的浏览器。

|操作系统|Office 版本|浏览器|
|:-----|:-----|:-----|
|任意|Office 网页版|在其中打开 Office 的浏览器。<br> (但请注意，Office web 版不会在 Internet Explorer 中打开。<br>尝试执行此操作会在 Edge.) 中打开Office web 版 |
|Mac|任意|将 Safari 与 WKWebView 配合使用|
|iOS|任意|将 Safari 与 WKWebView 配合使用|
|Android|任意|Chrome|

## <a name="perpetual-versions-of-office-on-windows"></a>Windows 上的 Office 永久版本

对于 Windows 上的 Office 永久版本，使用的浏览器由 Office 版本、许可证是零售许可证还是批量许可，以及是否安装了 Edge WebView2 基于 (Chromium) 。 Windows 版本并不重要，但请注意，Windows 7 之前的版本不支持 Office Web 外接程序，Office 2021在早于 Windows 10 的版本上不受支持。

若要确定 Office 2016 或 Office 2019 是零售许可还是批量许可，请使用 Office 版本和内部版本号的格式。  (对于 Office 2013 和 Office 2021，批量许可和零售之间的区别并不重要。) 

- **零售**：对于 Office 2016 和 2019，格式为 `YYMM (xxxxx.xxxxxx)`，以两个五位数字块结尾;例如 。 `2206 (Build 15330.20264`
- **批量许可**：
  - 对于 Office 2016，格式为 `16.0.xxxx.xxxxx`，以两个 *四* 位数字块结尾;例如 。 `16.0.5197.1000`
  - 对于 Office 2019，格式为 `1808 (xxxxx.xxxxxx)`，以两个 *五* 位数字块结尾;例如 。 `1808 (Build 10388.20027)` 请注意，年份和月份始终 `1808`为 。

| Office 版本 | 零售与批量许可 | 已安装基于 Edge WebView2 (Chromium 的) ？ | 浏览器 |
|:-----|:-----|:-----|:-----|
| Office 2013 | 无所谓 | 无所谓 | Internet Explorer 11 |
| Office 2016 | 批量许可 | 无所谓 | Internet Explorer 11 |
| Office 2019 | 批量许可 | 无所谓 | Internet Explorer 11 |
| Office 2016 到 Office 2019 | 零售版 | 否 | Microsoft Edge<sup>1，2</sup> 与原始 WebView (EdgeHTML) </br>如果未安装 Edge，则使用 Internet Explorer 11。 |
| Office 2016 到 Office 2019 | 零售版 | 是<sup>3</sup> | 具有基于 WebView2 (Chromium 的 Microsoft Edge<sup>1</sup>)  |
| Office 2021 | 无所谓 | 是<sup>3</sup> | 具有基于 WebView2 (Chromium 的 Microsoft Edge<sup>1</sup>)  |

<sup>1</sup> 使用 Microsoft Edge 时，Windows 讲述人 (有时称为“屏幕阅读器”) 读取 `<title>` 任务窗格中打开的页面中的标记。 在 Internet Explorer 11 中，讲述人读取任务窗格的标题栏，该标题栏来自 **\<DisplayName\>** 外接程序清单中的 值。

<sup>2</sup> 如果外接程序在 **\<Runtimes\>** 清单中包含 元素，则不会将 Microsoft Edge 与原始 WebView (EdgeHTML) 一起使用。 如果满足将 Microsoft Edge 与基于 WebView2 (Chromium) 配合使用的条件，则外接程序使用该浏览器。 否则，它将使用 Internet Explorer 11。 有关详细信息，请参阅[运行时](/javascript/api/manifest/runtimes)。

<sup>3</sup> 在 Windows 11 之前的 Windows 版本中，必须安装 WebView2 控件，以便 Office 可以嵌入它。 它随永久Office 2021或更高版本一起安装;但不会随 Microsoft Edge 一起自动安装。 如果你有早期版本的永久 Office，请使用有关在 [Microsoft Edge WebView2/嵌入 Web 内容中安装控件的说明。使用 Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)。

## <a name="microsoft-365-subscription-versions-of-office-on-windows"></a>Windows 上的 Office 的 Microsoft 365 订阅版本

对于 Windows 上的订阅 Office，使用的浏览器由操作系统、Office 版本以及是否安装 Edge WebView2 (Chromium) 决定。

|操作系统|Office 版本|已安装基于 Edge WebView2 (Chromium 的) ？|浏览器|
|:-----|:-----|:-----|:-----|
|Windows 7 | Microsoft 365| 无所谓 | Internet Explorer 11|
|Windows 8.1，<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | 否| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 版本&nbsp;<&nbsp;16.0.11629<sup>2</sup>| 无所谓|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 版本&nbsp;>=&nbsp;16.0.11629&nbsp;_和_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| 无所谓|Microsoft Edge<sup>1，3</sup> 与原始 WebView (EdgeHTML) |
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>窗口 11 | Microsoft 365 版本&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| 否 |Microsoft Edge<sup>1，3</sup> 与原始 WebView (EdgeHTML) |
|Windows 8.1<br>Windows 10，<br>Windows 11| Microsoft 365 版本&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| 是<sup>4</sup>|  具有基于 WebView2 (Chromium 的 Microsoft Edge<sup>1</sup>)  |

<sup>1</sup> 使用 Microsoft Edge 时，Windows 讲述人 (有时称为“屏幕阅读器”) 读取 `<title>` 任务窗格中打开的页面中的标记。 在 Internet Explorer 11 中，讲述人读取任务窗格的标题栏，该标题栏来自 **\<DisplayName\>** 外接程序清单中的 值。

<sup>2</sup> 有关更多详细信息，请参阅 [更新历史记录页](/officeupdates/update-history-office365-proplus-by-date) 以及如何 [查找 Office 客户端版本和更新通道](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) 。

<sup>3</sup> 如果外接程序在 **\<Runtimes\>** 清单中包含 元素，则不会将 Microsoft Edge 与原始 WebView (EdgeHTML) 一起使用。 如果满足将 Microsoft Edge 与基于 WebView2 (Chromium) 配合使用的条件，则外接程序使用该浏览器。 否则，无论 Windows 或 Microsoft 365 版本如何，它都会使用 Internet Explorer 11。 有关详细信息，请参阅[运行时](/javascript/api/manifest/runtimes)。

<sup>4</sup> 在 Windows 11 之前的 Windows 版本中，必须安装 WebView2 控件，以便 Office 可以嵌入它。 它随 Microsoft 365 版本 2101 或更高版本一起安装，但不会随 Microsoft Edge 一起自动安装。 如果你有早期版本的 Microsoft 365，请使用在 [Microsoft Edge WebView2/嵌入 Web 内容中安装控件的说明。使用 Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)。 在 16.0.14326.xxxxx 之前的 Microsoft 365 版本上，还必须HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2 **创建注册表项，** 并将其值设置为 `dword:00000001`。

## <a name="working-with-internet-explorer"></a>使用 Internet Explorer

Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果外接程序的任何用户具有使用 Internet Explorer 11 的平台，则若要使用 ECMAScript 2015 或更高版本的语法和功能，则有两个选项。

- 在 ECMAScript 2015 (也称为 ES6) 或更高版本的 JavaScript 或 TypeScript 中编写代码，然后使用 [babel](https://babeljs.io/) 或 [tsc](https://www.typescriptlang.org/index.html) 等编译器将代码编译为 ES5 JavaScript。
- 使用 ECMAScript 2015 或更高版本的 JavaScript 编写，但也加载使 IE 能够运行代码的 [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) 库（如 [core-js](https://github.com/zloirock/core-js) ）。

有关这些选项的详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md)。

此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。 若要了解详细信息，请参阅 [确定在运行时加载项是否在 Internet Explorer 中运行](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)。

## <a name="troubleshoot-microsoft-edge-issues"></a>排查 Microsoft Edge 问题

### <a name="service-workers-are-not-working"></a>服务辅助角色不工作

使用原始 Microsoft Edge WebView [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML) 时，Office 加载项不支持服务辅助角色。 [基于 Chromium 的 Edge WebView2](/microsoft-edge/hosting/webview2) 支持它们。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>任务窗格中不显示滚动条

默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。 适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) 属性，且应将其设置为 `scrollbar`。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载

[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。 发生这种情况时，它将自动重新加载该加载项。 为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误

一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。 按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>尝试下载 PDF 文件时出现错误

当 Edge 是浏览器时，不支持将 Blob 作为 PDF 文件直接下载到外接程序中。 解决方法是创建一个简单的 Web 应用程序，该应用程序将 Blob 下载为 PDF 文件。 在外接程序中，调用 `Office.context.ui.openBrowserWindow(url)` 方法并传递 Web 应用程序的 URL。 这将在 Office 外部的浏览器窗口中打开 Web 应用程序。

## <a name="wip-protected-documents"></a>WIP 保护的文档

在 [受 WIP 保护](/windows/uwp/enterprise/wip-hub)的文档中运行的外接程序永远不会 **将 Microsoft Edge 与基于 WebView2 (Chromium 的)** 配合使用。 在本文前面的 [Windows 版 Office 永久版本](#perpetual-versions-of-office-on-windows)和 [Windows 上的 Office 的 Microsoft 365 订阅版本](#microsoft-365-subscription-versions-of-office-on-windows)部分中，将 **Microsoft Edge 替换为原始 WebView (EdgeHTML)**，**将 Microsoft Edge 与基于 WebView2 (Chromium 的)**（无论后者出现在何处）。

若要确定文档是否受 WIP 保护，请执行以下步骤：

1. 打开此文件。
1. 选择功能区上的“ **文件** ”选项卡。
1. 选择“ **信息**”。
1. 在 **“信息** ”页面左上方的文件名正下方，启用 WIP 的文档将具有公文包图标，后跟 **由 Work (...)**。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
- [Office 加载项中的运行时](../testing/runtimes.md)
