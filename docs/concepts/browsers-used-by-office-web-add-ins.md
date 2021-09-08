---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 08/09/2021
localization_priority: Normal
ms.openlocfilehash: bda86e8bb7aacf72fbe26e86b7f062f362adbdd3
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937706"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office外接程序是 Web 应用程序，在 Office web 版 中运行时，使用 iFrame 显示，Office和移动客户端使用嵌入式浏览器控件。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入的浏览器和引擎都由用户计算机上安装的浏览器提供。

要使用的浏览器取决于：

- 计算机的操作系统。
- 外接程序是在 2013 Office web 版、Microsoft 365或非订阅Office中运行。

> [!IMPORTANT]
> **Internet Explorer外接程序Office中使用的内容**
>
> Microsoft 将终止对Internet Explorer的支持，但这不会显著Office外接程序。平台和 Office 版本（包括 Office 2019 的所有一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如本文所说明。 此外，提交到 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center)的加载项仍然需要支持这些组合，因此Internet Explorer对应用的支持。 有两 *个变化* ：
>
> - AppSource 不再使用作为浏览器Office web 版Internet Explorer加载项。 但 AppSource 仍测试使用 Office *版本的平台* 和桌面Internet Explorer。
> - Script Lab[工具](../overview/explore-with-script-lab.md)不再支持Internet Explorer。

下表显示在不同平台和操作系统中使用的浏览器。

|操作系统|Office 版本|安装了基于 (Chromium WebView2) Edge WebView2？|浏览器|
|:-----|:-----|:-----|:-----|
|任意|Office 网页版|不适用|在其中打开 Office 的浏览器。|
|Mac|任意|不适用|Safari|
|iOS|任意|不适用|Safari|
|Android|任意|不适用|Chrome|
|Windows 7、8.1、10 | 2013 Office或更高版本的非订阅|无关紧要|Internet Explorer 11|
|Windows 7 | Microsoft 365| 无关紧要 | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 ver. &nbsp; < &nbsp;1903| Microsoft 365 | 否| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; < &nbsp;16.0.11629<sup>1</sup>| 无关紧要|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; >= &nbsp;16.0.11629 &nbsp; _和_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| 无关紧要|Microsoft Edge<sup>2、3，</sup>具有原始 WebView (EdgeHTML) |
|Windows 10 ver. &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; >= &nbsp;16.0.13530.20424<sup>1</sup>| 否 |Microsoft Edge<sup>2、3，</sup>具有原始 WebView (EdgeHTML) |
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp;16.0.13530.20424<sup>1</sup>| 是<sup>4</sup>|  Microsoft Edge<sup>2</sup>与 WebView2 (Chromium基于)  |

<sup>1</sup>有关更多详细信息[，请参阅更新历史记录页](/officeupdates/update-history-office365-proplus-by-date)Office[客户端版本和更新](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)通道。

<sup>2</sup> Microsoft Edge时，Windows 10讲述人 (有时称为"屏幕阅读器") 读取任务窗格中打开的页面 `<title>` 中的标记。 如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。

<sup>3</sup>如果加载项在清单中包含 元素，则它将不会将 Microsoft Edge与原始 WebView (`<Runtimes>` EdgeHTML) 。 如果满足将 webView2 Microsoft Edge WebView2 (Chromium的条件) ，则外接程序会使用该浏览器。 否则，它将使用 Internet Explorer 11，而不考虑Windows或Microsoft 365版本。 有关详细信息，请参阅[运行时](../reference/manifest/runtimes.md)。

<sup>4</sup>必须安装可嵌入的 WebView2 控件Office嵌入它，并且它不会自动随 Edge 一起安装。 它随 Microsoft 365 2101 或更高版本一起安装。 如果你拥有早期版本的 Microsoft 365，请按照在 WebView2/嵌入 web Microsoft Edge安装[控件的说明...使用 Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)。

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果任何外接程序用户具有使用 Internet Explorer 11 的平台，则要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项。
>
> - 在 ECMAScript 2015 (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0](https://babeljs.io/) 或 [tsc）](https://www.typescriptlang.org/index.html)将代码编译为 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript[](https://en.wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js，](https://github.com/zloirock/core-js)它使 IE 能够运行代码）。
>
> 有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。

## <a name="troubleshooting-microsoft-edge-issues"></a>疑难Microsoft Edge疑难解答

### <a name="service-workers-are-not-working"></a>服务工作人员未工作

Office使用原始 WebView Microsoft Edge [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML)时，外接程序不支持服务工作人员。 它们受基于 Chromium [Edge WebView2 的支持](/microsoft-edge/hosting/webview2)。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>任务窗格中不显示滚动条

默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。 适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) 属性，且应将其设置为 `scrollbar`。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载

[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。 发生这种情况时，它将自动重新加载该加载项。 为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误

一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。 按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>尝试下载 PDF 文件时出错

当 Edge 为浏览器时，不支持在外接程序中直接将 blob 下载为 PDF 文件。 解决方法是创建一个简单的 Web 应用程序，将 blob 下载为 PDF 文件。 在外接程序中，调用 `Office.context.ui.openBrowserWindow(url)` 方法并传递 Web 应用程序的 URL。 这将在 Web 应用程序外部的浏览器窗口中Office。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
