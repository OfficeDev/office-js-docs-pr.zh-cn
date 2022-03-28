---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: a6dd2eceb320b9f88575c80f1f4a17becc06cbe5
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483368"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office外接程序是 Web 应用程序，当在外接程序中运行时，它们使用 iFrame Office web 版。 在Office客户端和移动客户端中，Office外接程序使用嵌入式浏览器控件 (也称为 webview) 。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入的浏览器和引擎都由用户计算机上安装的浏览器提供。

要使用的浏览器取决于：

- 计算机的操作系统。
- 外接程序是在 2013 Office web 版、Microsoft 365或非订阅Office中运行。

> [!IMPORTANT]
> **Internet Explorer加载项中Office仍使用**
>
> Microsoft 将终止对Internet Explorer的支持，但这不会显著Office外接程序。平台和 Office 版本（包括 Office 2019 之间的一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如本文所说明。 此外，提交到 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 的外接程序仍然需要支持这些组合Internet Explorer因此也支持这些组合。 有两 *个变化* ：
>
> - Office web 版中不再打开Internet Explorer。 因此，AppSource 不再使用 Office web 版 浏览器Internet Explorer测试加载项。 但 AppSource 仍测试使用 *Office 版本的平台* 和桌面Internet Explorer。
> - Script Lab[工具](../overview/explore-with-script-lab.md)不再支持Internet Explorer。

下表显示在不同平台和操作系统中使用的浏览器。

|操作系统|Office 版本|安装了基于 (Chromium WebView2) Edge WebView2？|浏览器|
|:-----|:-----|:-----|:-----|
|任意|Office 网页版|不适用|在其中打开 Office 的浏览器。<br> (但请注意，Office web 版将不会在 Internet Explorer 中打开。<br>尝试这样做将在 Edge.Office web 版中打开)  |
|Mac|任意|不适用|Safari|
|iOS|任意|不适用|Safari|
|Android|任意|不适用|Chrome|
|Windows 7、8.1、10、11 | 从 2013 Office 2019 Office非订阅|无关紧要|Internet Explorer 11|
|Windows 10、11 | 2021 Office更高版本的非订阅|是|Microsoft Edge <sup>1</sup> 与 WebView2 (Chromium基于) |
|Windows 7 | Microsoft 365| 无关紧要 | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | 否| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.116292<sup></sup>| 无关紧要|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629AND16.0.13530.204242&nbsp;&nbsp;<sup></sup><&nbsp;| 无关紧要|Microsoft Edge <sup>1，3 包含</sup>原始 WebView (EdgeHTML) |
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>窗口 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.204242<sup></sup>| 否 |Microsoft Edge <sup>1，3 包含</sup>原始 WebView (EdgeHTML) |
|Windows 8.1<br>Windows 10、<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.204242<sup></sup>| 是<sup>4</sup>|  Microsoft Edge <sup>1</sup> 与 WebView2 (Chromium基于)  |

<sup>1</sup> Microsoft Edge时，Windows 讲述人 (有时称为"屏幕阅读器") `<title>`读取任务窗格中打开的页面中的标记。 如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。

<sup>2</sup> 有关更多详细信息[，](/officeupdates/update-history-office365-proplus-by-date)请参阅更新历史记录页Office[客户端版本和更新](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)通道。

<sup>3</sup> 如果加载项在`<Runtimes>`清单中包含 元素，则它将不会将 Microsoft Edge与原始 WebView (EdgeHTML) 。 如果满足将 webView2 Microsoft Edge WebView2 (Chromium的条件) ，则外接程序会使用该浏览器。 否则，它将使用 Internet Explorer 11，而不考虑Windows或Microsoft 365版本。 有关详细信息，请参阅[运行时](/javascript/api/manifest/runtimes)。

<sup>4</sup> Windows之前的版本Windows 11，必须安装 WebView2 控件，以便Office嵌入它。 它随 Microsoft 365 版本 2101 或更高版本一起安装，并且具有一次购买 Office 2021 或更高版本;但它不会自动随 Microsoft Edge 一起安装。 如果你有早期版本的 Microsoft 365 或一次购买 Office，请按照说明在 [Microsoft Edge WebView2 / 嵌入 Web 内容 ...使用 Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)。 在Microsoft 365 16.0.14326.xxxxx 之前生成，还必须创建注册表项HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2，并设置其值。**** `dword:00000001`

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果任何外接程序的用户具有使用 Internet Explorer 11 的平台，那么要使用 ECMAScript 2015 或更高版本的语法和功能，你有两个选项。
>
> - 在 ECMAScript 2015 (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#](https://babeljs.io/) A0 或 [tsc](https://www.typescriptlang.org/index.html)）将代码编译为 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript 中编写，但也加载填充[](https://en.wikipedia.org/wiki/Polyfill_(programming))库（如 [core-js](https://github.com/zloirock/core-js)，它使 IE 能够运行代码）。
>
> 有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。 若要了解更多信息，请参阅 [在运行时确定加载项](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)是否Internet Explorer。

## <a name="troubleshooting-microsoft-edge-issues"></a>疑难Microsoft Edge疑难解答

### <a name="service-workers-are-not-working"></a>服务工作人员未工作

Office WebView、[EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML) 和原始外接程序Microsoft Edge不支持服务工作人员。 它们受基于 Chromium [Edge WebView2 的支持](/microsoft-edge/hosting/webview2)。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>任务窗格中不显示滚动条

默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。 适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) 属性，且应将其设置为 `scrollbar`。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载

[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。 发生这种情况时，它将自动重新加载该加载项。 为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误

一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。 按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>尝试下载 PDF 文件时出错

当 Edge 为浏览器时，不支持在外接程序中直接将 blob 下载为 PDF 文件。 解决方法是创建一个简单的 Web 应用程序，将 blob 下载为 PDF 文件。 在外接程序中，调用 方法 `Office.context.ui.openBrowserWindow(url)` 并传递 Web 应用程序的 URL。 这将在 Web 应用程序外部的浏览器窗口中Office。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
