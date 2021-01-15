---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 0bd231cc870322dd6f756defd14e4d67a69478b4
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771247"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office 外接程序是 Web 应用程序，在 Office 网页版中运行时，使用 iFrame 显示，在 Office 中对桌面客户端和移动客户端使用嵌入式浏览器控件。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入的浏览器和引擎都由用户计算机上安装的浏览器提供。

要使用的浏览器取决于：

- 计算机的操作系统。
- 加载项是在 Office 网页版、Microsoft 365 版还是非订阅版 Office 2013 或更高版本中运行。

下表显示在不同平台和操作系统中使用的浏览器。

|操作系统|Office 版本|Edge WebView2 (基于 Chromium) 安装？|浏览器|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|任意|Office 网页版|不适用|在其中打开 Office 的浏览器。|
|Mac|任意|不适用|Safari|
|iOS|任意|不适用|Safari|
|Android|任意|不适用|Chrome|
|Windows 7、8.1、10 | 非订阅 Office 2013 或更高版本|无关紧要|Internet Explorer 11|
|Windows 7 | Microsoft 365| 无关紧要 | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 版本。 &nbsp; < &nbsp;1903| Microsoft 365 | 否| Internet Explorer 11|
|Windows 10 版本。 &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; < &nbsp;16.0.11629<sup>1</sup>| 无关紧要|Internet Explorer 11|
|Windows 10 版本。 &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; >= &nbsp;16.0.11629 &nbsp; _和_ &nbsp; < &nbsp; 16.0.13530.20316 <sup>1</sup>| 无关紧要|具有原始 WebView 的 Microsoft Edge<sup>2、3</sup> (EdgeHTML) |
|Windows 10 版本。 &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; >= &nbsp;16.0.13530.20316<sup>1</sup>| 否 |具有原始 WebView 的 Microsoft Edge<sup>2、3</sup> (EdgeHTML) |
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp;16.0.13530.20316<sup>1</sup>| 是<sup>4</sup>|  Microsoft Edge<sup>2、3</sup> 和 WebView2 (基于 Chromium)  |

<sup>1</sup> 请参阅 [更新历史记录页](/officeupdates/update-history-office365-proplus-by-date) 以及如何查找 [Office 客户端版本](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) 和更新通道，了解更多详细信息。

<sup>2</sup> 使用 Microsoft Edge 时，Windows 10 讲述人 (有时称为"屏幕阅读器") 读取任务窗格中打开的页面 `<title>` 中的标记。 如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。

<sup>3</sup> 如果你的外接程序在清单中包含元素，它将使用 Internet Explorer 11，而不考虑 Windows 或 `Runtimes` Microsoft 365 版本。 有关详细信息，请参阅[运行时](../reference/manifest/runtimes.md)。

<sup>4</sup> 除了安装 Microsoft Edge 之外，还必须安装可嵌入的 WebView2 控件，以便 Office 可以嵌入它。 若要安装它，请参阅[Microsoft Edge WebView2/嵌入 Web 内容...使用 Microsoft Edge WebView2。](https://developer.microsoft.com/microsoft-edge/webview2/)


> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果任何加载项用户具有使用 Internet Explorer 11 的平台，那么要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项：
>
> - 在 ECMAScript 2015 (也称为 ES6) 或更高版本的 JavaScript 中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0 或](https://babeljs.io/) [tsc）](https://www.typescriptlang.org/index.html)将代码编译为 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript[](https://wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js）](https://github.com/zloirock/core-js)以允许 IE 运行代码。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge 问题疑难解答

### <a name="service-workers-are-not-working"></a>服务工作人员未工作

使用原始 [Microsoft Edge WebView](/microsoft-edge/hosting/webview) 时，Office 外接程序不支持服务工作人员。 它们受基于 [Chromium 的边缘 WebView2 的支持](/microsoft-edge/hosting/webview2)。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>任务窗格中不显示滚动条

默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。 适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) 属性，且应将其设置为 `scrollbar`。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载

[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。 发生这种情况时，它将自动重新加载该加载项。 为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误

一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。 按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>尝试下载 PDF 文件时出错

当 Edge 为浏览器时，不支持在外接程序中将 blob 直接下载为 PDF 文件。 解决方法是创建一个简单的 Web 应用程序，将 blob 下载为 PDF 文件。 在外接程序中，调用 `Office.context.ui.openBrowserWindow(url)` 该方法并传递 Web 应用程序的 URL。 这将在 Office 外部的浏览器窗口中打开 Web 应用程序。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
