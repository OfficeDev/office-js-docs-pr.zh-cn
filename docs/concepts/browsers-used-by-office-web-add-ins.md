---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1fedeb7f7e1e972a2a7fe4befa5a990ff8cc698d
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659652"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office 加载项是在Office web 版中运行时使用 iFrame 显示的 Web 应用程序。 在 Office 桌面和移动客户端中，Office 外接程序使用嵌入式浏览器控件 (也称为 Web 视图) 。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入式浏览器和引擎均由用户计算机上安装的浏览器提供。

要使用的浏览器取决于：

- 计算机的操作系统。
- 加载项是在 Office web 版、Microsoft 365 还是非订阅 Office 2013 或更高版本中运行。

> [!IMPORTANT]
> **仍在 Office 加载项中使用的 Internet Explorer**
>
> 平台和 Office 版本的某些组合（包括通过 Office 2019 的一次性购买版本）仍使用 Internet Explorer 11 附带的 Webview 控件来托管加载项，如本文所述。 建议 (但不需要) 继续支持这些组合（至少以最小方式）在 Internet Explorer Webview 中启动外接程序时为外接程序的用户提供正常故障消息。 请记住以下附加点：
>
> - Office web 版不再在 Internet Explorer 中打开。 因此，[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 不再使用 Internet Explorer 作为浏览器在Office web 版中测试加载项。
> - AppSource 仍会测试使用 Internet Explorer 的平台和 Office *桌面* 版本的组合，但仅当外接程序不支持 Internet Explorer 时才会发出警告;AppSource 不会拒绝加载项。
> - [Script Lab工具](../overview/explore-with-script-lab.md)不再支持 Internet Explorer。
>
> 有关在外接程序上支持 Internet Explorer 和配置正常故障消息的详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md)。

下表显示在不同平台和操作系统中使用的浏览器。

|操作系统|Office 版本|已安装基于 Edge WebView2 (Chromium) ？|浏览器|
|:-----|:-----|:-----|:-----|
|任意|Office 网页版|不适用|在其中打开 Office 的浏览器。<br> (但请注意，Office web 版不会在 Internet Explorer 中打开。<br>尝试这样做将在 Edge.) 中打开Office web 版 |
|Mac|任意|不适用|带 WKWebView 的 Safari|
|iOS|任意|不适用|带 WKWebView 的 Safari|
|Android|任意|不适用|Chrome|
|Windows 7、8.1、10、11 | 非订阅 Office 2013 到 Office 2019|无所谓|Internet Explorer 11|
|Windows 10、11 | 非订阅Office 2021或更高版本|是|Microsoft Edge<sup>1</sup> 和基于 WebView2 (Chromium的) |
|Windows 7 | Microsoft 365| 无所谓 | Internet Explorer 11|
|Windows 8.1，<br>Windows 10 ver。&nbsp;<&nbsp;1903| Microsoft 365 | 不支持| Internet Explorer 11|
|Windows 10 ver。&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.11629<sup>2</sup>| 无所谓|Internet Explorer 11|
|Windows 10 ver。&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;_和_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| 无所谓|Microsoft Edge<sup>1、3 和</sup> 原始 WebView (EdgeHTML) |
|Windows 10 ver。&nbsp;>=&nbsp;1903,<br>窗口 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| 不支持 |Microsoft Edge<sup>1、3 和</sup> 原始 WebView (EdgeHTML) |
|Windows 8.1<br>Windows 10、<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| 是<sup>4</sup>|  Microsoft Edge<sup>1</sup> 和基于 WebView2 (Chromium的)  |

<sup>1</sup> 使用 Microsoft Edge 时，Windows 讲述人 (有时称为“屏幕阅读器”，) 在任务窗格中打开的页面中读取 `<title>` 标记。 使用 Internet Explorer 11 时，讲述人将读取任务窗格的标题栏，该标题栏来自 **\<DisplayName\>** 加载项清单中的值。

<sup>2</sup> 有关更多详细信息，请参阅 [“更新历史记录”页](/officeupdates/update-history-office365-proplus-by-date) 以及如何 [查找 Office 客户端版本和更新通道](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) 。

<sup>3</sup> 如果加载项包含 **\<Runtimes\>** 清单中的元素，则它不会将 Microsoft Edge 与原始 WebView (EdgeHTML) 配合使用。 如果满足将 Microsoft Edge 与基于 WebView2 (Chromium) 使用的条件，加载项将使用该浏览器。 否则，无论 Windows 或 Microsoft 365 版本如何，它都使用 Internet Explorer 11。 有关详细信息，请参阅[运行时](/javascript/api/manifest/runtimes)。

<sup>4</sup> 在Windows 11之前的 Windows 版本上，必须安装 WebView2 控件，以便 Office 可以嵌入它。 它随 Microsoft 365 版本 2101 或更高版本一起安装，并且一次性购买Office 2021或更高版本;但它不会自动安装在 Microsoft Edge 中。 如果你有早期版本的 Microsoft 365 或一次性购买 Office，请使用有关在 [Microsoft Edge WebView2/Embed Web 内容上安装控件的说明...使用 Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)。 在 16.0.14326.xxxxx 之前的 Microsoft 365 版本上，还必须创建注册表项 **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** 并将其值设置为 `dword:00000001`。

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果加载项的任何用户都有使用 Internet Explorer 11 的平台，则要使用 ECMAScript 2015 或更高版本的语法和功能，可以使用两个选项。
>
> - 在 ECMAScript 2015 (也称为 ES6) 或更高版本 JavaScript 或 TypeScript 中编写代码，然后使用编译器（如 [babel](https://babeljs.io/) 或 [tsc](https://www.typescriptlang.org/index.html)）将代码编译到 ES5 JavaScript。
> - 在 ECMAScript 2015 或更高版本的 JavaScript 中编写，但还要加载 [一个多填充](https://en.wikipedia.org/wiki/Polyfill_(programming)) 库（如 [core-js](https://github.com/zloirock/core-js) ），使 IE 能够运行代码。
>
> 有关这些选项的详细信息，请参阅 [支持 Internet Explorer 11](../develop/support-ie-11.md)。
>
> 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。 若要了解详细信息，请参阅 [运行时确定外接程序是否在 Internet Explorer 中运行](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge 问题疑难解答

### <a name="service-workers-are-not-working"></a>服务工作者不起作用

使用原始 Microsoft Edge WebView Edge WebView [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML) 时，Office 加载项不支持服务工作者。 [基于 Chromium 的 Edge WebView2](/microsoft-edge/hosting/webview2) 支持这些功能。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>任务窗格中不显示滚动条

默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。 适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) 属性，且应将其设置为 `scrollbar`。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载

[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。 发生这种情况时，它将自动重新加载该加载项。 为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误

一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。 按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>获取尝试下载 PDF 文件的错误

当 Edge 是浏览器时，不支持将 Blob 直接下载为加载项中的 PDF 文件。 解决方法是创建一个简单的 Web 应用程序，将 Blob 下载为 PDF 文件。 在外接程序中，调用该 `Office.context.ui.openBrowserWindow(url)` 方法并传递 Web 应用程序的 URL。 这将在 Office 外部的浏览器窗口中打开 Web 应用程序。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
