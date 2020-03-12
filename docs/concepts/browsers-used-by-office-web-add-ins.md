---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 03/09/2020
localization_priority: Normal
ms.openlocfilehash: d53ea0da29c9d2cc1177d233eed9e3ee62a891f2
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596464"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office 加载项是使用 iFrames（在 Office 网页版中运行时）和使用 Office 桌面版和移动版客户端中的嵌入式浏览器控件显示的 Web 应用程序。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入的浏览器和引擎都是由安装在用户计算机上的浏览器提供的。

要使用的浏览器取决于：

- 计算机的操作系统。
- 加载项是在 Office 网页版、Office 365 还是非订阅版 Office 2013 或更高版本中运行。

下表显示在不同平台和操作系统中使用的浏览器。

|**操作系统/平台**|**Browser**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office 网页版|在其中打开 Office 的浏览器。|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows/非订阅版 Office 2013 或更高版本|Internet Explorer 11|
|Windows 10 版本 < 1903 / Office 365|Internet Explorer 11|
|Windows 10 版本 >= 1903 / Office 365 ver < 16.0.11629|Internet Explorer 11|
|Windows 10 版本 >= 1903 / Office 365 ver >= 16.0.11629|Microsoft Edge\*|

\*使用 Microsoft Edge 时，Windows 10 讲述人（有时称为“屏幕阅读器”）会读出页面中在任务窗格中打开的 `<title>` 标记。 如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。

> [!IMPORTANT]
> Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。 如果任何加载项用户安装的是使用 Internet Explorer 11 的平台，若要使用 ECMAScript 2015 或更高版本的语法和功能，则必须将 JavaScript 转换为 ES5 或使用填充代码。 此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge 问题疑难解答

### <a name="service-workers-are-not-working"></a>服务工作人员不工作

Office 外接程序不支持[Microsoft Edge web](/microsoft-edge/hosting/webview)上的服务工作线程。 请参阅[Office 外接程序概述](../overview/office-add-ins.md)，了解有关边缘 web 视图控件的最新支持的功能。 我们正在努力将基于 Chromium 的新[边缘 WebView2](/microsoft-edge/hosting/webview2)带到 Office 外接程序平台，我们预期将支持服务工作人员。

### <a name="chromium-based-edge-is-installed-on-my-development-computer-but-my-add-in-does-not-use-it"></a>在我的开发计算机上安装了基于 Chromium 的边缘，但我的加载项不使用它

[Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium)中的基本浏览器已更改为 Chromium。 在安装基于 Chromium 的边缘时，不会删除较早的 base （称为 "EdgeHTML"）。 Office 仍将使用加载项的 EdgeHTML 基础，直到在计算机上安装了支持 Chromium 的 Office 365 版本。 我们预计这些版本将在2020中发货。 它们可能会在预览体验频道中的年上半年显示。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>任务窗格中不显示滚动条

默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。 适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) 属性，且应将其设置为 `scrollbar`。 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载

[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。 发生这种情况时，它将自动重新加载该加载项。 为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误

一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。 按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。


## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
