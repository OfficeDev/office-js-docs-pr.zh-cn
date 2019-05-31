---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 05/28/2019
localization_priority: Priority
ms.openlocfilehash: 92218bb012ae9031ebfc429606885a0ec0ea85b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/30/2019
ms.locfileid: "34592127"
---
# <a name="browsers-used-by-office-add-ins"></a>Office 加载项使用的浏览器

Office 加载项是使用 iFrames（在 Office Online 中运行时）和使用 Office 桌面版和移动版客户端中的嵌入式浏览器控件显示的 Web 应用程序。 加载项还需要使用 JavaScript 引擎来运行 JavaScript。 嵌入式浏览器和引擎均由用户计算机上安装的浏览器提供。

要使用的浏览器取决于：

- 计算机的操作系统。
- 加载项是在 Office Online、Office 365 还是非订阅版 Office 2013 或更高版本中运行。

下表显示在不同平台和操作系统中使用的浏览器。

|**操作系统/平台**|**浏览器**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office Online|在其中打开 Office Online 的浏览器。|
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

> [!NOTE]
> 在它们公开发布之前，你需要是 Windows 预览体验成员才能获得 Windows 版本 1903 或更高版本，并且需要是 Office 预览体验成员才能获得 Office 版本 16.0.11629 或更高版本。
>
> 若要加入 Windows 预览体验成员：
> 
> 1. 转至 [Windows 预览体验成员](https://insider.windows.com)并单击链接以加入 Windows 预览体验成员。
> 2. 系统会将你引导至包含有关如何使用 Windows 设置支持 Windows 预览内部版本说明的页面。 按照说明执行操作。 选择更新频率时，请选择时间最短的选项。
>
> 若要加入 Office 预览体验成员：
> 
> 1. 转至 [Office 预览体验成员入门](https://insider.office.com/join)。
> 2. 按照加入页面上的说明操作。 系统要求你指定频道时，请选择预览体验成员。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](requirements-for-running-office-add-ins.md)
