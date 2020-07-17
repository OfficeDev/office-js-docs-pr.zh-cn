---
title: 在 Mac 上调试 Office 加载项
description: 了解如何使用 Mac 调试 Office 外接程序
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 12785a195c336e0de8c619379a3839bd15079b2c
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094125"
---
# <a name="debug-office-add-ins-on-a-mac"></a>在 Mac 上调试 Office 加载项

由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>在 Mac 上使用 Safari Web 检查器进行调试

如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。

要在 Mac 上调试 Office 加载项，必须拥有 Mac OS High Sierra 和 Mac Office 版本：16.9.1（内部版本 18012504）或更高版本。 如果没有 Office Mac 版本，则可以通过加入[Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个。

首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

然后，打开 Office 应用程序并[旁加载你的加载项](sideload-an-office-add-in-on-ipad-and-mac.md)。 右键单击加载项，应在上下文菜单中看到一个“**检查元素**”选项。 选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。

> [!NOTE]
> 如果你尝试使用检查器时对话框闪烁，请将 Office 更新到最新版本。 如果这样做未解决闪烁问题，请尝试以下解决方法：
> 1. 缩小对话框大小。
> 2. 选择“检查元素”****，这将在新窗口中打开。
> 3. 将对话框调整为原始大小。
> 4. 根据需要使用检查器。

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>在 Mac 上清除 Office 应用程序的缓存

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
