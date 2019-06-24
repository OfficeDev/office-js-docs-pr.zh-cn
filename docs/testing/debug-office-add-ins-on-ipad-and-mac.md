---
title: 在 Mac 上调试 Office 加载项
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 88f7cbf6c944a0f6510306cfe2d07db59e40bdeb
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126930"
---
# <a name="debug-office-add-ins-on-a-mac"></a>在 Mac 上调试 Office 加载项

你可以使用 Visual Studio 在 Windows 上开发和调试加载项，但不能使用它在 Mac 上调试加载项。由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>在 Mac 上使用 Safari Web 检查器进行调试

如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。

要在 Mac 上调试 Office 加载项，必须拥有 Mac OS High Sierra 和 Mac Office 版本：16.9.1（内部版本 18012504）或更高版本。 如果没有 Office Mac 内部版本，可以通过加入 [Office 365 开发人员计划](https://aka.ms/o365devprogram)获取一个版本。

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

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>在 Mac 上清除 Office 应用程序缓存

出于性能方面的考虑，加载项通常在 Mac 版 Office 中缓存。通常情况下，将通过重载加载项清除缓存。如果同一文档中存在多个加载项，则重载后自动清除缓存的过程可能不可靠。

在 Mac 上，通过删除 `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的内容可以手动清除缓存。 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
