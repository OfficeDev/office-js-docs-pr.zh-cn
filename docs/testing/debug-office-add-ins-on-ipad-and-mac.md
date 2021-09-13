---
title: 在 Mac 上调试 Office 加载项
description: 了解如何使用 Mac 调试Office加载项。
ms.date: 10/16/2020
ms.localizationpriority: medium
ms.openlocfilehash: 46104e5cbd9c81e56c1a83b6f49ae5883097b3e5
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149511"
---
# <a name="debug-office-add-ins-on-a-mac"></a>在 Mac 上调试 Office 加载项

由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>在 Mac 上使用 Safari Web 检查器进行调试

如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。

若要能够在 Mac Office加载项，你必须拥有 Mac OS High Sierra 和 Mac Office 版本 16.9.1 (版本18012504) 或更高版本。 如果你没有 Mac Office，可以通过加入开发人员计划获取一Microsoft 365[版本](https://developer.microsoft.com/office/dev-program)。

首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > Mac 应用商店版本Office不支持 `OfficeWebAddinDeveloperExtras` 标志。

然后，打开 Office 应用程序并[旁加载你的加载项](sideload-an-office-add-in-on-ipad-and-mac.md)。 右键单击加载项，应在上下文菜单中看到一个“**检查元素**”选项。 选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。

> [!NOTE]
> 如果你尝试使用检查器时对话框闪烁，请将 Office 更新到最新版本。 如果无法解决闪烁问题，请尝试以下解决方法。
>
> 1. 缩小对话框大小。
> 1. 选择“检查元素”，这将在新窗口中打开。
> 1. 将对话框调整为原始大小。
> 1. 根据需要使用检查器。

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>在 Mac 上清除 Office 应用程序的缓存

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
