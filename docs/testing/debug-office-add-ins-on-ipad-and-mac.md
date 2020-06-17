---
title: 在 Mac 上调试 Office 加载项
description: 了解如何使用 Mac 调试 Office 外接程序
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: bcdf16a858fa613a3248a728b45d68d02cf41a5e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611217"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="1eeba-103">在 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="1eeba-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="1eeba-p101">由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。</span><span class="sxs-lookup"><span data-stu-id="1eeba-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="1eeba-106">在 Mac 上使用 Safari Web 检查器进行调试</span><span class="sxs-lookup"><span data-stu-id="1eeba-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="1eeba-107">如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="1eeba-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="1eeba-108">要在 Mac 上调试 Office 加载项，必须拥有 Mac OS High Sierra 和 Mac Office 版本：16.9.1（内部版本 18012504）或更高版本。</span><span class="sxs-lookup"><span data-stu-id="1eeba-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="1eeba-109">如果没有 Office Mac 内部版本，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个版本。</span><span class="sxs-lookup"><span data-stu-id="1eeba-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="1eeba-110">首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="1eeba-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="1eeba-111">然后，打开 Office 应用程序并[旁加载你的加载项](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="1eeba-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="1eeba-112">右键单击加载项，应在上下文菜单中看到一个“**检查元素**”选项。</span><span class="sxs-lookup"><span data-stu-id="1eeba-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="1eeba-113">选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。</span><span class="sxs-lookup"><span data-stu-id="1eeba-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1eeba-114">如果你尝试使用检查器时对话框闪烁，请将 Office 更新到最新版本。</span><span class="sxs-lookup"><span data-stu-id="1eeba-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="1eeba-115">如果这样做未解决闪烁问题，请尝试以下解决方法：</span><span class="sxs-lookup"><span data-stu-id="1eeba-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="1eeba-116">缩小对话框大小。</span><span class="sxs-lookup"><span data-stu-id="1eeba-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="1eeba-117">选择“检查元素”\*\*\*\*，这将在新窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="1eeba-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="1eeba-118">将对话框调整为原始大小。</span><span class="sxs-lookup"><span data-stu-id="1eeba-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="1eeba-119">根据需要使用检查器。</span><span class="sxs-lookup"><span data-stu-id="1eeba-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="1eeba-120">在 Mac 上清除 Office 应用程序的缓存</span><span class="sxs-lookup"><span data-stu-id="1eeba-120">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
