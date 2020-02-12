---
title: 在 Mac 上调试 Office 加载项
description: ''
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 38aca8b9c5245ee83ed79c94497c26250d726245
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950934"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="4d9cd-102">在 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d9cd-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="4d9cd-p101">由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="4d9cd-105">在 Mac 上使用 Safari Web 检查器进行调试</span><span class="sxs-lookup"><span data-stu-id="4d9cd-105">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="4d9cd-106">如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-106">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="4d9cd-107">要在 Mac 上调试 Office 加载项，必须拥有 Mac OS High Sierra 和 Mac Office 版本：16.9.1（内部版本 18012504）或更高版本。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-107">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="4d9cd-108">如果没有 Office Mac 内部版本，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个版本。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-108">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="4d9cd-109">首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="4d9cd-109">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="4d9cd-110">然后，打开 Office 应用程序并[旁加载你的加载项](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-110">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="4d9cd-111">右键单击加载项，应在上下文菜单中看到一个“**检查元素**”选项。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-111">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="4d9cd-112">选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-112">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4d9cd-113">如果你尝试使用检查器时对话框闪烁，请将 Office 更新到最新版本。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-113">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="4d9cd-114">如果这样做未解决闪烁问题，请尝试以下解决方法：</span><span class="sxs-lookup"><span data-stu-id="4d9cd-114">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="4d9cd-115">缩小对话框大小。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-115">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="4d9cd-116">选择“检查元素”\*\*\*\*，这将在新窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-116">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="4d9cd-117">将对话框调整为原始大小。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-117">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="4d9cd-118">根据需要使用检查器。</span><span class="sxs-lookup"><span data-stu-id="4d9cd-118">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="4d9cd-119">在 Mac 上清除 Office 应用程序的缓存</span><span class="sxs-lookup"><span data-stu-id="4d9cd-119">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
