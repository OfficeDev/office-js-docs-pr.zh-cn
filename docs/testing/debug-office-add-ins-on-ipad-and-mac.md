---
title: 在 Mac 上调试 Office 加载项
description: 了解如何使用 Mac 调试Office加载项。
ms.date: 10/16/2020
localization_priority: Normal
ms.openlocfilehash: 98473e7c37b9ef5ee34d35f91688ccef65ac7d78
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350132"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="b714b-103">在 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b714b-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="b714b-p101">由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。</span><span class="sxs-lookup"><span data-stu-id="b714b-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="b714b-106">在 Mac 上使用 Safari Web 检查器进行调试</span><span class="sxs-lookup"><span data-stu-id="b714b-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="b714b-107">如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="b714b-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="b714b-108">若要能够在 Mac 上调试 Office 加载项，必须具有 Mac OS High Sierra 和 Mac Office 版本 16.9.1 (版本 18012504) 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="b714b-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office version 16.9.1 (build 18012504) or later.</span></span> <span data-ttu-id="b714b-109">如果你没有 Mac Office，可以通过加入开发人员计划获取Microsoft 365[一。](https://developer.microsoft.com/office/dev-program)</span><span class="sxs-lookup"><span data-stu-id="b714b-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="b714b-110">首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="b714b-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > <span data-ttu-id="b714b-111">Mac 应用商店版本Office不支持 `OfficeWebAddinDeveloperExtras` 标志。</span><span class="sxs-lookup"><span data-stu-id="b714b-111">Mac App Store builds of Office do not support the `OfficeWebAddinDeveloperExtras` flag.</span></span>

<span data-ttu-id="b714b-112">然后，打开 Office 应用程序并[旁加载你的加载项](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="b714b-112">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="b714b-113">右键单击加载项，应在上下文菜单中看到一个“**检查元素**”选项。</span><span class="sxs-lookup"><span data-stu-id="b714b-113">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="b714b-114">选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。</span><span class="sxs-lookup"><span data-stu-id="b714b-114">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b714b-115">如果你尝试使用检查器时对话框闪烁，请将 Office 更新到最新版本。</span><span class="sxs-lookup"><span data-stu-id="b714b-115">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="b714b-116">如果无法解决闪烁问题，请尝试以下解决方法。</span><span class="sxs-lookup"><span data-stu-id="b714b-116">If that doesn't resolve the flickering, try the following workaround.</span></span>
>
> 1. <span data-ttu-id="b714b-117">缩小对话框大小。</span><span class="sxs-lookup"><span data-stu-id="b714b-117">Reduce the size of the dialog.</span></span>
> 1. <span data-ttu-id="b714b-118">选择“检查元素”，这将在新窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="b714b-118">Choose **Inspect Element**, which opens in a new window.</span></span>
> 1. <span data-ttu-id="b714b-119">将对话框调整为原始大小。</span><span class="sxs-lookup"><span data-stu-id="b714b-119">Resize the dialog to its original size.</span></span>
> 1. <span data-ttu-id="b714b-120">根据需要使用检查器。</span><span class="sxs-lookup"><span data-stu-id="b714b-120">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="b714b-121">在 Mac 上清除 Office 应用程序的缓存</span><span class="sxs-lookup"><span data-stu-id="b714b-121">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
