---
title: 在 Mac 上调试 Office 加载项
description: ''
ms.date: 04/24/2019
localization_priority: Priority
ms.openlocfilehash: 6d77dd0d90e68c2147ffea67d12026fc194fa642
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517091"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="62b7c-102">在 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="62b7c-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="62b7c-p101">你可以使用 Visual Studio 在 Windows 上开发和调试加载项，但不能使用它在 Mac 上调试加载项。由于加载项是使用 HTML 和 JavaScript 开发的，因此它们可以跨平台工作，但不同浏览器呈现 HTML 的方式可能存在细微差别。本文介绍如何调试在 Mac 上运行的加载项。</span><span class="sxs-lookup"><span data-stu-id="62b7c-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="62b7c-106">在 Mac 上使用 Safari Web 检查器进行调试</span><span class="sxs-lookup"><span data-stu-id="62b7c-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="62b7c-107">如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="62b7c-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="62b7c-108">要在 Mac 上调试 Office 加载项，必须拥有 Mac OS High Sierra 和 Mac Office 版本：16.9.1（内部版本 18012504）或更高版本。</span><span class="sxs-lookup"><span data-stu-id="62b7c-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="62b7c-109">如果没有 Office Mac 内部版本，可以通过加入 [Office 365 开发人员计划](https://aka.ms/o365devprogram)获取一个版本。</span><span class="sxs-lookup"><span data-stu-id="62b7c-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="62b7c-110">首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="62b7c-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="62b7c-111">然后，打开 Office 应用程序并[旁加载你的加载项](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="62b7c-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="62b7c-112">右键单击加载项，应在上下文菜单中看到一个“**检查元素**”选项。</span><span class="sxs-lookup"><span data-stu-id="62b7c-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="62b7c-113">选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。</span><span class="sxs-lookup"><span data-stu-id="62b7c-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="62b7c-114">如果你尝试使用检查器时对话框闪烁，请将 Office 更新到最新版本。</span><span class="sxs-lookup"><span data-stu-id="62b7c-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="62b7c-115">如果这样做未解决闪烁问题，请尝试以下解决方法：</span><span class="sxs-lookup"><span data-stu-id="62b7c-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="62b7c-116">缩小对话框大小。</span><span class="sxs-lookup"><span data-stu-id="62b7c-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="62b7c-117">选择“检查元素”\*\*\*\*，这将在新窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="62b7c-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="62b7c-118">将对话框调整为原始大小。</span><span class="sxs-lookup"><span data-stu-id="62b7c-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="62b7c-119">根据需要使用检查器。</span><span class="sxs-lookup"><span data-stu-id="62b7c-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="62b7c-120">在 Mac 或 iPad 上清除 Office 应用程序缓存</span><span class="sxs-lookup"><span data-stu-id="62b7c-120">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="62b7c-p105">出于性能方面的考虑，外接程序通常在 Office for Mac 中缓存。通常情况下，将通过重载外接程序清除缓存。如果同一文档中存在多个外接程序，则重载后自动清除缓存的过程可能不可靠。</span><span class="sxs-lookup"><span data-stu-id="62b7c-p105">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="62b7c-124">在 Mac 上，通过删除 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的所有内容可以手动清除缓存。</span><span class="sxs-lookup"><span data-stu-id="62b7c-124">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

<span data-ttu-id="62b7c-p106">在 iPad 上，可以从外接程序中的 JavaScript 调用 `window.location.reload(true)` 来强制重载。或者，可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="62b7c-p106">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
