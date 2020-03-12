---
title: 使用 Office 对话框播放视频
description: 了解如何在 Office 对话框中打开和播放视频
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 9c65dfb9c0cf1adbc827be25b655e380dc39e2d2
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596527"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="2a1bb-103">使用 Office 对话框显示视频</span><span class="sxs-lookup"><span data-stu-id="2a1bb-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="2a1bb-104">本文介绍如何在 Office 外接程序对话框中播放视频。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="2a1bb-105">本文假定您熟悉使用 Office 对话框的基础知识，如在[Office 外接程序中使用 office 对话框 API](dialog-api-in-office-add-ins.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="2a1bb-106">若要使用 Office 对话框 API 在对话框中播放视频，请按照以下步骤操作：</span><span class="sxs-lookup"><span data-stu-id="2a1bb-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="2a1bb-107">创建包含 iframe 但不包含其他任何内容的页面。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="2a1bb-108">页面必须与主机页位于同一域中。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="2a1bb-109">有关主机页面的提示，请参阅[从主机页面打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="2a1bb-110">在 iframe `src`的属性中，指向联机视频的 URL。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="2a1bb-111">视频 URL 必须使用 HTTPS 协议。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="2a1bb-112">在本文中，我们将调用此页面 "video.dialogbox.html"。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="2a1bb-113">下面的示例展示了标记：</span><span class="sxs-lookup"><span data-stu-id="2a1bb-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="2a1bb-114">在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="2a1bb-115">如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="2a1bb-116">有关详细信息，请参阅[Office 对话框中的错误和事件](dialog-handle-errors-events.md)。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="2a1bb-117">有关在对话框中播放视频的示例，请参阅[视频占位图片设计模式](../design/first-run-experience-patterns.md#video-placemat)。</span><span class="sxs-lookup"><span data-stu-id="2a1bb-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![在外接程序对话框中播放视频的屏幕截图](../images/video-placemats-dialog-open.png)
