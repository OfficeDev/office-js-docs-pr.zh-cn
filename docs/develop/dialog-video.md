---
title: 使用 Office 对话框播放视频
description: 了解如何在 Office 对话框中打开和播放视频
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 407eec467ed8ed51350f6195a3607c430524e6b4
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650070"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>使用 Office 对话框显示视频

本文介绍如何在 Office 外接程序对话框中播放视频。

> [!NOTE]
> 本文假定您熟悉使用 Office 对话框的基础知识，如在[Office 外接程序中使用 office 对话框 API](dialog-api-in-office-add-ins.md)中所述。

若要使用 Office 对话框 API 在对话框中播放视频，请按照以下步骤操作：

1. 创建包含 iframe 但不包含其他任何内容的页面。 页面必须与主机页位于同一域中。 有关主机页面的提示，请参阅[从主机页面打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。 在 iframe `src`的属性中，指向联机视频的 URL。 视频 URL 必须使用 HTTPS 协议。 在本文中，我们将调用此页面 "video.dialogbox.html"。 下面的示例展示了标记：

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. 在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。
3. 如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。 有关详细信息，请参阅[Office 对话框中的错误和事件](dialog-handle-errors-events.md)。

有关在对话框中播放视频的示例，请参阅[视频占位图片设计模式](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat)。

![在外接程序对话框中播放视频的屏幕截图](../images/video-placemats-dialog-open.png)
