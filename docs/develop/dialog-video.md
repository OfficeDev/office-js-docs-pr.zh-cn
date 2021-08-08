---
title: 使用 Office 对话框播放视频
description: 了解如何在"开始"对话框中打开Office视频
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 4865b15f21238431f9cf48cebb9a35253cdba0695873b429ab249c33d329d7e4
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080752"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>使用Office对话框显示视频

本文介绍如何在加载项对话框中Office视频。

> [!NOTE]
> 本文认为你已熟悉使用 Office 对话框的基础知识，如在 Office 外接程序中使用 Office 对话框[API 中所述](dialog-api-in-office-add-ins.md)。

若要使用对话框 API 在对话框中Office视频，请按照以下步骤操作。

1. 创建包含 iframe 且不包含其他内容的页面。 页面必须与主机页在同一域中。 有关主机页的提醒，请参阅从主机页 [打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。 在 `src` iframe 的 属性中，指向联机视频的 URL。 视频 URL 必须使用 HTTPS 协议。 本文将此页面称为"video.dialogbox.html"。 下面是一个标注示例。

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. 在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。
3. 如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。 有关详细信息，请参阅错误[和事件在Office对话框中](dialog-handle-errors-events.md)。

有关在对话框中播放的视频示例，请参阅视频 [位置图设计模式](../design/first-run-experience-patterns.md#video-placemat)。

![Screenshot showing a video playing in an add-in dialog box in front of Excel.](../images/video-placemats-dialog-open.png)
