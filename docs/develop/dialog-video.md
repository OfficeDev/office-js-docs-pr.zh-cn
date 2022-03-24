---
title: 使用 Office 对话框播放视频
description: 了解如何在"开始"对话框中打开Office视频。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: a9f222f52d1ee22a4b0b84eb62ea24e6e48e63d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743506"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>使用Office对话框显示视频

本文介绍如何在加载项对话框中Office视频。

> [!NOTE]
> 本文要求你熟悉使用 Office 对话框的基础知识，如在 Office 加载项中使用 [Office 对话框 API 中所述](dialog-api-in-office-add-ins.md)。

若要使用对话框 API 在对话框中Office视频，请按照以下步骤操作。

1. 创建包含 iframe 且不包含其他内容的页面。 页面必须与主机页在同一域中。 有关主机页的提醒，请参阅从主机 [页打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。 `src`在 iframe 的 属性中，指向联机视频的 URL。 视频 URL 必须使用 HTTPS 协议。 本文将此页面称为"video.dialogbox.html"。 下面是一个标注示例。

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. 在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。
3. 如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。 有关详细信息，请参阅"[错误和事件Office对话框中](dialog-handle-errors-events.md)。

有关在对话框中播放的视频示例，请参阅视频 [位置设计模式](../design/first-run-experience-patterns.md#video-placemat)。

![Screenshot showing a video playing in an add-in dialog box in front of Excel.](../images/video-placemats-dialog-open.png)
