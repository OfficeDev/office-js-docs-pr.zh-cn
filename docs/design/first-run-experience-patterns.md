---
title: Office 外接程序的首次运行体验模式
description: 了解在 Office 外接程序中设计首次运行体验的最佳实践。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 00785df2cfd2f41b41917ea720c154e24b72f779
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132065"
---
# <a name="first-run-experience-patterns"></a>首次运行体验模式

首次运行体验模式 (FRE) 是对外接程序的用户介绍。 用户首次打开外接程序时，将会显示 FRE，其中提供有外接程序的函数、功能和/或权益相关的见解。 此体验有助于塑造用户对外接程序的印象，并提高用户继续使用你的外接程序的可能性。

## <a name="best-practices"></a>最佳做法

创建首次运行体验时，请按照以下最佳做法：

|允许事项|禁止事项|
|:------|:------|
|提供了外接程序中的主要操作的简要介绍。 | 不包括与入门无关的信息和标注。
|让用户有机会完成可以积极影响其外接程序使用的操作。 | 不要期望用户可以一次性学完全部内容。 重点关注可提供最大价值的操作。
|创建用户期望完成的富有吸引力的体验。 | 不要强制用户单击使用首次运行体验。 为用户提供可绕过首次运行体验的选项。 |

向用户显示首次运行体验一次还是定期显示对你的方案来说非常重要。 例如，如果只是定期使用外接程序，则用户可能不太熟悉外接程序，因此，再次使用首次运行体验可能会有用处。

根据需要应用以下模式，以创建或提升外接程序的首次运行体验。

## <a name="carousel"></a>旋转式传送

旋转式传送让用户能够在开始使用外接程序之前浏览一系列功能或信息页面。

*图1。允许用户提前或跳过轮播流的起始页*

![图示在 Office 桌面应用程序任务窗格的首次运行体验中显示轮播的步骤1。 在此示例中，任务窗格的右上部包含一个 "Skip" 操作。](../images/add-in-FRE-step-1.png)

*图2。将轮播屏幕的数量最小化，以有效地传递邮件所需的屏幕数量*

![图示在 Office 桌面应用程序任务窗格的首次运行体验中显示轮播的步骤2。 在此示例中，任务窗格中有3个轮播屏幕。](../images/add-in-FRE-step-2.png)

*图3。提供对操作的明确调用，以退出首次运行体验*

![图示在 Office 桌面应用程序任务窗格的首次运行体验中显示轮播的步骤3。 在此示例中，任务窗格的第三个和最后一个屏幕显示了开始使用的按钮。](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a>值占位图片

值占位通过徽标占位、明确的价值主张、功能亮点或汇总和行动号召传递外接程序的价值主张。

*图4。值占位图片，带有徽标、清除价值主张、功能摘要和行动要求*

![图中显示了在 Office 桌面应用程序任务窗格的首次运行体验中占位图片的值。 在此示例中，任务窗格显示加载项徽标、加载项说明以及开始使用按钮。](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a>视频占位图片

视频占位图片可以在用户开始使用外接程序之前向其显示视频。

*图5。第一次运行视频占位图片-屏幕包含视频中的静止图像和 "播放" 按钮，并清除 "操作-操作" 按钮*

![在 Office 桌面应用程序任务窗格的首次运行体验中显示视频占位图片的图示](../images/add-in-FRE-video.png)

*图6。视频播放器-在对话框窗口中显示有视频的用户*

![在背景中显示带有 Office 桌面应用程序和外接程序任务窗格的对话框窗口中的视频的插图](../images/add-in-FRE-video-dialog.png)
