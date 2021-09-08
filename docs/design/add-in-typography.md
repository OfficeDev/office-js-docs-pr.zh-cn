---
title: Office 加载项的版式准则
description: 了解在加载项中Office字样和字号。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 187267c20d119ca1b3d103f32a5fd665dc903a5a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937233"
---
# <a name="typography"></a>版式

Segoe 是 Office 的标准字样。 在外接程序中使用 Segoe，以与 Office 任务窗格、对话框和内容对象保持一致。 [Fabric Core](fabric-core.md) 可让你访问 Segoe。 它在方便使用的 CSS 类中为 Segoe 的全型斜坡提供了许多不同的字体粗细和大小。 并非所有 Fabric Core 大小和权重在加载项Office外观。 为了协调适应或避免冲突，请考虑使用 Fabric Core 类型渐变的子集。 下表列出了建议在加载项中Office Fabric Core 的基类。

> [!NOTE]
> 这些基类不包含文本颜色。 对白色背景上的大多数文本使用 Fabric Core 的"中性主要"。
>
> 若要了解有关可用版式的信息，请参阅 [Web 版式](https://developer.microsoft.com/fluentui#/styles/web/typography)。

|类型 |类 |大小 |权重 |建议的用法 |
|------ |----- |---- |------ |----------------- |
|主图|.ms-font-xxl |28 像素 | Segoe Light |<ul><li>此类大于 Office 中的所有其他版式元素。请谨慎使用以避免超越可视化层次结构。</li><li>避免在有限空间中的长字符串上使用。</li><li>在使用此类的文本周围提供充足的空白空间。</li><li>常用于首次运行的信息、特大元素或其他操作调用。</li></ul> |
|标题|.ms-font-xl |21 像素 |Segoe Light | <ul><li>此类匹配 Office 应用程序的任务窗格标题。</li><li>请谨慎使用以避免出现平面版式层次结构。</li><li>通常用作对话框、页面或内容标题等顶级元素。</li></ul> |
|副标题|.ms-font-l |17 像素 |Segoe Semilight | <ul><li>此类是标题下方的第一级元素。</li><li>常用作副标题、导航元素或组标头。</li><ul> |
|正文|.ms-font-m |14 像素 |Segoe Regular |<ul><li>通常用作加载项中的正文文本。</li><ul>|
|Caption|.ms-font-xs |11 像素 | Segoe Regular |<ul><li>通常由行、标题或字段标签用于时间戳等二级或三级文本。</li><ul>|
|Annotation|.ms-font-mi |10 像素 |Segoe Semibold |<ul><li>应极少使用类型渐变中的最小步长。它仅供不需要辨别的情况使用。</li><ul>|
