---
title: Office 外接程序的布局准则
description: ''
ms.date: 06/27/2018
ms.openlocfilehash: 421860162487a3f736b13f3b74833868509eaeb1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432289"
---
# <a name="layout"></a>布局
嵌入到 Office 中的每个 HTML 容器都将有一个布局。这些布局是外接程序的主屏幕。你将在其中创建使客户能够启动操作、修改设置、查看、滚动或导航内容的体验。设计在屏幕中具有一致布局的外接程序，以确保体验的连续性。如果你有客户熟悉使用的现有网站，请考虑重新使用现有网页中的布局。对它们进行调整以协调适应 Office HTML 容器。

有关布局指南，请参阅[任务窗格](task-pane-add-ins.md)、[内容](content-add-ins.md)和[对话框](dialog-boxes.md)。若要详细了解如何将 Office UI Fabric 组件装配到通用布局和用户体验流，请参阅[用户体验设计模式模板](ux-design-pattern-templates.md)。

请遵循下面的一般布局指南：

*   避免 HTML 容器上的边距过窄或过宽。20 像素是理想的默认值。
*   有意对齐元素。额外缩进和新对齐点应该有助于可视化层次结构。
*   Office 接口在 4 像素网格上。旨在使元素之间的填充保持在 4 的倍数。
*   界面过于拥挤可能导致混乱，并抑制触控交互的易用性。
*   在各个屏幕之间保持布局一致性。意外布局更改类似于视觉错误，这将导致对解决方案的信心和信任的缺失。
*   遵循公用的布局模式。约定可帮助用户了解如何使用界面。
*   避免冗余元素，如品牌或命令。
*   整合控件和视图，以避免需要过多地移动鼠标。
*   创建适应 HTML 容器宽度和高度的响应式体验。