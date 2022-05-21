---
title: Office 外接程序的布局准则
description: 获取有关如何在Office加载项中布局任务窗格或对话框的指南。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 134e8c01a5a6057f84ef2f4f62c290a161e94cfa
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628073"
---
# <a name="layout"></a>布局

嵌入到 Office 中的每个 HTML 容器都将有一个布局。这些布局是外接程序的主屏幕。你将在其中创建使客户能够启动操作、修改设置、查看、滚动或导航内容的体验。设计在屏幕中具有一致布局的外接程序，以确保体验的连续性。如果你有客户熟悉使用的现有网站，请考虑重新使用现有网页中的布局。对它们进行调整以协调适应 Office HTML 容器。

有关布局的指南，请参阅 [“任务”窗格](task-pane-add-ins.md)“ [内容](content-add-ins.md)。 有关如何将[Fluent UI React](using-office-ui-fabric-react.md)或[Office UI Fabric JS](fabric-core.md) 组件组合到常见布局和用户体验流的详细信息，请参阅 [UX 设计模式模板](ux-design-pattern-templates.md)。

对布局应用以下一般准则。

- 避免 HTML 容器上的边距过窄或过宽。20 像素是理想的默认值。
- 有意对齐元素。额外缩进和新对齐点应该有助于可视化层次结构。
- Office 接口在 4 像素网格上。旨在使元素之间的填充保持在 4 的倍数。
- 界面过于拥挤可能导致混乱，并抑制触控交互的易用性。
- 在各个屏幕之间保持布局一致性。意外布局更改类似于视觉错误，这将导致对解决方案的信心和信任的缺失。
- 遵循公用的布局模式。约定可帮助用户了解如何使用界面。
- 避免冗余元素，如品牌或命令。
- 整合控件和视图，以避免需要过多地移动鼠标。
- 创建适应 HTML 容器宽度和高度的响应式体验。
