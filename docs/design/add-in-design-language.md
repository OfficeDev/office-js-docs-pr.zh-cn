---
title: Office 加载项设计语言
description: ''
ms.date: 12/04/2017
localization_priority: Normal
ms.openlocfilehash: c4dcd4d08a52c101878b9739eeeb45c00836514e
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950402"
---
# <a name="office-add-in-design-language"></a>Office 加载项设计语言

Office 设计语言是一种简单明了的视觉对象系统，它可确保体验的一致性。它包含一组用于定义 Office 接口的可视化元素，包括：

- 标准字样
- 公用调色板
- 一组版式大小和权重
- 图标准则
- 共享图标资源
- 动画定义
- 常见组件

[Office UI Fabric](https://developer.microsoft.com/fabric) 是用于通过 Office 设计语言构建的官方前端框架。使用 Fabric 是可选的，但它是确保外接程序感觉像是 Office 的自然扩展的最快方法。利用 Fabric 来设计和构建补充 Office 的外接程序。

许多 Office 外接程序与先前存在的品牌相关联。你可以保留外接程序中的强大品牌及其视觉对象或组件语言。与 Office 集成的同时寻找保留自己的视觉对象语言的机会。寻找方法将 Office 颜色、版式、图标或其他样式元素置换为你自己品牌的元素。在插入客户熟悉的控件和组件时，寻找遵循通用外接程序布局或 UX 设计模式的方法。

在 Office 内插入基于主要品牌的 HTML 的 UI 会对客户产生不一致性。找到一个能够在 Office 中无缝整合的平衡点，同时与你的服务或父品牌保持明确一致。如果外接程序不适合 Office，通常是因为样式元素发生冲突。例如，版式过大和网格关闭、颜色对比度鲜明或太过强烈，或者相比 Office 动画过多且行为有差异。控件或组件的外观和行为与 Office 标准相差甚远。
