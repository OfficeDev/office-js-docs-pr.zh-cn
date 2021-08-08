---
title: Office 外接程序的图标准则
description: 大致了解如何设计图标以及外接程序命令的 Fresh 和 Monoline 设计样式。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 3f1c2d4cbe748d2ac214f0dd0c10a988435eeb4c34bbfe9c7b157b911cd4eb09
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082350"
---
# <a name="icons"></a>图标

图标是行为或概念的可视化表示形式。 它们通常用于为控件和命令添加含义。 实际或符号化的视觉对象使用户能够以与标记帮助用户浏览其环境的相同方式浏览 UI。 这些视觉对象应简单明了，并且只包含所需的详细信息，以使客户能够快速分析他们在选择控件时将会发生的操作。

Office 应用功能区界面具有标准视觉样式。 这可以确保一致性并熟悉各个 Office 应用程序。 这些准则将有助于你为解决方案设计一组适合作为 Office 固有部分的 PNG 资产。

许多 HTML 容器包含带有插图的控件。 使用 Fabric Core 的自定义字体Office外接程序中的样式图标。 Fabric [Core](fabric-core.md)提供的图标字体包含许多通用标志Office，你可以缩放、颜色和样式以满足你的需求。 如果你有带自己图标集的现有视觉语言，则可在 HTML 画布中随意使用。 构建自己带标准图标集的品牌的连续性是任何设计语言的重要组成部分。 请注意避免与 Office 隐喻产生冲突导致客户混淆。

## <a name="design-icons-for-add-in-commands"></a>加载项命令的设计图标

[外接程序命令](add-in-commands.md)添加按钮、文本和 Office UI 图标。 外接程序命令按钮应提供有意义的图标和标签，以便清楚地标识用户在使用命令时执行的操作。 以下文章提供了样式和生产指南，可帮助你设计与 Office 无缝集成。

- 有关单声道样式Microsoft 365，请参阅单声道[样式图标指南Office外接程序。](add-in-icons-monoline.md)
- 有关 2013+非订阅Office样式，请参阅全新样式图标指南Office[外接程序。](add-in-icons-fresh.md)

> [!NOTE]
> 必须选择一种样式或另一种样式，无论外接程序是在非订阅Microsoft 365运行，外接程序都将使用相同的Office。

## <a name="see-also"></a>另请参阅

- [加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [Excel、Word 和 PowerPoint 的加载项命令](../design/add-in-commands.md)
