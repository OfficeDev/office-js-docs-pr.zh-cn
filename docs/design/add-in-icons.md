---
title: Office 外接程序的图标准则
description: 概述如何为外接程序命令设计图标以及新的和 Monoline 的设计样式。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 35d8e0337b412a9ddebcde5be4db4db802e88269
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093838"
---
# <a name="icons"></a>图标

图标是行为或概念的可视化表示形式。 它们通常用于为控件和命令添加含义。 实际或符号化的视觉对象使用户能够以与标记帮助用户浏览其环境的相同方式浏览 UI。 这些视觉对象应简单明了，并且只包含所需的详细信息，以使客户能够快速分析他们在选择控件时将会发生的操作。

Office 应用程序功能区接口具有标准的视觉样式。 这可以确保一致性并熟悉各个 Office 应用程序。 这些准则将有助于你为解决方案设计一组适合作为 Office 固有部分的 PNG 资产。

许多 HTML 容器包含带有插图的控件。使用 Office UI Fabric 的自定义字体在外接程序中呈现 Office 样式图标。Fabric 的图标字体包含很多针对可缩放的常见 Office 隐喻、颜色和样式的字形以满足你的需要。如果你有带自己图标集的现有视觉语言，则可在 HTML 画布中随意使用。构建自己带标准图标集的品牌的连续性是任何设计语言的重要组成部分。请注意避免与 Office 隐喻产生冲突导致客户混淆。

## <a name="design-icons-for-add-in-commands"></a>加载项命令的设计图标

[外接程序命令](add-in-commands.md)添加按钮、文本和 Office UI 图标。 外接程序命令按钮应提供有意义的图标和标签，以便清楚地标识用户在使用命令时执行的操作。 以下文章提供了样式和生产准则，可帮助您设计与 Office 无缝集成的图标。

- 有关 Microsoft 365 的 Monoline 样式，请参阅[适用于 Office 外接程序的 Monoline 样式图标准则](add-in-icons-monoline.md)。
- 有关非订阅 Office 2013 + 的全新样式，请参阅[适用于 Office 外接程序的新样式图标指南](add-in-icons-fresh.md)。

> [!NOTE]
> 您必须选择一个样式或另一个，并且您的外接程序将使用相同的图标，无论它是在 Microsoft 365 还是在非订阅 Office 中运行。

## <a name="see-also"></a>另请参阅

- [加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [Excel、Word 和 PowerPoint 的加载项命令](../design/add-in-commands.md)
