---
title: Excel 加载项中的数据类型概述
description: Excel JavaScript API 中的数据类型使 Office 加载项开发人员能够使用带格式数字值、Web 图像、实体值、实体值中的数组以及作为数据类型的增强型错误。
ms.date: 11/03/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 5ff0d5a055c74eeff096d45ddb6c417615775431
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749390"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Excel 加载项中的数据类型概述（预览版）

> [!NOTE]
> 数据类型 API 目前仅在公共预览版中提供。 预览 API 可能会发生变更，不适合在生产环境中使用。 不要在生产环境或业务关键型文档中使用预览 API。

> [!IMPORTANT]
> 某些数据类型 API（如 `Range.valuesAsJSON`）处于活动开发阶段，尚未提供公共预览版。 本文旨在作为概念性介绍。 本文中介绍的尚未处于公共预览状态的概念将很快发布进入预览。

Excel JavaScript API 中的数据类型使加载项开发人员能够将复杂的数据结构组织为对象，例如带格式数字值、Web 图像和实体值。

在添加数据类型之前，Excel JavaScript API 已支持字符串、数字、布尔值和错误数据类型。 Excel UI 格式设置层能够向包含四种原始数据类型的单元格添加货币、日期和其他类型的格式设置，但此格式设置层仅控制 Excel UI 中原始数据类型的显示。 即使 Excel UI 中的单元格设置为货币或日期格式，基础数字值也不会更改。 基础值与 Excel UI 中带格式的显示之间的这一差距可能导致加载项计算过程中出现混淆和错误。 自定义数据类型是解决此差距的解决方案。

数据类型将 Excel JavaScript API 支持扩展到四种原始数据类型（字符串、数字、布尔值和错误）之外，将 Web 图像、带格式数字值、实体值、实体值中的数组，以及改进的错误数据类型等灵活的数据结构包括在内。 这些类型支持许多 [链接数据类型](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) 体验，在加载项计算过程中实现了精确和简化，并将 Excel 加载项的潜力扩展到 2 维网格之外。

## <a name="data-types-and-custom-functions"></a>数据类型和自定义函数

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

数据类型增强了自定义函数的功能。 自定义函数接受数据类型作为自定义函数的输入和自定义函数的输出，并且自定义函数对数据类型使用与 Excel JavaScript API 相同的 JSON 架构。 在自定义函数计算和求值时，对此数据类型 JSON 架构进行维护。 如果要详细了解如何将数据类型与自定义函数集成，请参阅 [自定义函数和数据类型核心概念](custom-functions-data-types-concepts.md)。

## <a name="see-also"></a>另请参阅

* [Excel 数据类型核心概念](excel-data-types-concepts.md)
* [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
* [自定义函数和数据类型概述](custom-functions-data-types-overview.md)