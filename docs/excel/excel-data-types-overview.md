---
title: Excel 加载项中的数据类型概述
description: Excel JavaScript API 中的数据类型使 Office 外接程序开发人员能够将格式化的数字值、Web 映像、实体、实体中的数组以及增强的错误作为数据类型进行处理。
ms.date: 10/14/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 92f541d3b1296de5545bfb0016448f49043abcba
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607434"
---
# <a name="overview-of-data-types-in-excel-add-ins"></a>Excel 加载项中的数据类型概述

数据类型将复杂数据结构组织为对象。 这包括格式化的数字值、Web 映像和实体作为 [实体卡](excel-data-types-entity-card.md)。

在添加数据类型之前，Excel JavaScript API 已支持字符串、数字、布尔值和错误数据类型。 Excel UI 格式设置层能够向包含四种原始数据类型的单元格添加货币、日期和其他类型的格式设置，但此格式设置层仅控制 Excel UI 中原始数据类型的显示。 即使 Excel UI 中的单元格设置为货币或日期格式，基础数字值也不会更改。 基础值与 Excel UI 中带格式的显示之间的这一差距可能导致加载项计算过程中出现混淆和错误。 数据类型 API 是解决此差距的解决方案。

数据类型将 Excel JavaScript API 支持扩展到四个原始数据类型之外， (字符串、数字、布尔值和错误) 包括 [Web 映像](excel-data-types-concepts.md#web-image-values)、 [格式化数字值](excel-data-types-concepts.md#formatted-number-values)、 [实体中的实体](excel-data-types-concepts.md#entity-values)、数组，以及改进 [的错误数据类型](excel-data-types-concepts.md#improved-error-support) 作为灵活的数据结构。 这些类型支持许多 [链接数据类型](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) 体验，在加载项计算过程中实现了精确和简化，并将 Excel 加载项的潜力扩展到 2 维网格之外。

若要了解如何使用数据类型 API，请从 [Excel 数据类型核心概念](excel-data-types-concepts.md) 文章开始。

> [!NOTE]
> 若要立即开始试验数据类型，请在 Excel 中安装 [Script Lab](../overview/explore-with-script-lab.md)并查看示 **例** 库中 **的数据类型** 部分。 还可以浏览 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/tree/prod/samples/excel/20-data-types) 存储库中的Script Lab示例。

## <a name="data-types-and-custom-functions"></a>数据类型和自定义函数

数据类型增强了自定义函数的功能。 自定义函数接受数据类型作为自定义函数的输入和自定义函数的输出，并且自定义函数对数据类型使用与 Excel JavaScript API 相同的 JSON 架构。 在自定义函数计算和求值时，对此数据类型 JSON 架构进行维护。 如果要详细了解如何将数据类型与自定义函数集成，请参阅[自定义函数和数据类型](custom-functions-data-types-concepts.md)。

## <a name="see-also"></a>另请参阅

- [Excel 数据类型核心概念](excel-data-types-concepts.md)
- [使用具有实体值数据类型的卡片](excel-data-types-entity-card.md)
- [自定义函数和数据类型](custom-functions-data-types-concepts.md)
- [在 Excel 中创建和浏览数据类型](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)