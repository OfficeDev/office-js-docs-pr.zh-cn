---
title: Excel JavaScript API 概述
description: 详细了解 Excel Javascript API
ms.date: 04/05/2021
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 4b512db9028d56e9de6dcb31d03ffb0cd0d83ea6
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152617"
---
# <a name="excel-javascript-api-overview"></a>Excel JavaScript API 概述

Excel 加载项通过使用 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API 包括两个 JavaScript 对象模型：

* **Excel JavaScript API**：下面是针对 Excel 的 [应用程序特定 API](../../develop/application-specific-api-model.md)。[Excel JavaScript API](/javascript/api/excel) 随 Office 2016 引入，提供强类型的 Excel 对象，可用于访问工作表、区域、表、图表等。

* **通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。

本文档的此部分重点介绍 Excel JavaScript API，你将使用该 API 开发面向 Excel 网页版或 Excel 2016 或更高版本的加载项中的大部分功能。有关通用 API 的信息，请参阅[常用 JavaScript API 对象模型](../../develop/office-javascript-api-object-model.md)。

## <a name="learn-object-model-concepts"></a>了解对象模型概念

有关重要对象模型概念的信息，请参见 [Office 加载项中的 Excel JavaScript 对象模型](../../excel/excel-add-ins-core-concepts.md)。

有关使用 Excel JavaScript API 访问 Excel 中对象的实际经验，请完成 [Excel 加载项教程](../../tutorials/excel-tutorial.md)。

## <a name="learn-api-capabilities"></a>了解 API 功能

针对每个主要的 Excel API 功能都有一篇或多篇文章，用于探讨该功能的作用以及相关的对象模型。

* [图表](../../excel/excel-add-ins-charts.md)
* [备注](../../excel/excel-add-ins-comments.md)
* [条件格式](../../excel/excel-add-ins-conditional-formatting.md)
* [自定义函数](../../excel/custom-functions-overview.md)
* [数据验证](../../excel/excel-add-ins-data-validation.md)
* [事件](../../excel/excel-add-ins-events.md)
* [数据透视表](../../excel/excel-add-ins-pivottables.md)
* [区域](../../excel/excel-add-ins-ranges-get.md)与[单元格](../../excel/excel-add-ins-cells.md)
* [RangeAreas（多个区域）](../../excel/excel-add-ins-multiple-ranges.md)
* [性状](../../excel/excel-add-ins-shapes.md)
* [表格](../../excel/excel-add-ins-tables.md)
* [工作簿和应用程序级 API](../../excel/excel-add-ins-workbooks.md)
* [工作表](../../excel/excel-add-ins-worksheets.md)

有关 Excel JavaScript API 对象模型的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。

## <a name="try-out-code-samples-in-script-lab"></a>试用 Script Lab 中的代码示例

使用 [Script Lab](../../overview/explore-with-script-lab.md) 快速熟悉一系列展示如何使用 API 完成任务的内置示例。 你可以运行 Script Lab 中的示例来立即查看任务窗格或工作表中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。

## <a name="see-also"></a>另请参阅

* [Excel 加载项文档](../../excel/index.yml)
* [Excel 加载项概述](../../excel/excel-add-ins-overview.md)
* [Excel JavaScript API 参考](/javascript/api/excel)
* [Office 客户端应用程序和 Office 加载项的平台可用性](../../overview/office-add-in-availability.md)
* [使用特定于应用程序的 API 模型](../../develop/application-specific-api-model.md)
