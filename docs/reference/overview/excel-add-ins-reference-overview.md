---
title: Excel JavaScript API 概述
description: 详细了解 Excel Javascript API
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e589bd7ce814211759cc731d828e9c180339ea1f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293658"
---
# <a name="excel-javascript-api-overview"></a>Excel JavaScript API 概述

Excel 加载项通过使用 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API 包括两个 JavaScript 对象模型：

* **Excel JavaScript API**：下面是针对 Excel 的[应用程序特定 API](../../develop/application-specific-api-model.md)。 [Excel JavaScript API](/javascript/api/excel) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。

* **通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。

文档的本部分着重介绍了 Excel JavaScript API，它可用于开发面向 Excel 网页版或 Excel 2016 或更高版本的加载项中的大部分功能。 有关通用 API 的信息，请参阅[常见 JavaScript API 对象模型](../../develop/office-javascript-api-object-model.md)。

## <a name="learn-programming-concepts"></a>了解编程概念

有关重要编程概念的信息，请参阅 [Excel JavaScript API 基本编程概念](../../excel/excel-add-ins-core-concepts.md)。

有关使用 Excel JavaScript API 访问 Excel 中对象的实际经验，请完成 [Excel 加载项教程](../../tutorials/excel-tutorial.md)。

## <a name="learn-api-capabilities"></a>了解 API 功能

每个主要的 Excel API 功能都有一篇文章，探讨该功能的作用以及相关的对象模型。

* [图表](../../excel/excel-add-ins-charts.md)
* [备注](../../excel/excel-add-ins-comments.md)
* [条件格式](../../excel/excel-add-ins-conditional-formatting.md)
* [自定义函数](../../excel/custom-functions-overview.md)
* [数据验证](../../excel/excel-add-ins-data-validation.md)
* [事件](../../excel/excel-add-ins-events.md)
* [多个范围 (RangeArea)](../../excel/excel-add-ins-multiple-ranges.md)
* [数据透视表](../../excel/excel-add-ins-pivottables.md)
* [范围](../../excel/excel-add-ins-ranges.md)和[高级范围 API](../../excel/excel-add-ins-ranges-advanced.md)
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
