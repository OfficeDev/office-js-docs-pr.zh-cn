---
title: Excel JavaScript API 数据类型核心概念
description: 了解在 Office 加载项中使用 Excel 数据类型的核心概念。
ms.date: 11/03/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a5d4915638d67c67679095eb03bc04a48e9196dd
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749383"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel 数据类型核心概念（预览版）

> [!NOTE]
> 数据类型 API 目前仅在公共预览版中提供。 预览 API 可能会发生变更，不适合在生产环境中使用。 不要在生产环境或业务关键型文档中使用预览 API。

> [!IMPORTANT]
> 本文中所述的一些数据类型概念，如 `Range.valuesAsJSON` 处于活动开发阶段，尚未在公共预览版中提供。 本文旨在作为概念性介绍。 本文中介绍的尚未处于公共预览状态的概念将很快发布进入预览。

本文介绍如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 来处理数据类型。 它引入了对数据类型开发至关重要的核心概念。

## <a name="core-concepts"></a>核心概念

使用 `Range.valuesAsJSON` 属性处理数据类型值。 此属性类似于 [Range.values](/javascript/api/excel/excel.range#values)，但 `Range.values` 只返回四种基本类型：字符串、数字、布尔或错误值。 `Range.valuesAsJSON` 可以返回有关这四种基本类型的扩展信息，此属性可以返回数据类型，例如带格式数字值、实体和 Web 图像。

### <a name="json-schema"></a>JSON 架构

数据类型使用一致的 JSON 架构，该架构定义数据的 [CellValueType](/javascript/api/excel/excel.cellvaluetype) 和其他信息，如 `basicValue`、`numberFormat` 或 `address`。 每个 `CellValueType` 都具有符合该类型的可用属性。 例如，`webImage` 类型包括 [altText](/javascript/api/excel/excel.webimagecellvalue#altText) 和 [attribution](/javascript/api/excel/excel.webimagecellvalue#attribution) 属性。 以下部分显示带格式数字值、实体值和 Web 图像数据类型的 JSON 代码示例。

## <a name="formatted-number-values"></a>带格式数字值

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) 对象使 Excel 加载项能够定义值的 `numberFormat` 属性。 分配后，此数字格式将使用该值进行计算，并可由函数返回。

以下 JSON 代码示例显示了带格式数字值。 代码示例中的 `myDate` 带格式数字值在 Excel UI 中显示为 **1/16/1990**。

```json
// This is an example of the JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>实体值

实体值是数据类型的容器，类似于面向对象的编程中的对象。 实体还支持数组作为实体值的属性。 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) 对象允许加载项定义属性，如 `type`、`text` 和 `properties`。 `properties` 属性使实体值能够定义并包含其他数据类型。

以下 JSON 代码示例显示了包含文本、图像、日期和其他文本值的实体值。

```json
// This is an example of the JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }
};
```

## <a name="web-image-values"></a>Web 图像值

[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) 对象创建将图像存储为 [实体](#entity-values) 或作为区域中独立值的一部分的功能。 此对象提供许多属性，包括 `address`、`altText` 和 `relatedImagesAddress`。

以下 JSON 代码示例演示了如何表示 Web 图像。

```json
// This is an example of the JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw"
};
```

## <a name="improved-error-support"></a>改进的错误支持

数据类型 API 将现有 Excel UI 错误公开为对象。 现在，因为这些错误可作为对象访问，加载项就可以定义或检索属性，如 `type`、`errorType` 和 `errorSubType`。

下面是通过数据类型扩展支持的所有错误对象的列表。

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

每个错误对象都可以通过 `errorSubType` 属性访问枚举，并且此枚举包含有关错误的其他数据。 例如，`BlockedErrorCellValue` 错误对象可以访问 [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype) 枚举。 `BlockedErrorCellValueSubType` 枚举提供有关导致错误原因的其他数据。

## <a name="see-also"></a>另请参阅

- [ Excel 加载项中的数据类型的概述](excel-data-types-overview.md)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
- [自定义函数和数据类型概述](custom-functions-data-types-overview.md)