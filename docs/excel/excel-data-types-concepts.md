---
title: Excel JavaScript API 数据类型核心概念
description: 了解在 Office 加载项中使用 Excel 数据类型的核心概念。
ms.date: 05/18/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 61485451bf5e0d7dff96a5f4f215def49425e571
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628082"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel 数据类型核心概念（预览版）

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

本文介绍如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 来处理数据类型。 它引入了对数据类型开发至关重要的核心概念。

## <a name="core-concepts"></a>核心概念

使用 [`Range.valuesAsJson`](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member) 属性处理数据类型值。 此属性类似于 [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member)，但 `Range.values` 只返回四种基本类型：字符串、数字、布尔或错误值。 `Range.valuesAsJson` 返回有关这四种基本类型的扩展信息，并且此属性可以返回数据类型，例如带格式数字值、实体和 Web 图像。

`valuesAsJson` 属性返回 [CellValue](/javascript/api/excel/excel.cellvalue) 类型别名，而这是以下数据类型的 [联合](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types)。

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

[CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) 对象是和 `*CellValue` 类型其余部分的 [交集](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types)。 它本身不是数据类型。 `CellValueExtraProperties` 对象的属性与所有数据类型一起使用，用于指定与覆盖单元格值相关的详细信息。

### <a name="json-schema"></a>JSON 架构

每个数据类型都使用为该类型设计的 JSON 元数据架构。 这将定义数据的 [CellValueType](/javascript/api/excel/excel.cellvaluetype) 以及有关单元格的其他信息，例如 `basicValue`、`numberFormat` 或 `address`。 每个 `CellValueType` 都具有符合该类型的可用属性。 例如，`webImage` 类型包括 [altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member) 和 [attribution](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member) 属性。 以下部分显示带格式数字值、实体值和 Web 图像数据类型的 JSON 代码示例。

每个数据类型的 JSON 元数据架构还包括一个或多个只读属性，这些属性在计算遇到不兼容的方案时使用，例如 Excel 版本不符合数据类型功能的最低内部版本号要求。 属性 `basicType` 是每个数据类型的 JSON 元数据的一部分，它始终是只读属性。 当 `basicType` 数据类型不受支持或格式不正确时，属性用作回退。

## <a name="formatted-number-values"></a>带格式数字值

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) 对象使 Excel 加载项能够定义值的 `numberFormat` 属性。 分配后，此数字格式将使用该值进行计算，并可由函数返回。

以下 JSON 代码示例显示了格式化数字值的完整架构。 代码示例中的 `myDate` 带格式数字值在 Excel UI 中显示为 **1/16/1990**。 如果不满足数据类型功能的最低兼容性要求，则计算使用 `basicValue` 代替格式化的数字。

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>实体值

实体值是数据类型的容器，类似于面向对象的编程中的对象。 实体还支持数组作为实体值的属性。 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) 对象允许加载项定义属性，如 `type`、`text` 和 `properties`。 `properties` 属性使实体值能够定义并包含其他数据类型。

`basicType` 和 `basicValue` 属性定义了如果未满足使用数据类型的最低兼容性要求，计算将如何读取此实体数据类型。 在该方案中，此实体数据类型显示为 **#VALUE!** Excel UI 中的错误。

以下 JSON 代码示例显示了包含文本、图像、日期和其他文本值的实体值的完整架构。

```TypeScript
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

实体值还提供可创建实体的卡片的 `layouts` 属性。 该卡片在 Excel UI 中显示为模式窗口，并且可以显示实体值中包含的其他信息，而不仅显示单元格中可见的信息。 要了解详细信息，请参阅 [使用具有实体值数据类型的卡片](excel-data-types-entity-card.md)。

## <a name="web-image-values"></a>Web 图像值

[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) 对象创建将图像存储为 [实体](#entity-values) 的一部分或作为区域中独立值的功能。此对象提供许多属性，包括 `address`、`altText` 和 `relatedImagesAddress`。

`basicType` 和 `basicValue`属性定义了如果未满足使用数据类型功能的最低兼容性要求，计算将如何读取 Web 图像数据类型。 在该方案中，此 Web 图像数据类型显示为 **#VALUE!** Excel UI 中的错误。

以下 JSON 代码示例显示了 Web 图像的完整架构。

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
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
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

每个错误对象都可以通过 `errorSubType` 属性访问枚举，并且此枚举包含有关错误的其他数据。 例如，`BlockedErrorCellValue` 错误对象可以访问 [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype) 枚举。 `BlockedErrorCellValueSubType` 枚举提供有关导致错误原因的其他数据。

## <a name="see-also"></a>另请参阅

- [ Excel 加载项中的数据类型的概述](excel-data-types-overview.md)
- [使用具有实体值数据类型的卡片](excel-data-types-entity-card.md)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
- [自定义函数和数据类型](custom-functions-data-types-concepts.md)
