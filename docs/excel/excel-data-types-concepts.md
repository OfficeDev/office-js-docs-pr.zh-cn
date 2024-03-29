---
title: Excel JavaScript API 数据类型核心概念
description: 了解在 Office 加载项中使用 Excel 数据类型的核心概念。
ms.date: 10/14/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 65a69838500733f8be08a15a99baa167a946b82a
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607448"
---
# <a name="excel-data-types-core-concepts"></a>Excel 数据类型核心概念

本文介绍如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 来处理数据类型。 它引入了对数据类型开发至关重要的核心概念。

## <a name="the-valuesasjson-property"></a>`valuesAsJson` 属性

用于 `valuesAsJson` [NamedItem](/javascript/api/excel/excel.nameditem)) 的属性 (或单`valueAsJson`一属性对于在 Excel 中创建数据类型而言是不可或缺的。 此属性是 `values` 属性的扩展，例如 [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member)。 `values` 和 `valuesAsJson` 属性都用于访问单元格中的值， 但 `values` 属性仅返回以下四种基本类型之一：字符串、数字、布尔或错误（作为字符串）。 对比来看，`valuesAsJson` 返回有关这四种基本类型的扩展信息，并且此属性可以返回数据类型，例如带格式的数值、实体和 Web 图像。

以下对象提供 `valuesAsJson` 属性。

- [NamedItem](/javascript/api/excel/excel.nameditem) (为 `valueAsJson`) 
- [NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)
- [Range](/javascript/api/excel/excel.range)
- [RangeView](/javascript/api/excel/excel.rangeview)
- [TableColumn](/javascript/api/excel/excel.tablecolumn)
- [TableRow](/javascript/api/excel/excel.tablerow)

> [!NOTE]
> 某些单元格值会根据用户的区域设置而更改。 `valuesAsJsonLocal` 属性提供本地化支持，并且可用于与 `valuesAsJson` 相同的所有对象。

## <a name="cell-values"></a>单元格值

`valuesAsJson` 属性返回 [CellValue](/javascript/api/excel/excel.cellvalue) 类型别名，而这是以下数据类型的 [联合](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types)。

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

`CellValue` 类型别名还返回 [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) 对象，这是和 `*CellValue` 类型其余部分的 [交集](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types)。 它本身不是数据类型。 `CellValueExtraProperties` 对象的属性与所有数据类型一起使用，用于指定与覆盖单元格值相关的详细信息。

### <a name="json-schema"></a>JSON 架构

`valuesAsJson` 返回的每个单元格值类型使用为该类型设计的 JSON 元数据架构。 除了每个数据类型特有的附加属性，这些 JSON 元数据架构都具有共同的 `type`、`basicType` 和 `basicValue` 属性。

`type` 定义数据的 [CellValueType](/javascript/api/excel/excel.cellvaluetype)。 始终 `basicType` 为只读，在数据类型不受支持或格式不正确时用作回退。 `basicValue` 与将由 `values` 属性返回的值匹配。 `basicValue` 在计算遇到不兼容方案时用作回退，例如不支持数据类型功能的较旧版本的 Excel。 是`basicValue`只读的`ArrayCellValue`类型、`EntityCellValue``LinkedEntityCellValue`数据类型和`WebImageCellValue`数据类型。

除了所有数据类型共享的三个字段之外，每个 `*CellValue` 的 JSON 元数据架构都有根据该类型的可用属性。 例如，[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) 类型包括 `altText` 和 `attribution` 属性，而 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) 类型则提供 `properties` 和 `text` 字段。

以下部分显示带格式数字值、实体值和 Web 图像数据类型的 JSON 代码示例。

## <a name="formatted-number-values"></a>带格式数字值

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) 对象使 Excel 加载项能够定义值的 `numberFormat` 属性。 分配后，此数字格式将使用该值进行计算，并可由函数返回。

以下 JSON 代码示例显示了格式化数字值的完整架构。 代码示例中的 `myDate` 带格式数字值在 Excel UI 中显示为 **1/16/1990**。 如果不满足数据类型功能的最低兼容性要求，则计算使用 `basicValue` 代替格式化的数字。

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A read-only property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

通过打开 [Script Lab](../overview/explore-with-script-lab.md)并查看数据类型：示 **例** 库中的 [格式化数字](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-formatted-number.yaml)片段，开始试验格式化的数字值。

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
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

实体值还提供可创建实体的卡片的 `layouts` 属性。 该卡片在 Excel UI 中显示为模式窗口，并且可以显示实体值中包含的其他信息，而不仅显示单元格中可见的信息。 要了解详细信息，请参阅 [使用具有实体值数据类型的卡片](excel-data-types-entity-card.md)。

若要探索实体数据类型，请首先在 Excel [中Script Lab](../overview/explore-with-script-lab.md)并打开[数据类型：从示例库中的表代码段中的数据创建实体卡](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml)。 [数据类型：具有引用](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-references.yaml)和[数据类型的实体值：实体值归属属性](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-attribution.yaml)代码片段提供对实体功能的更深入了解。

### <a name="linked-entities"></a>已链接实体

已链接实体值或 [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) 对象是实体值的一种类型。 这些对象集成外部服务提供的数据，并可以将此数据显示为[实体卡片](excel-data-types-entity-card.md)，例如常规实体值。 通过 Excel UI 提供的[股票和地理位置数据类型](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877)是已链接实体值。

## <a name="web-image-values"></a>Web 图像值

The [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) object creates the ability to store an image as part of an [entity](#entity-values) or as an independent value in a range. This object offers many properties, including `address`, `altText`, and `relatedImagesAddress`.

`basicType` 和 `basicValue`属性定义了如果未满足使用数据类型功能的最低兼容性要求，计算将如何读取 Web 图像数据类型。 在该方案中，此 Web 图像数据类型显示为 **#VALUE!** Excel UI 中的错误。

以下 JSON 代码示例显示了 Web 图像的完整架构。

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

通过打开 [Script Lab](../overview/explore-with-script-lab.md)并选择 [“数据类型”来](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-web-image.yaml)试用 Web 映像数据类型：示 **例** 库中的 Web 映像片段。

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

通过查看 [“数据类型”](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-error-values.yaml)了解有关数据类型错误对象的详细信息：在 [Script Lab](../overview/explore-with-script-lab.md)示 **例** 库中设置错误值代码段。

## <a name="next-steps"></a>后续步骤

了解实体数据类型如何将 Excel 外接程序的潜力扩展到具有 [实体值数据类型的“使用卡](excel-data-types-entity-card.md) ”一文的二维网格之外。

使用 [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) 存储库[中的 Excel 示例中的创建和浏览数据类型](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)，通过生成和旁加载在工作簿中创建和编辑数据类型的外接程序，更深入地试验数据类型。

## <a name="see-also"></a>另请参阅

- [ Excel 加载项中的数据类型的概述](excel-data-types-overview.md)
- [使用具有实体值数据类型的卡片](excel-data-types-entity-card.md)
- [在 Excel 中创建和浏览数据类型](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [自定义函数和数据类型](custom-functions-data-types-concepts.md)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)