---
title: Excel JavaScript API 数据类型核心概念
description: 了解在 Office 加载项中使用 Excel 数据类型的核心概念。
ms.date: 01/14/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a769010ad46af7bba2210d9a6f9d66082cb3f815
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074306"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel 数据类型核心概念（预览版）

> [!NOTE]
> 数据类型 API 目前仅在公共预览版中提供。 预览 API 可能会发生变更，不适合在生产环境中使用。 我们建议你仅在测试和开发环境中试用它们。 不要在生产环境或业务关键型文档中使用预览 API。
>
> 若要使用预览 API：
>
> - 必须在内容分发网络 （CDN） （https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)） 上引用 **beta** 库。 用于 TypeScript 编译和 IntelliSense 的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)位于 CDN 和 [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) 中。 可以使用 `npm install --save-dev @types/office-js-preview` 来安装这些类型。 有关其他信息，请参阅 [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM 包自述文件。
> - 可能需要加入 [Office 预览体验计划](https://insider.office.com)才能访问更新的 Office 版本。
>
> 若要在 Windows 版 Office 中试用数据类型，则 Excel 内部版本号必须大于或等于 16.0.14626.10000。 若要尝试 Mac 版 Office 中的数据类型集成，Excel 内部版本号必须大于或等于 16.55.21102600。

本文介绍如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 来处理数据类型。 它引入了对数据类型开发至关重要的核心概念。

## <a name="core-concepts"></a>核心概念

使用 [`Range.valuesAsJson`](/javascript/api/excel/excel.range#valuesAsJson) 属性处理数据类型值。 此属性类似于 [Range.values](/javascript/api/excel/excel.range#values)，但 `Range.values` 只返回四种基本类型：字符串、数字、布尔或错误值。 `Range.valuesAsJson` 可以返回有关这四种基本类型的扩展信息，此属性可以返回数据类型，例如带格式数字值、实体和 Web 图像。

### <a name="json-schema"></a>JSON 架构

每个数据类型都使用为该类型设计的 JSON 元数据架构。 这将定义数据的 [CellValueType](/javascript/api/excel/excel.cellvaluetype) 以及有关单元格的其他信息，例如 `basicValue`、`numberFormat` 或 `address`。 每个 `CellValueType` 都具有符合该类型的可用属性。 例如，`webImage` 类型包括 [altText](/javascript/api/excel/excel.webimagecellvalue#altText) 和 [attribution](/javascript/api/excel/excel.webimagecellvalue#attribution) 属性。 以下部分显示带格式数字值、实体值和 Web 图像数据类型的 JSON 代码示例。

每个数据类型的 JSON 元数据架构还包括一个或多个只读属性，这些属性在计算遇到不兼容的方案时使用，例如 Excel 版本不符合数据类型功能的最低内部版本号要求。 属性 `basicType` 是每个数据类型的 JSON 元数据的一部分，它始终是只读属性。 当 `basicType` 数据类型不受支持或格式不正确时，属性用作回退。

## <a name="formatted-number-values"></a>带格式数字值

[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) 对象使 Excel 加载项能够定义值的 `numberFormat` 属性。 分配后，此数字格式将使用该值进行计算，并可由函数返回。

以下 JSON 代码示例显示了格式化数字值的完整架构。 代码示例中的 `myDate` 带格式数字值在 Excel UI 中显示为 **1/16/1990**。 如果不满足数据类型功能的最低兼容性要求，则计算使用 `basicValue` 代替格式化的数字。

```json
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.CellValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>实体值

实体值是数据类型的容器，类似于面向对象的编程中的对象。 实体还支持数组作为实体值的属性。 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) 对象允许加载项定义属性，如 `type`、`text` 和 `properties`。 `properties` 属性使实体值能够定义并包含其他数据类型。

`basicType` 和 `basicValue` 属性定义了如果未满足使用数据类型的最低兼容性要求，计算将如何读取此实体数据类型。 在该方案中，此实体数据类型显示为 **#VALUE!** Excel UI 中的错误。

以下 JSON 代码示例显示了包含文本、图像、日期和其他文本值的实体值的完整架构。

```json
// This is an example of the complete JSON for an entity value.
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
    }, 
    basicType: Excel.CellValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

## <a name="web-image-values"></a>Web 图像值

[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) 对象创建将图像存储为 [实体](#entity-values) 或作为区域中独立值的一部分的功能。 此对象提供许多属性，包括 `address`、`altText` 和 `relatedImagesAddress`。

`basicType` 和 `basicValue`属性定义了如果未满足使用数据类型功能的最低兼容性要求，计算将如何读取 Web 图像数据类型。 在该方案中，此 Web 图像数据类型显示为 **#VALUE!** Excel UI 中的错误。

以下 JSON 代码示例显示了 Web 图像的完整架构。

```json
// This is an example of the complete JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.CellValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
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
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
- [自定义函数和数据类型](custom-functions-data-types-concepts.md)