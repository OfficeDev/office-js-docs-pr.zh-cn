---
title: 自定义函数和数据类型核心概念
description: 了解在自定义函数Excel数据类型的核心概念。
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 3b7e735f78ca7b6dcdffa3bd5e8ba9c9d3093766
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749404"
---
# <a name="custom-functions-and-data-types-core-concepts-preview"></a>自定义函数和数据类型的核心概念 (预览) 

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

数据类型通过扩展对Excel字符串、数字、布尔值和错误值 (四种数据类型的支持来增强 JavaScript API) 。 数据类型包括对实体值中的格式化数字值、Web 图像、实体值和数组的支持。 自定义函数接受数据类型作为输入和输出值，这扩展了自定义函数的计算能力。

若要了解有关在加载项中Excel数据类型，请参阅Excel[数据类型核心概念](excel-data-types-concepts.md)。

## <a name="how-custom-functions-handle-data-types"></a>自定义函数如何处理数据类型

自定义函数可以识别数据类型，并接受它们作为参数值。 自定义函数可以新建一数据类型返回值的值。 自定义函数对数据类型使用与 JavaScript API Excel相同的 JSON 架构，此 JSON 架构在自定义函数计算和评估时进行维护。

> [!NOTE]
> 自定义函数不支持由数据类型提供的增强错误对象的完整功能。 自定义函数可以接受数据类型错误对象，但它不会在整个计算过程中得到维护。 目前，自定义函数仅支持 [CustomFunctions.Error 对象中包含的错误](custom-functions-errors.md)。

## <a name="enable-data-types-for-custom-functions"></a>为自定义函数启用数据类型

若要使用此功能，需要手动更新 JSON 元数据。 对于更临时的测试，你可以自定义Script Lab设置，而不是手动更新 JSON 元数据。 以下各节更详细地概述了这些步骤。

### <a name="manually-update-json-metadata"></a>手动更新 JSON 元数据

自定义函数项目包括 JSON 元数据文件。 此 JSON 元数据文件与数据类型 API 使用的 JSON 架构不同。 若要将数据类型与自定义函数集成，必须手动更新自定义函数 JSON 元数据文件，以包括 属性 `allowCustomDataForDataTypeAny` 。 将此属性设置为 `true` 。

有关手动 JSON 创建过程的完整说明，请参阅手动为自定义函数 [创建 JSON 元数据](custom-functions-json.md)。 有关[此属性的其他详细信息，请参阅 allowCustomDataForDataTypeAny。](custom-functions-json.md#allowcustomdatafordatatypeany-preview)

### <a name="script-lab-option"></a>Script Lab选项

除了上一节所述的手动 JSON 元数据更新之外，自定义函数与数据类型的集成还可用于测试 Script Lab。 若要了解有关 javaScript Script Lab，请参阅使用[Office JavaScript API Script Lab。](../overview/explore-with-script-lab.md) 若要使用 Script Lab测试此功能，请执行以下步骤来更新设置。

1. 打开"Script Lab **代码**"任务窗格。
1. 在右下角 **，选择设置** 按钮。
1. Go to the **User 设置** tab and enter `allowCustomDataForDataTypeAny: true` .

![Screenshot showing the steps to enable data types for custom functions in Script Lab.](../images/custom-functions-script-lab-data-type.png)

## <a name="output-a-formatted-number-value"></a>输出带格式的数值

下面的代码示例演示如何使用自定义函数数据类型 [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) 对象。 该函数采用基本数字和格式设置作为输入参数，并返回一个格式化数据类型值作为输出。

```js
/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}
```

## <a name="input-an-entity-value"></a>输入实体值

下面的代码示例显示了一个自定义函数，该函数将 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) 数据类型作为输入。 如果 `attribute` 参数设置为 `text` ，则函数返回 `text` 实体值的 属性。 否则，函数将 `basicValue` 返回实体值的 属性。

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {any} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type == "Entity") {
        if (attribute == "text") {
            return value.text;
        } else {
            return value.properties[attribute].basicValue;
        }
    } else {
        return JSON.stringify(value);
    }
}
```

## <a name="see-also"></a>另请参阅

* [自定义函数和数据类型概述](custom-functions-data-types-overview.md)
* [ Excel 加载项中的数据类型的概述](excel-data-types-overview.md)
* [Excel 数据类型核心概念](excel-data-types-concepts.md)
* [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
