---
ms.date: 01/08/2019
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中的自定义函数的元数据（预览）
ms.openlocfilehash: 79fb5ac0ebcbe91b154c48cb10aad65451b6e5ee
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770628"
---
# <a name="custom-functions-metadata-preview"></a>自定义函数元数据（预览）

在 Excel 加载项中定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含 JSON 元数据文件，该文件提供 Excel 注册自定义函数并使其可供最终用户使用所需的信息。 本文介绍了 JSON 元数据文件的格式。

有关为启用自定义函数必须在加载项项目中包含的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>示例元数据

以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。 此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST", 
      "description":  "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> 在 [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub 存储库中提供了完整的示例 JSON 文件。

## <a name="functions"></a>functions 

`functions` 属性是自定义函数对象的一个数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  是否必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否  |  最终用户在 Excel 中看到的函数的说明。 例如，**将摄氏度值转换为华氏度**。 |
|  `helpUrl`  |  string  |   否  |  提供有关函数的信息的 URL。 （它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。 |
| `id`     | string | 是 | 函数的唯一 ID。 此 ID 只能包含字母数字字符和句点，设置后不应更改。 |
|  `name`  |  string  |  是  |  最终用户在 Excel 中看到的函数的名称。 在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。 |
|  `options`  |  object  |  否  |  使用户能够自定义 Excel 执行函数的方式和时间。 有关详细信息，请参阅[选项](#options)。 |
|  `parameters`  |  array  |  是  |  定义函数的输入参数的数组。 有关详细信息，请参阅[参数](#parameters)。 |
|  `result`  |  object  |  是  |  定义函数返回的信息类型的对象。 有关详细信息，请参阅[结果](#result)。 |

## <a name="options"></a>options

`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。 下表列出了 `options` 对象的属性。

|  属性  |  数据类型  |  是否必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  否<br/><br/>默认值为 `false`。  |  如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。 如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。 （请***不要***在 `parameters` 属性中注册此参数）。 在函数正文中，必须将处理程序分配给 `caller.onCanceled` 成员。 有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function)。 |
|  `stream`  |  boolean  |  否<br/><br/>默认值为 `false`。  |  如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。 （请***不要***在 `parameters` 属性中注册此参数）。 函数不应存在 `return` 语句。 相反，结果值将作为 `caller.setResult` 回调方法的参数传递。 有关详细信息，请参阅[流式处理函数](custom-functions-overview.md#streaming-functions)。 |
|  `volatile`  | boolean | 否 <br/><br/>默认值为 `false`。 | <br /><br /> 如果为 `true`，则该函数会在每次 Excel 重新计算时（而不是仅当公式的从属值发生更改时）进行重新计算。 函数不能同时为流式处理和可变。 如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。 |

## <a name="parameters"></a>参数

`parameters` 属性是参数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  是否必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否 |  参数的说明。 这显示在 Excel 的 intelliSense 中。  |
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。  |
|  `name`  |  string  |  是  |  参数的名称。 此名称显示在 Excel 的 intelliSense 中。  |
|  `type`  |  string  |  否  |  参数的数据类型。 可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。 如果未指定此属性，则数据类型默认为 **any**。 |
|  `optional`  | boolean | 否 | 如果为 `true`，则参数是可选的。 |

>[!NOTE]
> 如果可选参数的 `type` 属性未指定或设置为 `any`，则可能会发现 IDE 中的 lint 错误以及当将函数输入到 Excel 的单元格中时未显示可选参数等问题。 预计将于 2018 年 12 月有所改变。

## <a name="result"></a>结果

`result` 对象定义函数返回的信息类型。 下表列出了 `result` 对象的属性。

|  属性  |  数据类型  |  是否必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。 |
|  `type`  |  string  |  是  |  参数的数据类型。 必须是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。 |

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
