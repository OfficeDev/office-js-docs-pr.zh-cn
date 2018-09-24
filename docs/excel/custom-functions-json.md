---
ms.date: 09/20/2018
description: 在 Excel 中定义自定义函数的元数据。
title: 在 Excel 中的自定义函数的元数据
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062142"
---
# <a name="custom-functions-metadata"></a>自定义函数元数据

在 Excel 加载项内定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含一个 JSON 元数据文件，它提供 Excel 需要用来注册自定义函数并使其为最终用户可用的信息。 本文介绍了 JSON 元数据文件的格式。

> [!NOTE]
> 有关其他文件的信息，你必须在加载项项目中加入其他文件才能启用自定义函数，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md#learn-the-basics)。

## <a name="example-metadata"></a>元数据示例

下面的示例显示用于定义自定义函数的加载项的 JSON 元数据文件的内容。 下面示例中的各节提供了有关此 JSON 示例中各个属性的详细信息。

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
> [OfficeDev/Excel-Custom-Functions GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)中提供了完整的示例 JSON 文件。

## <a name="functions"></a>函数 

 `functions` 属性是自定义函数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No  |  Excel UI 中显示的函数的说明。 例如，**将摄氏度值转换为华氏度**。 |
|  `helpUrl`  |  string  |   No  |  用户可在其中获取有关函数的信息的 URL。 （它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。 |
| `id`     | string | Yes | 函数的唯一 ID。 设置之后，不应更改此 ID。 |
|  `name`  |  string  |  Yes  |  用户选择函数时出现在 Excel 用户界面中的函数名称（以命名空间为前缀）。 它不必与 JavaScript 中定义的函数名称相同。 |
|  `options`  |  object  |  No  |  使你可以自定义 Excel 执行函数的方式和时间等的某些方面。 有关详细信息，请参阅[选项对象](#options-object)。 |
|  `parameters`  |  array  |  Yes  |  定义函数的输入参数的数组。 有关详细信息，请参阅[参数数组](#parameters-array)。 |
|  `result`  |  object  |  Yes  |  定义函数返回的信息类型的对象。 有关详细信息，请参阅[结果对象](#result-object)。 |

## <a name="options"></a>选项

`options` 对象使你可以自定义 Excel 执行函数的方式和时间等的某些方面。 下表列出了 `options` 对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  否，默认值为 `false` 。  |  如果为 `true`，每当用户执行的操作会取消函数时，Excel 将调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。 如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。 （请***不要*** 在 `parameters` 属性中注册此参数）。 在函数的主体中，必须将处理程序分配给 `caller.onCanceled` 成员。 有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function)。 |
|  `stream`  |  boolean  |  否，默认值为 `false` 。  |  如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。 （请***不要*** 在 `parameters` 属性中注册此参数）。 函数不应存在 `return` 语句。 相反，结果值将作为 `caller.setResult` 回调方法的参数传递。 有关详细信息，请参阅[流式函数](custom-functions-overview.md#streamed-functions)。 |

## <a name="parameters"></a>参数

`parameters` 属性是参数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  参数的描述。  |
|  `dimensionality`  |  string  |  No  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。  |
|  `name`  |  string  |  Yes  |  参数的名称。 此名称显示在 Excel 的 intelliSense 中。  |
|  `type`  |  string  |  No  |  参数的数据类型。 必须是**布尔值**、**数字**或**字符串**。  |

## <a name="result"></a>结果

`results` 对象定义函数返回的信息类型。 下表列出了 `result` 对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  No  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。 |
|  `type`  |  string  |  Yes  |  参数的数据类型。 必须是**布尔值**、**数字**或**字符串**。  |

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数的最佳做法](custom-functions-best-practices.md)