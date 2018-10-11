---
ms.date: 09/27/2018
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中的自定义函数的元数据
ms.openlocfilehash: e8af13b8855d6c5e1a3b1ce99edb24445e066756
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459236"
---
# <a name="custom-functions-metadata-preview"></a>自定义函数元数据 （预览）

在 Excel 加载项内定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含一个 JSON 元数据文件，它提供 Excel 需要用来注册自定义函数并使其为最终用户可用的信息。 本文介绍了 JSON 元数据文件的格式。

有关其他文件的信息，你必须在加载项项目中加入其他文件才能启用自定义函数，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>元数据示例

下面的示例显示用于定义自定义函数的加载项的 JSON 元数据文件的内容。下面示例中的各节提供了有关此 JSON 示例中各个属性的详细信息。

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
> [OfficeDev/Excel-Custom-Functions ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)GitHub 存储库中提供了完整的示例 JSON 文件。

## <a name="functions"></a>functions 

 `functions` 属性是自定义函数对象的数组。下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否  |  最终用户在 Excel 中看到的函数的描述。 例如，**将摄氏度值转换为华氏度**。 |
|  `helpUrl`  |  string  |   否  |  提供有关函数的信息的 URL。 （它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。 |
| `id`     | string | 是 | 函数的唯一 ID。设置之后，不应更改此 ID。 |
|  `name`  |  string  |  Yes  |  最终用户在 Excel 中看到的函数的名称。在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。 |
|  `options`  |  object  |  否  |  使你可以自定义 Excel 执行函数的方式和时间等的某些方面。有关详细信息，请参阅[选项对象](#options-object)。 |
|  `parameters`  |  数组  |  Yes  |  定义函数的输入参数的数组。有关详细信息，请参阅[参数数组](#parameters-array)。 |
|  `result`  |  object  |  是  |  定义函数返回的信息类型的对象。有关详细信息，请参阅[结果对象](#result-object)。 |

## <a name="options"></a>options

 `options` 对象使你可以自定义 Excel 执行函数的方式和时间等的某些方面。下表列出 `options` 对象的属性。

|  属性  |  数据类型  |  必需  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  布尔值  |  否<br/><br/>默认值为 `false`。  |  如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。如果您使用此选项，Excel 将使用其他 `caller` 参数调用 JavaScript 函数。（请***不要***在 `parameters` 属性中注册此参数）。在函数的正文中，必须将处理程序分配给 `caller.onCanceled` 成员。有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function) 。 |
|  `stream`  |  boolean  |  否<br/><br/>默认值为 `false`。  |  如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。此选项对于快速变化的数据源（如股票价格）非常有用。如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。（请***不要***在 `parameters` 属性中注册此参数）。函数不应存在 `return` 语句。相反，结果值将作为 `caller.setResult` 回调方法的参数传递。有关详细信息，请参阅[流函数](custom-functions-overview.md#streaming-functions)。 |

## <a name="parameters"></a>parameters

`parameters` 属性是参数对象的数组。下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  参数的描述。  |
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。  |
|  `name`  |  string  |  是  |  参数的名称。此名称显示在 Excel 的 IntelliSense 中。  |
|  `type`  |  string  |  No  |  参数的数据类型。必须是 **boolean**、 **number** 或 **string**。  |

## <a name="result"></a>result

 `results` 对象定义函数返回的信息类型。下表列出 `result` 对象的属性。

|  属性  |  数据类型  |  必需  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。 |
|  `type`  |  string  |  Yes  |  参数的数据类型。必须是 **boolean**、 **number** 或 **string**。  |

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数运行时](custom-functions-runtime.md)
* [自定义函数最佳做法](custom-functions-best-practices.md)
* [Excel 自定义函数教程](excel-tutorial-custom-functions.md)