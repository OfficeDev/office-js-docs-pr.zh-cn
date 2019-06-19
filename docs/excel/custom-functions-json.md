---
ms.date: 06/17/2019
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中自定义函数的元数据
localization_priority: Normal
ms.openlocfilehash: a7715bcdd125d44ec887f8b779ac0673b4a12af0
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059858"
---
# <a name="custom-functions-metadata"></a>自定义函数元数据

在 Excel 加载项中定义[自定义函数](custom-functions-overview.md)时, 加载项项目包含 JSON 元数据文件, 该文件提供了 Excel 注册自定义函数并使其可供最终用户使用的信息。

此文件的生成方式为:

- 您, 在手写 JSON 文件中
- 从您在函数开头输入的 JSDoc 注释

自定义函数在用户首次运行外接程序且在所有工作簿中对同一用户可用时注册。

本文介绍了 JSON 元数据文件的格式, 假定您正在手动编写元数据文件。 有关 JSDoc 注释 JSON 文件生成的信息, 请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。

有关为启用自定义函数必须在加载项项目中包含的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。

托管 JSON 文件的服务器上的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)，以便自定义函数在 Excel Online 中正常工作。

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
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
> [OfficeDev/Excel 自定义函数](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub 存储库的提交历史记录中提供了完整的示例 JSON 文件。 随着项目已调整为自动生成 JSON, 手写 JSON 的完整示例仅在项目的早期版本中可用。

## <a name="functions"></a>functions 

`functions` 属性是自定义函数对象的一个数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否  |  最终用户在 Excel 中看到的函数的说明。 例如，**将摄氏度值转换为华氏度**。 |
|  `helpUrl`  |  string  |   否  |  提供有关函数的信息的 URL。 （它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。 |
| `id`     | string | 是 | 函数的唯一 ID。 此 ID 只能包含字母数字字符和句点，设置后不应更改。 |
|  `name`  |  string  |  是  |  最终用户在 Excel 中看到的函数的名称。 在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。 |
|  `options`  |  object  |  否  |  使用户能够自定义 Excel 执行函数的方式和时间。 有关详细信息，请参阅[选项](#options)。 |
|  `parameters`  |  array  |  是  |  定义函数的输入参数的数组。 有关详细信息，请参阅[参数](#parameters)。 |
|  `result`  |  object  |  是  |  定义函数返回的信息类型的对象。 有关详细信息，请参阅[结果](#result)。 |

## <a name="options"></a>options

`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。 下表列出了 `options` 对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  否<br/><br/>默认值为 `false`。  |  如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。 可取消函数通常仅用于返回单个结果的异步函数, 并需要处理对数据请求的取消操作。 函数不能同时为流式处理和可取消。 有关详细信息, 请参阅[Make a 流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)结尾附近的注释。 |
|  `requiresAddress`  | boolean | 否 <br/><br/>默认值为 `false`。 | <br /><br /> 如果为 true, 则自定义函数可以访问调用自定义函数的单元格的地址。 若要获取调用自定义函数的单元格的地址, 请在自定义函数中使用 context。 有关详细信息，请参阅[确定调用自定义函数的单元格](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function)。 不能将自定义函数同时设置为流式处理和 requiresAddress。 使用此选项时, "调用" 参数必须是在 options 中传递的最后一个参数。 |
|  `stream`  |  boolean  |  否<br/><br/>默认值为 `false`。  |  如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 函数不应存在 `return` 语句。 相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。 有关详细信息，请参阅[流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。 |
|  `volatile`  | boolean | 否 <br/><br/>默认值为 `false`。 | <br /><br /> 如果为 `true`，则该函数会在每次 Excel 重新计算时（而不是仅当公式的从属值发生更改时）进行重新计算。 函数不能同时为流式处理和可变。 如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。 |

## <a name="parameters"></a>参数

`parameters` 属性是参数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否 |  参数的说明。 这显示在 Excel 的 intelliSense 中。  |
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。  |
|  `name`  |  string  |  是  |  参数的名称。 此名称显示在 Excel 的 intelliSense 中。  |
|  `type`  |  string  |  否  |  参数的数据类型。 可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。 如果未指定此属性，则数据类型默认为 **any**。 |
|  `optional`  | boolean | 否 | 如果为 `true`，则参数是可选的。 |

## <a name="result"></a>结果

`result` 对象定义函数返回的信息类型。 下表列出了 `result` 对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。 |

## <a name="next-steps"></a>后续步骤
了解[有关命名函数](custom-functions-naming.md)或了解如何使用前面所述的手写 JSON 方法对[函数进行本地化](custom-functions-localize.md)的最佳做法。

## <a name="see-also"></a>另请参阅

* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [自定义函数参数选项](custom-functions-parameter-options.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)