---
ms.date: 05/06/2020
description: 在 Excel 中定义自定义函数的 JSON 元数据，并将您的函数 id 和 name 属性相关联。
title: Excel 中自定义函数的元数据
localization_priority: Normal
ms.openlocfilehash: 848a65a0eda7b8cfd6a28df16b44dbbfc207c7b9
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275992"
---
# <a name="custom-functions-metadata"></a>自定义函数元数据

如 "[自定义函数概述](custom-functions-overview.md)" 一文中所述，自定义函数项目必须包括 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能注册函数，使其可供使用。 自定义函数在用户首次运行外接程序且在所有工作簿中对同一用户可用时注册。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

我们建议在可能的情况（而不是创建自己的 JSON 文件）中使用 JSON 自动生成。 自动生成不易出现用户错误， `yo office` 搭建文件已包含此文件。 有关 JSDoc 注释 JSON 文件生成的过程的详细信息，请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。

不过，您可以从头开始创建自定义函数项目;它要求您执行以下操作：

- 编写 JSON 文件。
- 检查您的清单文件是否已连接到您的 JSON 文件。
- `id`在脚本文件中关联函数和 `name` 属性，以便注册您的函数

下图说明了使用 `yo office` 搭建文件和从草稿写入 JSON 之间的差异。
![使用 Yo 办公室和编写自己的 JSON 的差异的图像](../images/custom-functions-json.png)

> [!NOTE]
> 请记住，如果不使用生成器，请将清单连接到您创建的 JSON 文件，并通过 `<Resources>` XML 清单文件中的节 `yo office` 。

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>创作元数据并连接到清单

在项目中创建一个 JSON 文件，并提供有关函数中的函数的所有详细信息，如函数的参数。 有关函数属性的完整列表，请参阅[以下元数据示例](#json-metadata-example)和[元数据参考](#metadata-reference)。

请确保您的 XML 清单文件引用了节中的 JSON 文件 `<Resources>` ，如下例所示。

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a>JSON 元数据示例

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
      "description": "Count up from zero",
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
      "description": "Get the second highest number from a range",
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
> [OfficeDev/Excel 自定义函数](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub 存储库的提交历史记录中提供了完整的示例 JSON 文件。 随着项目已调整为自动生成 JSON，手写 JSON 的完整示例仅在项目的早期版本中可用。

## <a name="metadata-reference"></a>元数据参考

### <a name="functions"></a>functions

`functions` 属性是自定义函数对象的一个数组。 下表列出了每个对象的属性。

| 属性      | 数据类型 | 必需 | 说明                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | 否       | 最终用户在 Excel 中看到的函数的说明。 例如，**将摄氏度值转换为华氏度**。                                                            |
| `helpUrl`     | string    | 否       | 提供有关函数的信息的 URL。 （它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。                      |
| `id`          | string    | 是      | 函数的唯一 ID。 此 ID 只能包含字母数字字符和句点，设置后不应更改。                                            |
| `name`        | string    | 是      | 最终用户在 Excel 中看到的函数的名称。 在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。 |
| `options`     | object    | 否       | 使用户能够自定义 Excel 执行函数的方式和时间。 有关详细信息，请参阅[选项](#options)。                                                          |
| `parameters`  | array     | 是      | 定义函数的输入参数的数组。 有关详细信息，请参阅[参数](#parameters)。                                                                             |
| `result`      | object    | 是      | 定义函数返回的信息类型的对象。 有关详细信息，请参阅[结果](#result)。                                                                 |

### <a name="options"></a>options

`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。 下表列出了 `options` 对象的属性。

| 属性          | 数据类型 | 必需                               | 说明                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | boolean   | 否<br/><br/>默认值为 `false`。  | 如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。 可取消函数通常仅用于返回单个结果的异步函数，并需要处理对数据请求的取消操作。 函数不能同时为流式处理和可取消。 有关详细信息，请参阅[Make a 流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)结尾附近的注释。 |
| `requiresAddress` | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果为 `true` ，则自定义函数可以访问调用自定义函数的单元格的地址。 若要获取调用自定义函数的单元格的地址，请在自定义函数中使用 context。 不能将自定义函数同时设置为流式处理和 requiresAddress。 使用此选项时，"调用" 参数必须是在 options 中传递的最后一个参数。                                              |
| `stream`          | boolean   | 否<br/><br/>默认值为 `false`。  | 如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 函数不应存在 `return` 语句。 相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。 有关详细信息，请参阅[流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。                                                                                                                                                                |
| `volatile`        | boolean   | 否 <br/><br/>默认值为 `false`。 | <br /><br /> 如果 `true` 为，则函数会在 Excel 重新计算时重新计算，而不是仅在公式的依赖值发生更改时进行重新计算。 函数不能同时为流式处理和可变。 如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a>参数

`parameters` 属性是参数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否 |  参数的说明。 这将显示在 Excel 的 IntelliSense 中。  |
|  `dimensionality`  |  string  |  否  |  必须是**标量**（非数组值）或**矩阵**（二维数组）。  |
|  `name`  |  string  |  是  |  参数的名称。 此名称显示在 Excel 的 IntelliSense 中。  |
|  `type`  |  string  |  否  |  参数的数据类型。 可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。 如果未指定此属性，则数据类型默认为 **any**。 |
|  `optional`  | boolean | 否 | 如果为 `true`，则参数是可选的。 |
|`repeating`| boolean | 否 | 如果 `true` 为，则参数将从指定的数组中填充。 请注意，根据定义，所有重复参数均被视为可选参数。  |

### <a name="result"></a>结果

`result` 对象定义函数返回的信息类型。 下表列出了 `result` 对象的属性。

| 属性         | 数据类型 | 必需 | 说明                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | 否       | 必须是**标量**（非数组值）或**矩阵**（二维数组）。 |

## <a name="associating-function-names-with-json-metadata"></a>将函数名称与 JSON 元数据相关联

若要使函数正常工作，需要将函数的 `id` 属性与 JavaScript 实现相关联。 请确保存在关联，否则将无法注册该函数，也无法在 Excel 中使用它。 下面的代码示例演示如何使用方法进行关联 `CustomFunctions.associate()` 。 该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

下面的 JSON 显示了与上一个自定义函数 JavaScript 代码相关联的 JSON 元数据。

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。

- 在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。

- 在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。 也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。

- 在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。 你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。

- 在 JavaScript 文件中，使用每个函数的后面指定自定义函数关联 `CustomFunctions.associate` 。

以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。 `id`和 `name` 属性值以大写形式表示，这是描述自定义函数的最佳做法。 仅当您手动准备自己的 JSON 文件，而不是使用自动生成时，才需要添加此 JSON。 有关自动生成的详细信息，请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a>后续步骤

了解[有关命名函数](custom-functions-naming.md)或了解如何使用前面所述的手写 JSON 方法对[函数进行本地化](custom-functions-localize.md)的最佳做法。

## <a name="see-also"></a>另请参阅

- [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
- [自定义函数参数选项](custom-functions-parameter-options.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
