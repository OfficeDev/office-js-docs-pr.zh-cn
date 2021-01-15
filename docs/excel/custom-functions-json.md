---
ms.date: 12/22/2020
description: 在 Excel 中为自定义函数定义 JSON 元数据，并关联函数 ID 和名称属性。
title: 在 Excel 中为自定义函数手动创建 JSON 元数据
localization_priority: Normal
ms.openlocfilehash: 80a71c640caacbd865b0dd253f03258a64c9b1bf
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735548"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>手动为自定义函数创建 JSON 元数据

如自定义函数概述[](custom-functions-overview.md)文章中所述，自定义函数项目必须同时包括 JSON 元数据文件和脚本 (JavaScript 或 TypeScript) 文件以注册函数，使其可供使用。 自定义函数在用户首次运行外接程序时注册，之后，自定义函数可供所有工作簿中的同一用户使用。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

我们建议尽可能使用 JSON 自动生成，而不是创建自己的 JSON 文件。 自动生成不太容易出现用户错误，基架文件 `yo office` 已包含此错误。 有关 JSDoc 标记和 JSON 自动生成过程的信息，请参阅自动生成 [自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)。

但是，你可以从头开始创建自定义函数项目。 此过程需要您：

- 编写 JSON 文件。
- 检查清单文件是否连接到 JSON 文件。
- 在脚本 `id` 文件中关联函数和属性 `name` ，以便注册函数。

下图说明了使用基架文件和从头开始编写 `yo office` JSON 之间的差异。

![使用 Yo Office 和编写自己的 JSON 之间的差异的图像](../images/custom-functions-json.png)

> [!NOTE]
> 请记住，如果不使用生成器，请通过 XML 清单文件中部分将清单连接到创建的 JSON `<Resources>` `yo office` 文件。

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>创作元数据并连接到清单

在项目中创建 JSON 文件，并提供函数的所有详细信息，如函数的参数。 有关[函数属性的完整](#json-metadata-example)[列表，](#metadata-reference)请参阅以下元数据示例和元数据引用。

确保 XML 清单文件引用该节中的 JSON 文件， `<Resources>` 如以下示例所示。

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
> [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub 存储库的提交历史记录中提供了完整的 JSON 文件示例。 由于项目已调整为自动生成 JSON，因此手写 JSON 的完整示例仅在项目的早期版本中可用。

## <a name="metadata-reference"></a>元数据参考

### <a name="functions"></a>functions

`functions` 属性是自定义函数对象的一个数组。 下表列出了每个对象的属性。

| 属性      | 数据类型 | 必需 | 说明                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | 否       | 最终用户在 Excel 中看到的函数的说明。 例如，**将摄氏度值转换为华氏度**。                                                            |
| `helpUrl`     | string    | 否       | 提供有关函数的信息的 URL。 （它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。                      |
| `id`          | string    | 是      | 函数的唯一 ID。 此 ID 只能包含字母数字字符和句点，设置后不应更改。                                            |
| `name`        | string    | 是      | 最终用户在 Excel 中看到的函数的名称。 在 Excel 中，此函数名称以 XML 清单文件中指定的自定义函数命名空间作为前缀。 |
| `options`     | object    | 否       | 使用户能够自定义 Excel 执行函数的方式和时间。 有关详细信息，请参阅[选项](#options)。                                                          |
| `parameters`  | array     | 是      | 定义函数的输入参数的数组。 有关详细信息 [，](#parameters) 请参阅参数。                                                                             |
| `result`      | object    | 是      | 定义函数返回的信息类型的对象。 有关详细信息，请参阅[结果](#result)。                                                                 |

### <a name="options"></a>options

`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。 下表列出了 `options` 对象的属性。

| 属性          | 数据类型 | 必需                               | 说明 |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | boolean   | 否<br/><br/>默认值为 `false`。  | 如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。 可取消函数通常仅用于返回单个结果且需要处理数据请求取消的异步函数。 函数不能同时使用和 `stream` `cancelable` 属性。 |
| `requiresAddress` | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果 `true` ，自定义函数可以访问调用它的单元格的地址。 `address`调用参数[的属性包含](custom-functions-parameter-options.md#invocation-parameter)调用自定义函数的单元格的地址。 函数不能同时使用和 `stream` `requiresAddress` 属性。 |
| `requiresParameterAddresses` | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果 `true` ，自定义函数可以访问函数的输入参数的地址。 此属性必须与结果对象的属性结合使用，并且 `dimensionality` [](#result) `dimensionality` 必须设置为 `matrix` 。 有关详细信息 [，请参阅"检测参数的地址](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) "。 |
| `stream`          | boolean   | 否<br/><br/>默认值为 `false`。  | 如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 函数不应存在 `return` 语句。 相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。 有关详细信息，请参阅"[创建流式处理函数"。](custom-functions-web-reqs.md#make-a-streaming-function) |
| `volatile`        | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果为 ，函数将每次 Excel 重新计算时重新计算，而不是仅在公式的 `true` 从属值发生更改时重新计算。 函数不能同时使用和 `stream` `volatile` 属性。 如果将 `stream` and `volatile` 属性都设置为 `true` ，则可变属性将被忽略。 |

### <a name="parameters"></a>参数

`parameters` 属性是参数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否 |  参数的说明。 这将显示在 Excel IntelliSense。  |
|  `dimensionality`  |  string  |  否  |  必须 (一个非数组值) 或 (一 `scalar` `matrix` 个二维数组) 。  |
|  `name`  |  string  |  是  |  参数的名称。 此名称显示在 Excel IntelliSense。  |
|  `type`  |  string  |  否  |  参数的数据类型。 可以是 `boolean` 、、或 ，它允许您使用前三种类型 `number` `string` `any` 中的任意一种。 如果未指定此属性，则数据类型默认值 `any` 。 |
|  `optional`  | boolean | 否 | 如果为 `true`，则参数是可选的。 |
|`repeating`| boolean | 否 | If `true` ，参数从指定的数组填充。 请注意，根据定义，所有重复参数都被视为可选参数。  |

### <a name="result"></a>结果

`result` 对象定义函数返回的信息类型。 下表列出了 `result` 对象的属性。

| 属性         | 数据类型 | 必需 | 说明                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | 否       | 必须是 (数组值，) 二维 (`scalar` `matrix` 数组值) 。 |
| `type` | string    | 否       | 结果数据类型。 可以是 `boolean` `number` `string` 、、或 (，它允许您使用前 `any` 三种类型中的任意) 。 如果未指定此属性，则数据类型默认值 `any` 。 |

## <a name="associating-function-names-with-json-metadata"></a>将函数名称与 JSON 元数据相关联

若要使函数正常工作，需要将函数 `id` 的属性与 JavaScript 实现关联。 请确保存在关联，否则函数将不会注册且在 Excel 中不可使用。 下面的代码示例演示如何使用该方法进行 `CustomFunctions.associate()` 关联。 该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。

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

以下 JSON 显示与以前的自定义函数 JavaScript 代码关联的 JSON 元数据。

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

- 在 JavaScript 文件中，在每个函数之后指定 `CustomFunctions.associate` 自定义函数关联。

以下示例显示对应于前面 JavaScript 代码示例中定义的函数的 JSON 元数据。 和 `id` 属性值为大写，这是描述自定义函数 `name` 时的最佳操作。 只有在手动准备自己的 JSON 文件而不是使用自动生成时，才需要添加此 JSON。 有关自动生成的信息，请参阅自动生成 [自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)。

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

了解[命名函数的](custom-functions-naming.md)最佳实践，或了解如何使用前面描述的手写[](custom-functions-localize.md)JSON 方法本地化函数。

## <a name="see-also"></a>另请参阅

- [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
- [自定义函数参数选项](custom-functions-parameter-options.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
