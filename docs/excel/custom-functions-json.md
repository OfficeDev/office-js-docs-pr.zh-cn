---
title: 在 Excel 中手动为自定义函数创建 JSON 元数据
description: 在 Excel 中定义自定义函数的 JSON 元数据，并关联函数 ID 和名称属性。
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4bc9139b3e46bc64749a58537737db2f048ee82
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68540996"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>为自定义函数手动创建 JSON 元数据

如 [自定义函数概述](custom-functions-overview.md) 文章中所述，自定义函数项目必须包括 JSON 元数据文件和脚本 (JavaScript 或 TypeScript) 文件才能注册函数，使其可供使用。 当用户首次运行外接程序时以及之后，自定义函数将注册到所有工作簿中的同一用户。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

建议尽可能使用 JSON 自动生成，而不是创建自己的 JSON 文件。 自动生成不太容易出现用户错误， `yo office` 基架文件已包含此错误。 有关 JSDoc 标记和 JSON 自动生成过程的详细信息，请参阅 [自定义函数的自动生成 JSON 元数据](custom-functions-json-autogeneration.md)。

但是，可以从头开始创建自定义函数项目。 此过程要求执行以下操作：

- 编写 JSON 文件。
- 检查清单文件是否已连接到 JSON 文件。
- 将函 `id` 数和 `name` 属性关联到脚本文件中，以便注册函数。

下图说明了使用 `yo office` 基架文件和从头开始写入 JSON 之间的区别。

![使用 Office 外接程序的 Yeoman 生成器和编写自己的 JSON 之间的差异图像。](../images/custom-functions-json.png)

> [!NOTE]
> 如果未使用 Office 外接程序的 [Yeoman 生成器](../develop/yeoman-generator-overview.md)，请记得通过 **\<Resources\>** XML 清单文件中的节将清单连接到创建的 JSON 文件。

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>创作元数据并连接到清单

在项目中创建一个 JSON 文件，并提供有关其中函数的所有详细信息，例如函数的参数。 有关函数属性的完整列表，请参阅 [以下元数据示例](#json-metadata-example) 和 [元数据参考](#metadata-reference) 。

确保 XML 清单文件引用节中的 **\<Resources\>** JSON 文件，类似于以下示例。

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
  "allowCustomDataForDataTypeAny": true,
  "allowErrorForDataTypeAny": true,
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
> [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub 存储库的提交历史记录中提供了完整的示例 JSON 文件。 由于项目已调整为自动生成 JSON，因此手写 JSON 的完整示例仅在项目的早期版本中可用。

## <a name="metadata-reference"></a>元数据参考

### <a name="allowcustomdatafordatatypeany"></a>allowCustomDataForDataTypeAny

该 `allowCustomDataForDataTypeAny` 属性是布尔数据类型。 将此值设置为 `true` 允许自定义函数接受数据类型作为参数并返回值。 若要了解详细信息，请参阅 [自定义函数和数据类型](custom-functions-data-types-concepts.md)。

> [!NOTE]
> 与其他大多数 JSON 元数据属性不同，它是顶级属性， `allowCustomDataForDataTypeAny` 不包含子属性。 有关如何设置此属性的格式的示例，请参阅前面的 [JSON 元数据代码示例](#json-metadata-example) 。

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

该 `allowErrorForDataTypeAny` 属性是布尔数据类型。 设置值以 `true` 允许自定义函数将错误作为输入值进行处理。 当设置为`true`输入值时`allowErrorForDataTypeAny`，具有类型`any`或`any[][]`可以接受错误的所有参数作为输入值。 默认 `allowErrorForDataTypeAny` 值为 `false`.

> [!NOTE]
> 与其他 JSON 元数据属性不同，它是顶级属性， `allowErrorForDataTypeAny` 不包含子属性。 有关如何设置此属性的格式的示例，请参阅前面的 [JSON 元数据代码示例](#json-metadata-example) 。

### <a name="functions"></a>functions

`functions` 属性是自定义函数对象的一个数组。 下表列出了每个对象的属性。

| 属性      | 数据类型 | 是否必需 | 说明                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | 否       | 最终用户在 Excel 中看到的函数的说明。 例如，**将摄氏度值转换为华氏度**。                                                            |
| `helpUrl`     | string    | 否       | 提供有关函数的信息的 URL。 （它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。                      |
| `id`          | string    | 是      | 函数的唯一 ID。 此 ID 只能包含字母数字字符和句点，设置后不应更改。                                            |
| `name`        | string    | 是      | 最终用户在 Excel 中看到的函数的名称。 在 Excel 中，此函数名称由 XML 清单文件中指定的自定义函数命名空间作为前缀。 |
| `options`     | object    | 否       | 使用户能够自定义 Excel 执行函数的方式和时间。 有关详细信息，请参阅[选项](#options)。                                                          |
| `parameters`  | array     | 是      | 定义函数的输入参数的数组。 有关详细信息，请参阅 [参数](#parameters) 。                                                                             |
| `result`      | object    | 是      | 定义函数返回的信息类型的对象。 有关详细信息，请参阅[结果](#result)。                                                                 |

### <a name="options"></a>options

`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。 下表列出了 `options` 对象的属性。

| 属性          | 数据类型 | 是否必需                               | 说明 |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | boolean   | 否<br/><br/>默认值为 `false`。  | 如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。 可取消函数通常仅用于返回单个结果并需要处理数据请求取消的异步函数。 函数不能同时使用这些 `stream` 属性和 `cancelable` 属性。 |
| `requiresAddress` | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果 `true`是，自定义函数可以访问调用它的单元格的地址。 `address` [调用参数](custom-functions-parameter-options.md#invocation-parameter)的属性包含调用自定义函数的单元格的地址。 函数不能同时使用这些 `stream` 属性和 `requiresAddress` 属性。 |
| `requiresParameterAddresses` | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果 `true`是，自定义函数可以访问函数输入参数的地址。 此属性必须与`dimensionality`[结果](#result)对象的属性结合使用，并且`dimensionality`必须设置为 `matrix`。 有关详细信息，请参阅 [“检测参数的地址](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) ”。 |
| `stream`          | boolean   | 否<br/><br/>默认值为 `false`。  | 如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 函数不应存在 `return` 语句。 而是将结果值作为回调函数的 `StreamingInvocation.setResult` 参数传递。 有关详细信息，请参阅 [“创建流式处理”函数](custom-functions-web-reqs.md#make-a-streaming-function)。 |
| `volatile`        | boolean   | 否 <br/><br/>默认值为 `false`。 | 如果 `true`，函数会在每次 Excel 重新计算时重新计算，而不是仅在公式的依赖值发生更改时重新计算。 函数不能同时使用这些 `stream` 属性和 `volatile` 属性。 如果这 `stream` 两个属性都 `volatile` 设置为 `true`，则会忽略易失性属性。 |

### <a name="parameters"></a>参数

`parameters` 属性是参数对象的数组。 下表列出了每个对象的属性。

|  属性  |  数据类型  |  是否必需  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  否 |  参数的说明。 这会显示在 Excel 的 IntelliSense 中。  |
|  `dimensionality`  |  string  |  否  |  必须 `scalar` (非数组值) 或 `matrix` (二维数组) 。  |
|  `name`  |  string  |  是  |  参数的名称。 此名称显示在 Excel 的 IntelliSense 中。  |
|  `type`  |  string  |  否  |  参数的数据类型。 可以是`boolean`，`number``string`或`any`，它允许你使用前三种类型中的任何一种。 如果未指定此属性，则数据类型默认为 `any`。 |
|  `optional`  | boolean | 否 | 如果为 `true`，则参数是可选的。 |
|`repeating`| boolean | 否 | 如果 `true`是，则从指定数组填充参数。 请注意，按定义，所有重复参数的函数都被视为可选参数。  |

### <a name="result"></a>result

`result` 对象定义函数返回的信息类型。 下表列出了 `result` 对象的属性。

| 属性         | 数据类型 | 是否必需 | 说明                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | 否       | 必须 `scalar` (非数组值) 或 `matrix` (二维数组) 。 |
| `type` | string    | 否       | 结果的数据类型。 可以是`boolean`、`number``string`或`any` (，可用于使用前三种类型中的任何一种) 。 如果未指定此属性，则数据类型默认为 `any`。 |

## <a name="associating-function-names-with-json-metadata"></a>将函数名称与 JSON 元数据相关联

若要使函数正常工作，需要将函数的 `id` 属性与 JavaScript 实现相关联。 请确保存在关联，否则该函数不会注册且在 Excel 中不可用。 下面的代码示例演示如何使用该 `CustomFunctions.associate()` 函数建立关联。 该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。

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

- 在 JavaScript 文件中，指定每个函数后使用的 `CustomFunctions.associate` 自定义函数关联。

下面的示例显示了与前面 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。 和`id``name`属性值采用大写形式，这是描述自定义函数时的最佳做法。 仅当手动准备自己的 JSON 文件而不使用自动生成时，才需要添加此 JSON。 有关自动生成的详细信息，请参阅 [自定义函数的自动生成 JSON 元数据](custom-functions-json-autogeneration.md)。

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
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

了解 [命名函数或](custom-functions-naming.md) 了解如何使用前面描述的手写 JSON 方法 [本地化函](custom-functions-localize.md) 数的最佳做法。

## <a name="see-also"></a>另请参阅

- [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
- [自定义函数参数选项](custom-functions-parameter-options.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
