---
ms.date: 06/18/2019
description: 了解在 Excel 中开发自定义函数的最佳实践。
title: 自定义函数最佳实践
localization_priority: Normal
ms.openlocfilehash: 7c836119a783f5cc7e1e7f4f52f1d21b86091bfe
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127931"
---
# <a name="custom-functions-best-practices"></a>自定义函数最佳实践

本文介绍了在 Excel 中开发自定义函数的最佳实践。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a>将函数名称与 JSON 元数据相关联

如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能构成完整的函数。 如果您使用`yo office`的是 JSON 元数据, 则可以从代码注释生成。 否则, 您需要手动生成 JSON 元数据文件。

若要使函数正常工作, 需要将函数的`id`属性与 JavaScript 实现相关联。 请确保存在关联, 否则将不会调用该函数。 下面的代码示例演示如何使用`CustomFunctions.associate()`方法进行关联。 该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。

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
    },
  ]
}
```


在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。

* 在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。

* 在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。 也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。

* 在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。 你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。

* 在 JavaScript 文件中, 使用`CustomFunctions.associate`每个函数的后面指定自定义函数关联。

以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。 `id`和`name`属性值以大写形式表示, 这是描述自定义函数的最佳做法。 仅当您手动准备自己的 JSON 文件, 而不是使用自动生成时, 才需要添加此 JSON。 有关自动生成的详细信息, 请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。

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

## <a name="additional-considerations"></a>其他注意事项

避免从自定义函数中直接或间接访问文档对象模型 (DOM) (例如, 使用 jQuery)。 在 Windows 的 Excel 中, 自定义函数使用[JavaScript 运行时](custom-functions-runtime.md), 自定义函数无法访问 DOM。

## <a name="next-steps"></a>后续步骤
了解如何[使用自定义函数执行 web 请求](custom-functions-web-reqs.md)。

## <a name="see-also"></a>另请参阅

* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [自定义函数元数据](custom-functions-json.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
