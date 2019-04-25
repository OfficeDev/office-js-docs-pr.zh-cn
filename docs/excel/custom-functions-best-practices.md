---
ms.date: 01/08/2019
description: 了解在 Excel 中开发自定义函数的最佳实践。
title: 自定义函数最佳实践（预览）
localization_priority: Normal
ms.openlocfilehash: 4efcd0ba5efb0dc7450192694e8f0750de43b8a8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448607"
---
# <a name="custom-functions-best-practices-preview"></a>自定义函数最佳实践（预览）

本文介绍了在 Excel 中开发自定义函数的最佳实践。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a>故障排除

1. 如果要在 Windows 版 Office 中测试外接程序，则应启用**[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)**，以解决外接程序的 XML 清单文件及多个安装和运行时条件问题。 运行时日志记录将 `console.log` 语句写入日志文件，以帮你发现问题。

2. 如果一个或多个自定义函数与以前注册的外接程序的自定义函数冲突, 则不会加载外接程序。 在这种情况下, 您可以删除现有加载项, 或者如果在开发加载项时遇到此错误, 则可以在清单中指定不同的命名空间名称。

3. 若要向 Excel 自定义函数团队报告有关此故障排除方法的反馈，请发送团队反馈。 若要执行此操作，请选择“**文件|反馈|发送哭脸**”。 发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。

## <a name="associating-function-names-with-json-metadata"></a>将函数名称与 JSON 元数据相关联

如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能构成完整的函数。 若要使函数正常工作, 需要将 id 与 JavaScript 实现相关联。 请确保存在关联, 否则将不会调用该函数。

以下代码示例展示了如何执行此关联操作。 该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。

* 在 JSON 元数据文件中，函数的 `name` 和 `id` 只能使用大写字母。 不要使用大小写字母混合或仅使用小写字母。 如果这样做，你最终可能会得到两个值，这些值只会因情况而异，从而导致意外覆盖你的函数。 例如，`id` 值为 **add** 的函数对象稍后可以通过声明在 `id` 值为 **ADD** 的函数对象文件中覆盖。 此外，`name` 属性还会定义最终用户将在 Excel 中看到的函数名称。 使用大写字母作为每个自定义函数的名称可在 Excel 中提供一致的体验，其中所有内置函数名称均为大写。

* 在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。

* 在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。 也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。 

* 在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。 你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。

* 在 JavaScript 文件中，请在同一位置指定所有自定义函数关联。 例如，以下代码示例定义了两个自定义函数，并接着指定了这两个函数的关联信息。

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。 请注意，在此文件中，`id` 和 `name` 属性为大写字母。 

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

## <a name="declaring-optional-parameters"></a>声明可选参数 

在 Excel for Windows（版本 1812 或更高版本）中，可以声明自定义函数的可选参数。 当用户在 Excel 中调用函数时，可选参数将显示在括号中。 例如，具有一个名为 `parameter1` 的必需参数和一个名为 `parameter2` 的可选参数的函数 `FOO` 将在 Excel 中显示为 `=FOO(parameter1, [parameter2])`。

若要使某个参数可选，请在定义函数的 JSON 元数据文件中将 `"optional": true` 添加到该参数。 以下示例显示对于函数 `=ADD(first, second, [third])` 会是怎么样的。 请注意，可选 `[third]` 参数后跟两个必需参数。 必需参数将先显示在 Excel 的公式 UI 中。

```json
{
    "id": "ADD",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
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
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

定义包含一个或多个可选参数的函数时，应指定未定义可选参数时会发生什么情况。 在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。 如果未定义 `zipCode` 参数，则会将默认值设置为 98052。 如果未定义 `dayOfWeek` 参数，则会将其设置为星期三。

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a>其他注意事项

为创建一个可在多个平台（Office 外接程序的关键租户之一）上运行的外接程序，请勿访问自定义函数中的文档对象模型 (DOM) 或使用 jQuery 等这类依赖于 DOM 的库。 在自定义函数会使用 [JavaScript 运行时](custom-functions-runtime.md)的 Windows 版 Excel 中，自定义函数无法访问 DOM。

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
