---
ms.date: 09/27/2018
description: 了解 Excel 自定义函数的最佳做法和建议的模式。
title: 自定义函数的最佳做法
ms.openlocfilehash: 4590682a9efa3048605686763f9af28f2fad20a4
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348112"
---
# <a name="custom-functions-best-practices-preview"></a>自定义函数的最佳做法（预览）

本文介绍在 Excel 中开发自定义函数的最佳做法。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>错误处理

构建定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 自定义函数的错误处理与 [Excel JavaScript API 的错误处理大体相同](excel-add-ins-error-handling.md)。 在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="debugging"></a>调试

目前，调试 Excel 自定义函数的最佳方法是到首先在 **Excel Online** 内[旁加载](../testing/sideload-office-add-ins-for-testing.md)加载项。然后，可以通过结合下列方法使用[浏览器自带的 F12 键调试工具](../testing/debug-add-ins-in-office-online.md)来调试自定义函数：

- 使用自定义函数代码中的 `console.log` 语句实时发送输出到控制台。

- 使用自定义函数代码内的 `debugger;` 语句来指定 F12 窗口打开时执行将暂停的断点。 例如，如果 F12 窗口打开时以下函数运行，则执行将在 `debugger;` 语句上暂停，使你能够在函数返回前手动检查参数值。  `debugger;` 语句在 F12 窗口未打开时在 Excel Online 中不起作用。 目前，`debugger;` 语句在 Excel for Windows 中不起作用。

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

如果加载项未能注册，请[验证为托管加载项应用程序的 Web 服务器正确配置了 SSL 证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

如果你在 Office 2016 桌面中测试加载项，可以启用[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)，以调试加载项 XML 清单文件以及若干安装和运行时条件等问题。


## <a name="mapping-function-names-to-json-metadata"></a>将函数名称映射到 JSON 元数据

如[自定义函数概述](custom-functions-overview.md)文章中所述，自定义函数项目必须包含一个 JSON 元数据文件，它提供了 Excel 注册并向最终用户提供自定义函数所需的信息。 此外，在定义自定义函数的 JavaScript 文件内，你必须提供信息以指定 JSON 元数据文件中的哪个函数对象对应于 JavaScript 文件中的每个自定义函数。

例如，下面的代码示例定义自定义函数 `add`，然后指定函数 `add` 对应于 JSON 元数据文件中 `id` 属性值是 **ADD** 的对象。

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

在 JavaScript 文件中创建自定义函数并指定 JSON 元数据文件中的对应信息时，请记住以下最佳做法。

* 在 JavaScript 文件中，以小骆驼拼写法指定函数名称。 例如，函数名称 `addTenToInput` 用小骆驼拼写法编写：名称的第一个单词开头以小写字母开头，名称中的每个后续单词开头以大写字母开头。

* 在 JSON 元数据文件中，以大写字母指定每个 `name` 属性的值。 `name` 属性定义了最终用户将在 Excel 中看到的函数名称。 在每个自定义函数名称中使用大写字母为最终用户提供了一致的体验，在 Excel 中所有内置函数名称都是大写。

* 在 JSON 元数据文件中，以大写字母指定每个 `id` 属性的值。 这样使  JavaScript 代码中的 `CustomFunctionMappings` 语句对应于 JSON 元数据文件中的 `id` 属性的部分显而易见（前提是函数名称使用小骆驼拼写法，如前面所建议）。

* 在 JSON 元数据文件中，确保每个 `id` 属性的值在文件范围内是唯一的。 即，元数据文件中没有两个函数对象具有相同的 `id` 值。 此外，不要在元数据文件中指定两个仅大小写不同的 `id` 值。 例如，不要将一个函数对象的 `id` 值定义为 **add**，而将另一函数对象的 `id` 值定义为 **ADD**。

* JSON 元数据文件中的 `id` 属性值在映射到相应的 JavaScript 函数名称后，不要更改其值。 可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不应在 `id` 属性值确定后再更改其值。

* 在 JavaScript 文件中，在相同位置指定所有自定义函数映射。 例如，下面的代码示例定义两个自定义的函数，并指定这两个函数的映射信息。

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

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    下面的示例演示对应于此 JavaScript 代码示例中定义的函数的 JSON 元数据。

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

## <a name="see-also"></a>另请参阅

- [在 Excel 中创建自定义函数](custom-functions-overview.md)
- [自定义函数元数据](custom-functions-json.md)
- [Excel 自定义函数运行时](custom-functions-runtime.md)
