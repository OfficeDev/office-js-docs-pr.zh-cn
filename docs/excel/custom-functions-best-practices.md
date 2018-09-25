---
ms.date: 09/20/2018
description: 了解 Excel 自定义函数的最佳实践和建议的模式。
title: 自定义函数的最佳实践
ms.openlocfilehash: 4fe0ddc36ce1b08ea360bb556121e76cd57c3823
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004908"
---
# <a name="custom-functions-best-practices"></a>自定义函数的最佳实践

本文介绍在 Excel 中开发自定义函数的最佳实践。

## <a name="error-handling"></a>错误处理

构建定义了自定义函数的加载项时，请务必加入错误处理逻辑，以便处理运行时错误。 自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](excel-add-ins-error-handling.md)。 在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi.com/comments/" + x; 
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
目前，调试 Excel 自定义函数的最佳方法是首先在 **Excel Online** 中[旁加载](../testing/sideload-office-add-ins-for-testing.md)加载项。 然后，您可以[使用浏览器内部的 F12 调试工具](../testing/debug-add-ins-in-office-online.md)来调试自定义函数。 使用自定义函数代码中的 `console.log` 语句发送输出到实时控制台。

如果加载项无法注册，请[验证为托管加载项应用程序的 Web 服务器正确配置了 SSL 证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

如果您在 Office 2016 桌面中测试加载项，您可以启用[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)，以调试加载项 XML 指令清单文件以及若干安装和运行时条件的问题。 


## <a name="mapping-names"></a>映射名称

默认情况下，JavaScript 文件中的自定义函数的名称通常完全使用大写字母声明，并与用户在 Excel 中看到的函数名称相对应。 但是，可以通过使用 `CustomFunctionsMappings` 对象将一个或多个函数名称从 JavaScript 文件映射到不同值来更改此设置，用户将在 Excel 中看到的函数名称对应于这些值。 如果您使用难以处理大写函数名称的 uglifier、webpack 或 import 语法，这非常有用。 `CustomFunctionsMappings` 对使用 JavaScript 的项目可能是可选的，但如果您的项目使用 TypeScript， 则必须使用它。  
  
下面的代码示例定义一个“键-值”对，将 JavaScript 函数名称 `plusFortyTwo` 映射到 Excel UI 中的 `ADD42` 函数名称。 当用户选择 Excel 中的 `ADD42` 函数时，JavaScript 函数 `plusFortyTwo` 将运行。

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

下面的代码示例定义两个“键-值”对。 第一对将 JavaScript 函数名称 `plusFifty` 映射到 Excel UI 中的 `ADD50` 函数名称，第二对将 JavaScript 函数名称 `plusOneHundred` 映射到 Excel UI 中的 `ADD100` 函数名称。 当用户选择 Excel 中的 `ADD50` 函数时，JavaScript 函数 `plusFifty` 将运行。 当用户选择 Excel 中的 `ADD100` 函数时，JavaScript 函数 `plusOneHundred` 将运行。

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a>另请参阅

- [在 Excel 中创建自定义函数](custom-functions-overview.md)
- [自定义函数元数据](custom-functions-json.md)
- [Excel 自定义函数运行时](custom-functions-runtime.md)
