---
ms.date: 05/08/2019
description: Excel 自定义函数中的常见问题疑难解答。
title: 自定义函数疑难解答
localization_priority: Priority
ms.openlocfilehash: 999b1fb9b89050ab5c6bcf87e1aac9d2fce13702
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952052"
---
# <a name="troubleshoot-custom-functions"></a>自定义函数疑难解答

开发自定义函数时，创建和测试函数可能会遇到产品错误。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

若要解决这些问题，可以[启用运行时日志记录以捕获错误](#enable-runtime-logging)，并参考[Excel 的本机错误消息](#check-for-excel-error-messages)。 另外，检查常见错误，例如未正确[有未解析的 promise](#ensure-promises-return) 以及忘记[关联函数](#my-functions-wont-load-associate-functions)。

## <a name="enable-runtime-logging"></a>启用运行时日志记录

如果在 Windows 上的 Office 中测试加载项，应[启用运行时日志记录](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。 运行时日志记录将 `console.log` 语句传递给创建的单独日志文件，以帮助发现问题。 这些语句涵盖了各种错误，其中包括加载项的 XML 清单文件、运行时条件或自定义函数安装的相关错误。  有关运行时日志记录的详细信息，请参阅[使用运行时日志记录调试加载项](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。  

### <a name="check-for-excel-error-messages"></a>检查 Excel 错误消息

Excel 有许多内置错误消息，如果存在计算错误，系统会将向单元格返回这些错误消息。 自定义函数仅使用以下错误消息：`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A` 和 `#BUSY!`。

通常情况下，这些错误可能对应于你在 Excel 中熟悉的错误。 有一些特定于自定义函数的异常，如下所示：

- `#NAME` 错误通常意味着注册函数时出错。
- `#VALUE` 错误通常是指函数的脚本文件中出现了错误。
- `#N/A` 错误也可能是注册的函数无法运行的迹象。 这通常是因为缺少 `CustomFunctions.associate` 命令。
- `#REF!` 错误可能指示函数名称与已存在的加载项中的函数名称相同。

## <a name="clear-the-office-cache"></a>清除 Office 缓存

与自定义函数相关的信息由 Office 缓存。 有时候，开发和反复重新加载带有自定义函数的加载项时，变更可能不会显示。 可以通过清除 Office 缓存修复此问题。 有关详细信息，请参阅[使用清单验证和解决问题](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)一文中的“清除 Office 缓存”部分。

## <a name="common-issues"></a>常见问题

### <a name="my-functions-wont-load-associate-functions"></a>我的函数无法加载：关联函数

在自定义函数的脚本文件中，需要将每个自定义函数与在 [JSON 元数据文件](custom-functions-json.md)中指定的 ID 相关联。 使用 `CustomFunctions.associate()` 方法可实现此操作。 通常，在每个函数之后或脚本文件的末尾调用此方法。 如果没有关联自定义函数，它将不起作用。

下面的示例显示了一个 add 函数，后跟一个与相应的 JSON ID `ADD` 相关联的函数名称 `add`。

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

有关此过程的更多信息，请参阅[将函数名称与 JSON 元数据相关联](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)。

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>无法从 localhost 打开加载项：使用本地环回异常

如果看到错误“我们无法从 localhost 打开此加载项”，则需要启用本地环回异常。 有关如何执行此操作的详细信息，请参阅[此 Microsoft 支持文章](https://support.microsoft.com/zh-CN/help/4490419/local-loopback-exemption-does-not-work)。

### <a name="ensure-promises-return"></a>确保返回 promise

在 Excel 等待自定义函数完成时，它会在单元格中 显示 #BUSY!。 如果自定义函数代码返回一个 promise，但 promise 不返回结果，则 Excel 将继续显示 #BUSY!。 查看函数以确保所有 promise 都正确地向单元格返回结果。

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>错误：开发服务器已在端口 3000 上运行

有时候，运行 `npm start` 时，你可能会看到开发服务器已在端口 3000（或加载项使用的任何端口）上运行的错误。 可以通过运行 `npm stop` 或关闭 Node.js 窗口停止开发服务器运行。 但在某些情况下，开发服务器可能需要几分钟才能实际停止运行。

## <a name="reporting-feedback"></a>报告反馈

如果遇到本文中未记录的问题，请告诉我们。 有两种方法可以报告问题。

### <a name="in-excel-on-windows-or-mac"></a>在 Wndows 或 Mac 上的 Excel 中

如果使用 Windows 版 Excel 或 Mac 版 Excel，可以直接从 Excel 向 Office 扩展性团队报告反馈。 为此，请选择“文件”->“反馈”->“发送哭脸”****。 发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。

### <a name="in-github"></a>在 Github 中

可以随时通过任何文档页底部的“内容反馈”功能提交所遇到的问题，也可以[直接向自定义功能存储库提交新问题](https://github.com/OfficeDev/Excel-Custom-Functions/issues)。

## <a name="next-steps"></a>后续步骤
了解如何[调试自定义函数](custom-functions-debugging.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据自动生成](custom-functions-json-autogeneration.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [让自定义函数与 XLL 用户定义的函数兼容](make-custom-functions-compatible-with-xll-udf.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
