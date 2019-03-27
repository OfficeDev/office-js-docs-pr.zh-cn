---
ms.date: 03/19/2019
description: Excel 自定义函数中的常见问题疑难解答。
title: 自定义函数疑难解答（预览版）
localization_priority: Priority
ms.openlocfilehash: 19c3dcccce7618289dc49c3f61ce781744c24369
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871337"
---
# <a name="troubleshoot-custom-functions"></a>自定义函数疑难解答

开发自定义函数时，创建和测试函数可能会遇到产品错误。

若要解决这些问题，可以[启用运行时日志记录以捕获错误](#enable-runtime-logging)，并参考[Excel 的本机错误消息](#check-for-excel-error-messages)。 另外，检查常见错误，例如未正确[验证 SSL 证书](#verify-ssl-certificates)、[有未解析的 promise](#ensure-promises-return)，以及忘记[关联函数](#associate-your-functions)。

## <a name="enable-runtime-logging"></a>启用运行时日志记录

如果在 Windows 上的 Office 中测试加载项，应[启用运行时日志记录](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。 运行时日志记录将 `console.log` 语句传递给创建的单独日志文件，以帮助发现问题。 这些语句涵盖了各种错误，其中包括加载项的 XML 清单文件、运行时条件或自定义函数安装的相关错误。  有关运行时日志记录的详细信息，请参阅[使用运行时日志记录调试加载项](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。  

### <a name="check-for-excel-error-messages"></a>检查 Excel 错误消息

Excel 有许多内置错误消息，如果存在计算错误，系统会将向单元格返回这些错误消息。 自定义函数仅使用以下错误消息：`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A` 和 `#BUSY!`。

## <a name="common-issues"></a>常见问题

### <a name="my-add-in-wont-load-verify-certifications"></a>我的加载项无法加载：验证证书

如果加载项无法安装，请验证是否为托管加载项的 Web 服务器正确配置了 SSL 证书。 通常，如果 SSL 证书存在问题，将会在 Excel 警告中看到一条错误消息，提示无法正确安装加载项。 有关详细信息，请参阅[添加自签名证书作为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

### <a name="my-functions-wont-load-associate-functions"></a>我的函数无法加载：关联函数

在自定义函数的脚本文件中，需要将每个自定义函数与在 [JSON 元数据文件](custom-functions-json.md)中指定的 ID 相关联。 使用 `CustomFunctions.associate()` 方法可实现此操作。 通常，在每个函数之后或脚本文件的末尾调用此方法。 如果没有关联自定义函数，它将不起作用。

下面的示例显示了一个 add 函数，后跟一个与相应的 JSON ID `ADD` 相关联的函数名称 `add`。

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

有关此过程的更多信息，请参阅[将函数名称与 JSON 元数据相关联](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)。

### <a name="ensure-promises-return"></a>确保返回 promise

在 Excel 等待自定义函数完成时，它会在单元格中 显示 #BUSY!。 如果自定义函数代码返回一个 promise，但 promise 不返回结果，则 Excel 将继续显示 #BUSY!。 查看函数以确保所有 promise 都正确地向单元格返回结果。

## <a name="reporting-feedback"></a>报告反馈

如果遇到本文中未记录的问题，请告诉我们。 有两种方法可以报告问题。

### <a name="in-excel-on-windows-or-mac"></a>在 Wndows 或 Mac 上的 Excel 中

如果使用 Excel for Windows 或 Excel for Mac，可以直接从 Excel 向 Office 扩展性团队报告反馈。 为此，请选择“文件”->“反馈”->“发送哭脸”****。 发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。

### <a name="in-github"></a>在 Github 中

可以随时通过任何文档页底部的“内容反馈”功能提交所遇到的问题，也可以[直接向自定义功能存储库提交新问题](https://github.com/OfficeDev/Excel-Custom-Functions/issues)。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
