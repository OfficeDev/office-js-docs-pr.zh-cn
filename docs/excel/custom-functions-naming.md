---
title: 自定义函数的命名Excel
description: 了解自定义函数Excel的要求，并避免常见的命名错误。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: bfc850fb2a40e7736006930c63489ec7e0c9912b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936677"
---
# <a name="custom-functions-naming-guidelines"></a>自定义函数命名准则

自定义函数由 JSON 元数据文件的 和 属性 `id` `name` 标识。

- 函数 `id` 用于唯一标识 JavaScript 代码中的自定义函数。
- 该 `name` 函数用作显示名称用户显示Excel。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

函数 `name` 可以不同于 函数 `id` ，例如用于本地化目的。 一般情况下，函数的 应保持与 `name` `id` 相同（如果没有理由区别的话）。

一个函数 `name` ， `id` 并共享一些常见要求。

- 函数只能使用字符 `id` A 到 Z、从零到九的数字、下划线和句点。

- 函数可以使用任何 `name` Unicode 字母字符、下划线和句点。

- 两个 `name` `id` 函数和 都必须以字母开头，且最小限制为三个字符。

Excel对内置函数名称使用大写字母 (如 `SUM`) 。 最好将大写字母用于自定义函数 `name` `id` 。

函数 `name` 不应与以下函数相同：

- A1 到 XFD1048576 之间的任何单元格，或 R1C1 到 R1048576C16384 之间的任何单元格。

- 任何 Excel 4.0 宏函数 (，例如 `RUN` `ECHO` ，) 。  有关这些函数的完整列表，请参阅[本Excel宏函数参考文档](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)。

## <a name="naming-conflicts"></a>命名冲突

如果函数 `name` 与已存在的加载项中的函数相同 `name` **，#REF！** 错误将显示在工作簿中。

若要修复命名冲突，请在加载项中 `name` 更改 ，然后再次尝试 函数。 您还可以使用冲突的名称卸载外接程序。 或者，如果您要在不同环境中测试外接程序，请尝试使用不同的命名空间来区分您的函数 (如 `NAMESPACE_NAMEOFFUNCTION`) 。

## <a name="best-practices"></a>最佳做法

- 请考虑向函数添加多个参数，而不是创建名称相同或相似的多个函数。
- 避免函数名称中的缩写不明确。 简洁性比简洁性更重要。 选择类似 的名称 `=INCREASETIME` ，而不是 `=INC` 。
- 函数名称应指示函数的操作，例如 =GETZIPCODE 而不是 ZIPCODE。
- 对执行类似操作的函数一致地使用相同的动词。 例如，使用 `=DELETEZIPCODE` 和 `=DELETEADDRESS` ，而不是 `=DELETEZIPCODE` 和 `=REMOVEADDRESS` 。
- 命名流式处理函数时，请考虑在函数描述中添加该效果的注释或添加到函数 `STREAM` 名称的末尾。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>本地化函数名称

可以使用单独的 JSON 文件本地化不同语言的函数名称，并替代加载项清单文件中的值。 避免为函数提供 或 作为另一种语言Excel内置函数，因为这可能与本地化函数 `id` `name` 冲突。

有关本地化的完整信息，请参阅 [本地化自定义函数](custom-functions-localize.md)

## <a name="next-steps"></a>后续步骤

了解 [错误处理最佳做法](custom-functions-errors.md)。

## <a name="see-also"></a>另请参阅

* [手动为自定义函数创建 JSON 元数据](custom-functions-json.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
