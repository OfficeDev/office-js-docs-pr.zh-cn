---
ms.date: 11/06/2020
description: 了解 Excel 自定义函数名称的要求并避免常见命名缺陷。
title: Excel 中自定义函数的命名准则
localization_priority: Normal
ms.openlocfilehash: eefd703c63311934435657bf9e6159662f908a95
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071611"
---
# <a name="custom-functions-naming-guidelines"></a>自定义函数命名准则

`id` `name` 在 JSON 元数据文件中，自定义函数由和属性标识。

- 函数 `id` 用于唯一标识 JavaScript 代码中的自定义函数。
- 函数 `name` 用作在 Excel 中向用户显示的显示名称。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

函数 `name` 可以与函数不同，例如 `id` 出于本地化目的。 通常情况下， `name` `id` 如果没有理由让函数与相同，则函数应保持不变。

函数的 `name` 并 `id` 共享一些常见要求：

- 函数 `id` 可能只使用字符 A 到 Z、从零到九、下划线和句点。

- 函数 `name` 可能使用任何 Unicode 字母字符、下划线和句点。

- 这两个函数都 `name` `id` 必须以字母开头，并且最小限制为三个字符。

Excel 使用大写字母作为内置函数名称 (如 `SUM`) 。 将大写字母用作自定义函数 `name` 和 `id` 最佳实践。

函数 `name` 不应如下所示：

- A1 到 XFD1048576 之间的任何单元格，或从 R1C1 到 R1048576C16384 之间的任何单元格。

- 任何 Excel 4.0 宏函数 (例如 `RUN` ， `ECHO`) 。  有关这些函数的完整列表，请参阅 [此 Excel 宏函数参考文档](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)。

## <a name="naming-conflicts"></a>命名冲突

如果您的函数与 `name` 已存在的外接程序中的函数相同 `name` ，则 **#REF！** 错误将出现在工作簿中。

若要修复命名冲突，请更改 `name` 外接程序中的，然后再次尝试该函数。 此外，还可以使用冲突的名称卸载加载项。 或者，如果要在不同的环境中测试外接程序，请尝试使用不同的命名空间来区分函数 (如 `NAMESPACE_NAMEOFFUNCTION`) 。

## <a name="best-practices"></a>最佳做法

- 请考虑向函数中添加多个参数，而不是使用相同或相似的名称创建多个函数。
- 避免函数名称中不明确的缩写。 清晰度比简洁性更重要。 选择一个名称（ `=INCREASETIME` 而不是） `=INC` 。
- 函数名称应指示函数的操作，如 = GETZIPCODE 而不是邮政编码。
- 对执行类似操作的函数始终使用相同的动作。 例如，使用 `=DELETEZIPCODE` 和 `=DELETEADDRESS` ，而不是 `=DELETEZIPCODE` 和 `=REMOVEADDRESS` 。
- 在命名流式处理函数时，请考虑在函数的说明中添加对该效果的注释或添加 `STREAM` 到函数名称的末尾。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>对函数名称进行本地化

您可以使用单独的 JSON 文件本地化不同语言的函数名称，并在外接程序清单文件中重写值。 避免为您的函数 `id` 提供 `name` 另一种语言的内置 Excel 函数，因为这可能会与本地化函数发生冲突。

有关本地化的完整信息，请参阅 [本地化自定义函数](custom-functions-localize.md)

## <a name="next-steps"></a>后续步骤
了解 [错误处理最佳实践](custom-functions-errors.md)。

## <a name="see-also"></a>另请参阅

* [手动创建自定义函数的 JSON 元数据](custom-functions-json.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
