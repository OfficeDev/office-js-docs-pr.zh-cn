---
ms.date: 05/03/2019
description: 了解 Excel 自定义函数名称的要求并避免出现常见命名缺陷。
title: Excel 中自定义函数的命名准则
localization_priority: Normal
ms.openlocfilehash: 64420171a90b29732745891cb691b8cd4309c53d
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706076"
---
# <a name="naming-guidelines"></a>命名准则

自定义函数由 JSON 元数据文件中的**id**和**name**属性标识。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

- 函数`id`用于唯一标识 JavaScript 代码中的自定义函数。 
- 函数`name`用作在 Excel 中向用户显示的显示名称。 

函数`name`可以与函数`id`不同, 例如出于本地化目的。 通常情况下, 如果没有`name`明显的原因, 函数应`id`保持与的相同。

函数的`name`并`id`共享一些常见要求:

- 函数`id`可能只使用字符 A 到 Z、从零到九、下划线和句点。

- 函数`name`可能使用任何 Unicode 字母字符、下划线和句点。

- 这两`name`个`id`函数都必须以字母开头, 并且最小限制为三个字符。

Excel 使用大写字母作为内置函数名称 (例如`SUM`)。 因此, 请考虑将大写字母用作自定义函数`name`和`id`最佳实践。

函数的`name`名称不应与以下相同:

- A1 到 XFD1048576 之间的任何单元格, 或从 R1C1 到 R1048576C16384 之间的任何单元格。

- 任何 Excel 4.0 宏函数 (例如`RUN`, `ECHO`)。  有关这些函数的完整列表, 请参阅[本文](https://www.microsoft.com/en-us/download/details.aspx?id=1465)。

## <a name="naming-conflicts"></a>命名冲突

如果您的`name`函数与已存在的外`name`接程序中的函数相同, 则 **#REF!** 错误将出现在工作簿中。

若要修复命名冲突, 请更改`name`外接程序中的, 然后再次尝试该函数。 此外, 还可以使用冲突的名称卸载加载项。 或者, 如果要在不同的环境中测试外接程序, 请尝试使用不同的命名空间来区分您的函数`NAMESPACE_NAMEOFFUNCTION`(如)。

## <a name="best-practices"></a>最佳做法

- 请考虑向函数中添加多个参数, 而不是使用相同或相似的名称创建多个函数。
- 函数名称应指示函数的操作, 例如 ( `=GETZIPCODE`而不是) `ZIPCODE`。
- 避免函数名称中不明确的缩写。 清晰度比简洁性更重要。 选择一个名称 ( `=INCREASETIME`而不`=INC`是)。
- 对执行类似操作的函数始终使用相同的动作。 `=DELETEZIPCODE`例如, 使用`=DELETEADDRESS`和, 而不是`=DELETEZIPCODE`和`=REMOVEADDRESS`。
- 在命名流式处理函数时, 请考虑在函数的说明中添加对该效果的注释或`STREAM`添加到函数名称的末尾。

## <a name="localizing-function-names"></a>对函数名称进行本地化

您可以使用单独的 JSON 文件本地化不同语言的函数名称, 并在外接程序清单文件中重写值。 作为一种最佳做法, 应避免在另`id`一`name`种语言中为函数提供内置 Excel 函数, 因为这可能会与本地化函数发生冲突。

有关本地化的完整信息, 请参阅[本地化自定义函数](custom-functions-localize.md)

## <a name="next-steps"></a>后续步骤
了解[错误处理最佳实践](custom-functions-errors.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
