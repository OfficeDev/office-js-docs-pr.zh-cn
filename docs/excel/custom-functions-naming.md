---
ms.date: 02/08/2019
description: 了解 Excel 自定义函数名称的要求并避免出现常见命名缺陷。
title: Excel 中自定义函数的命名准则 (预览)
localization_priority: Normal
ms.openlocfilehash: bdf31879fb6e750fb9dea51f66c55dbc83a2dc90
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/22/2019
ms.locfileid: "30203844"
---
# <a name="naming-guidelines"></a>命名准则

自定义函数由 JSON 元数据文件中的**id**和**name**属性标识。 函数 id 用于唯一标识 JavaScript 代码中的自定义函数。 函数名称将用作在 Excel 中向用户显示的显示名称。 函数名可以与函数 ID 不同, 例如出于本地化目的。 但通常, 如果没有理由让它们不同, 则应将其保持与 ID 相同。

函数名称和函数 id 共享一些常见要求:

- 它们必须仅使用字母数字字符 (包括 Unicode)、0到9、下划线和句点。

- 它们必须以字母开头, 最小限制为三个字符。

Excel 使用大写字母作为内置函数名称 (例如`SUM`)。 因此, 请考虑将大写字母用作自定义函数名称和函数 id 作为最佳实践。

函数名称不应按如下方式命名:

- A1 到 XFD1048576 之间的任何单元格, 或从 R1C1 到 R1048576C16384 之间的任何单元格。

- 任何 Excel 4.0 宏函数 (例如`RUN`, `ECHO`)。  有关这些函数的完整列表, 请参阅[本文](https://www.microsoft.com/en-us/download/details.aspx?id=1465)。

## <a name="naming-conflicts"></a>命名冲突

如果您的函数名称与已存在的外接程序中的函数名称相同, 则 **#REF!** 错误将出现在工作簿中。

若要修复名称冲突, 请更改外接程序中的名称, 然后重试该函数。 此外, 还可以使用冲突的名称卸载加载项。 或者, 如果要在不同的环境中测试外接程序, 请尝试使用不同的命名空间来区分您的函数 (如 NAMESPACE_NAMEOFFUNCTION)。

此外, 还应考虑你希望用户在你的外接程序中使用这些功能的方式。 在许多情况下, 将多个参数添加到函数中是有意义的, 而不是使用相同或相似的名称来创建多个函数。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
