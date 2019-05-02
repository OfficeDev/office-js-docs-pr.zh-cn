---
ms.date: 04/30/2019
description: 了解如何实现易失性和脱机流式处理自定义函数。
title: 函数中的可变值 (预览)
localization_priority: Normal
ms.openlocfilehash: 63618adecff57398e1630e6b5ab43c0dbc753b36
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527301"
---
## <a name="volatile-values-in-functions"></a>函数中的可变值

可变函数是值在每次计算单元格时更改的函数。 即使函数的所有参数都不变, 该值也可以更改。 每当 Excel 重新计算时，这些函数即会重新计算。 例如，假设某个单元格调用函数 `NOW`。 每当调用 `NOW` 时，它将自动返回当前的日期和时间。

Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。 可参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)，来获取 Excel 可变函数的完整列表。

利用自定义函数, 您可以创建自己的可变函数, 这在处理日期、时间、随机编号和建模时可能很有用。 例如, [Monte Carlo 模拟](https://en.wikipedia.org/wiki/Monte_Carlo_method
)要求生成随机输入以确定最佳解决方案。

如果选择自动生成 JSON 文件, 则使用 JSDOC 注释标记`@volatile`声明一个可变函数。 有关自动生成的详细信息, 请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数元数据](custom-functions-json.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
