---
ms.date: 05/03/2019
description: 了解如何实现易失性和脱机流式处理自定义函数。
title: 函数中的可变值
localization_priority: Normal
ms.openlocfilehash: 1ca3edc3de2d9ac5f2171004f89466352c5cfa1e
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627995"
---
# <a name="volatile-values-in-functions"></a>函数中的可变值

可变函数是值在每次计算单元格时更改的函数。 即使函数的所有参数都不变, 该值也可以更改。 每当 Excel 重新计算时，这些函数即会重新计算。 例如，假设某个单元格调用函数 `NOW`。 每当调用 `NOW` 时，它将自动返回当前的日期和时间。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。 可参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)，来获取 Excel 可变函数的完整列表。

利用自定义函数, 您可以创建自己的可变函数, 这在处理日期、时间、随机编号和建模时可能很有用。 例如, [Monte Carlo 模拟](https://en.wikipedia.org/wiki/Monte_Carlo_method
)要求生成随机输入以确定最佳解决方案。

如果选择自动生成 JSON 文件, 则使用 JSDOC 注释标记`@volatile`声明一个可变函数。 有关自动生成的详细信息, 请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。

## <a name="next-steps"></a>后续步骤
了解如何[在自定义函数中保存状态](custom-functions-save-state.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数参数选项](custom-functions-parameter-options.md)
* [自定义函数元数据](custom-functions-json.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
