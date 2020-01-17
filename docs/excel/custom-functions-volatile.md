---
ms.date: 01/14/2020
description: 了解如何实现易失性和脱机流式处理自定义函数。
title: 函数中的可变值
localization_priority: Normal
ms.openlocfilehash: 57a41578f400b10806fc169fed09db7d7a66ce84
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217025"
---
# <a name="volatile-values-in-functions"></a>函数中的可变值

可变函数是值在每次计算单元格时更改的函数。 即使函数的所有参数都不变，该值也可以更改。 每当 Excel 重新计算时，这些函数即会重新计算。 例如，假设某个单元格调用函数 `NOW`。 每当调用 `NOW` 时，它将自动返回当前的日期和时间。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。 可参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)，来获取 Excel 可变函数的完整列表。

利用自定义函数，您可以创建自己的可变函数，这在处理日期、时间、随机编号和建模时可能很有用。 例如， [Monte Carlo 模拟](https://en.wikipedia.org/wiki/Monte_Carlo_method)要求生成随机输入以确定最佳解决方案。

如果选择自动生成 JSON 文件，则使用 JSDoc 注释标记`@volatile`声明一个可变函数。 有关自动生成的详细信息，请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。

可变自定义函数的示例如下所示，模拟掷出六个侧骰子的情况。

![显示自定义函数的 gif，该函数返回随机值以模拟掷出的六边骰子](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>后续步骤
了解如何[在自定义函数中保存状态](custom-functions-save-state.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数参数选项](custom-functions-parameter-options.md)
* [自定义函数元数据](custom-functions-json.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
