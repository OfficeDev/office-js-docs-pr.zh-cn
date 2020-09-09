---
title: 在编辑单元格时延迟执行
description: 了解如何在编辑单元格时延迟执行 Excel 方法。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: eb33f4cb7cce3b1f8642e00f432e708e90b5b895
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409381"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>在编辑单元格时延迟执行

`Excel.run` 具有一个在 [RunOptions](/javascript/api/excel/excel.runoptions) 对象中采用的重载。 这包含一组影响函数运行时平台行为的属性。 目前，支持以下属性：

* `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。 未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
