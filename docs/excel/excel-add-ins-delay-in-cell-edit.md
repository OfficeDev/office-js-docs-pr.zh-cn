---
title: 编辑单元格时延迟执行
description: 了解如何在编辑单元格时延迟 Excel.run 方法的执行。
ms.date: 09/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 246faebf593e16b342606d975573a4c29279cc42
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149528"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>编辑单元格时延迟执行

`Excel.run`具有一个重载，该重载[Excel。RunOptions](/javascript/api/excel/excel.runoptions)对象。 这包含一组影响函数运行时平台行为的属性。 当前支持以下属性。

- `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。 未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
