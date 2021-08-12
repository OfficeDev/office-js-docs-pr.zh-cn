---
title: 编辑单元格时延迟执行
description: 了解如何在编辑单元格时延迟 Excel.run 方法的执行。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 41bbfba3894bcef0c1fd075ce76557dfdc4ba4721b7bc7b19ca21756b86ccc4d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084279"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>编辑单元格时延迟执行

`Excel.run`具有一个重载，该重载接受[Excel。RunOptions](/javascript/api/excel/excel.runoptions)对象。 这包含一组影响函数运行时平台行为的属性。 当前支持以下属性。

- `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。 未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
