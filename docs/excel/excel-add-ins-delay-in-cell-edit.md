---
title: 编辑单元格时延迟执行
description: 了解如何在编辑单元格时延迟 Excel.run 函数的执行。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c434fddf70c89d49712c96a42db772d67168a1fb
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958531"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>编辑单元格时延迟执行

`Excel.run` 具有在 [Excel.RunOptions](/javascript/api/excel/excel.runoptions) 对象中接受的重载。 这包含一组影响函数运行时平台行为的属性。 当前支持以下属性。

- `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 当 `true`用户退出单元格编辑模式时，批处理请求会延迟并运行。 当 `false`用户处于单元格编辑模式 (导致用户) 出错时，批处理请求会自动失败。 未指定任何 `delayForCellEdit` 属性的默认行为等效于它所在的 `false`时间。

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
