---
title: 编辑单元格时延迟执行
description: 了解如何在编辑单元格时延迟 Excel.run 方法的执行。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1abcdb382150db486033b32d2521207ab0b7f28f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889217"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>编辑单元格时延迟执行

`Excel.run` 具有在 [Excel.RunOptions](/javascript/api/excel/excel.runoptions) 对象中接受的重载。 这包含一组影响函数运行时平台行为的属性。 当前支持以下属性。

- `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 当 `true`用户退出单元格编辑模式时，批处理请求会延迟并运行。 当 `false`用户处于单元格编辑模式 (导致用户) 出错时，批处理请求会自动失败。 未指定任何 `delayForCellEdit` 属性的默认行为等效于它所在的 `false`时间。

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
