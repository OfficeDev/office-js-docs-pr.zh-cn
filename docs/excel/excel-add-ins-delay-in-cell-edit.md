---
title: 编辑单元格时延迟执行
description: 了解如何在编辑单元格时延迟 Excel.run 方法的执行。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: b7b28064ef4d313639391e63cba780351b5623f9
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349516"
---
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="80a16-103">编辑单元格时延迟执行</span><span class="sxs-lookup"><span data-stu-id="80a16-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="80a16-104">`Excel.run`具有一个重载，该重载接受[Excel。RunOptions](/javascript/api/excel/excel.runoptions)对象。</span><span class="sxs-lookup"><span data-stu-id="80a16-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="80a16-105">这包含一组影响函数运行时平台行为的属性。</span><span class="sxs-lookup"><span data-stu-id="80a16-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="80a16-106">当前支持以下属性。</span><span class="sxs-lookup"><span data-stu-id="80a16-106">The following property is currently supported.</span></span>

- <span data-ttu-id="80a16-107">`delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。</span><span class="sxs-lookup"><span data-stu-id="80a16-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="80a16-108">若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。</span><span class="sxs-lookup"><span data-stu-id="80a16-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="80a16-109">若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。</span><span class="sxs-lookup"><span data-stu-id="80a16-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="80a16-110">未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。</span><span class="sxs-lookup"><span data-stu-id="80a16-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
