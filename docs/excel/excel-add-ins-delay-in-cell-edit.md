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
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="fc32b-103">在编辑单元格时延迟执行</span><span class="sxs-lookup"><span data-stu-id="fc32b-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="fc32b-104">`Excel.run` 具有一个在 [RunOptions](/javascript/api/excel/excel.runoptions) 对象中采用的重载。</span><span class="sxs-lookup"><span data-stu-id="fc32b-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="fc32b-105">这包含一组影响函数运行时平台行为的属性。</span><span class="sxs-lookup"><span data-stu-id="fc32b-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="fc32b-106">目前，支持以下属性：</span><span class="sxs-lookup"><span data-stu-id="fc32b-106">The following property is currently supported:</span></span>

* <span data-ttu-id="fc32b-107">`delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。</span><span class="sxs-lookup"><span data-stu-id="fc32b-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="fc32b-108">若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。</span><span class="sxs-lookup"><span data-stu-id="fc32b-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="fc32b-109">若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。</span><span class="sxs-lookup"><span data-stu-id="fc32b-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="fc32b-110">未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。</span><span class="sxs-lookup"><span data-stu-id="fc32b-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
