---
title: 清单文件中的 EquivalentAddin 元素
description: 为等效的 COM 外接程序或 XLL 指定向后兼容性。
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 425b926901b7325665eeede04263f74e4b854d50
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718284"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="c99bc-103">EquivalentAddin 元素</span><span class="sxs-lookup"><span data-stu-id="c99bc-103">EquivalentAddin element</span></span>

<span data-ttu-id="c99bc-104">为等效的 COM 外接程序或 XLL 指定向后兼容性。</span><span class="sxs-lookup"><span data-stu-id="c99bc-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="c99bc-105">**外接类型：** 任务窗格，自定义函数</span><span class="sxs-lookup"><span data-stu-id="c99bc-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="c99bc-106">语法</span><span class="sxs-lookup"><span data-stu-id="c99bc-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="c99bc-107">包含于</span><span class="sxs-lookup"><span data-stu-id="c99bc-107">Contained in</span></span>

[<span data-ttu-id="c99bc-108">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="c99bc-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="c99bc-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="c99bc-109">Must contain</span></span>

[<span data-ttu-id="c99bc-110">类型</span><span class="sxs-lookup"><span data-stu-id="c99bc-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="c99bc-111">可以包含</span><span class="sxs-lookup"><span data-stu-id="c99bc-111">Can contain</span></span>

<span data-ttu-id="c99bc-112">[ProgId](progid.md)
[文件名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="c99bc-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="c99bc-113">备注</span><span class="sxs-lookup"><span data-stu-id="c99bc-113">Remarks</span></span>

<span data-ttu-id="c99bc-114">若要将 COM 加载项指定为等效的`ProgId`加载项，请同时提供和`Type`元素。</span><span class="sxs-lookup"><span data-stu-id="c99bc-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="c99bc-115">若要将 XLL 指定为等效的外接程序，请同时`FileName`提供`Type`和元素。</span><span class="sxs-lookup"><span data-stu-id="c99bc-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="c99bc-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c99bc-116">See also</span></span>

- [<span data-ttu-id="c99bc-117">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="c99bc-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="c99bc-118">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="c99bc-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)