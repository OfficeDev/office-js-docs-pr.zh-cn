---
title: 清单文件中的 EquivalentAddin 元素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356848"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="c119c-102">EquivalentAddin 元素</span><span class="sxs-lookup"><span data-stu-id="c119c-102">EquivalentAddin element</span></span>

<span data-ttu-id="c119c-103">为等效的 COM 外接程序或 XLL 指定向后兼容性。</span><span class="sxs-lookup"><span data-stu-id="c119c-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="c119c-104">**外接类型:** 任务窗格, 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c119c-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="c119c-105">语法</span><span class="sxs-lookup"><span data-stu-id="c119c-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="c119c-106">包含于</span><span class="sxs-lookup"><span data-stu-id="c119c-106">Contained in</span></span>

[<span data-ttu-id="c119c-107">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="c119c-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="c119c-108">必须包含</span><span class="sxs-lookup"><span data-stu-id="c119c-108">Must contain</span></span>

[<span data-ttu-id="c119c-109">Type</span><span class="sxs-lookup"><span data-stu-id="c119c-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="c119c-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="c119c-110">Can contain</span></span>

<span data-ttu-id="c119c-111">[ProgID](progid.md)
[文件名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="c119c-111">[ProgID](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="c119c-112">说明</span><span class="sxs-lookup"><span data-stu-id="c119c-112">Remarks</span></span>

<span data-ttu-id="c119c-113">若要将 COM 加载项指定为等效的`ProgID`加载项, 请同时提供和`Type`元素。</span><span class="sxs-lookup"><span data-stu-id="c119c-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgID` and `Type` elements.</span></span> <span data-ttu-id="c119c-114">若要将 XLL 指定为等效的外接程序, 请同时`FileName`提供`Type`和元素。</span><span class="sxs-lookup"><span data-stu-id="c119c-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="c119c-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c119c-115">See also</span></span>

- [<span data-ttu-id="c119c-116">使自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="c119c-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="c119c-117">使您的 Office 外接程序与现有的 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="c119c-117">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)