---
title: 清单文件中的 EquivalentAddin 元素
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059921"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="5faf7-102">EquivalentAddin 元素</span><span class="sxs-lookup"><span data-stu-id="5faf7-102">EquivalentAddin element</span></span>

<span data-ttu-id="5faf7-103">为等效的 COM 外接程序或 XLL 指定向后兼容性。</span><span class="sxs-lookup"><span data-stu-id="5faf7-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="5faf7-104">**外接类型:** 任务窗格, 自定义函数</span><span class="sxs-lookup"><span data-stu-id="5faf7-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="5faf7-105">语法</span><span class="sxs-lookup"><span data-stu-id="5faf7-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="5faf7-106">包含于</span><span class="sxs-lookup"><span data-stu-id="5faf7-106">Contained in</span></span>

[<span data-ttu-id="5faf7-107">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="5faf7-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="5faf7-108">必须包含</span><span class="sxs-lookup"><span data-stu-id="5faf7-108">Must contain</span></span>

[<span data-ttu-id="5faf7-109">Type</span><span class="sxs-lookup"><span data-stu-id="5faf7-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="5faf7-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="5faf7-110">Can contain</span></span>

<span data-ttu-id="5faf7-111">[ProgId](progid.md)
[文件名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="5faf7-111">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="5faf7-112">说明</span><span class="sxs-lookup"><span data-stu-id="5faf7-112">Remarks</span></span>

<span data-ttu-id="5faf7-113">若要将 COM 加载项指定为等效的`ProgId`加载项, 请同时提供和`Type`元素。</span><span class="sxs-lookup"><span data-stu-id="5faf7-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="5faf7-114">若要将 XLL 指定为等效的外接程序, 请同时`FileName`提供`Type`和元素。</span><span class="sxs-lookup"><span data-stu-id="5faf7-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="5faf7-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5faf7-115">See also</span></span>

- [<span data-ttu-id="5faf7-116">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="5faf7-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="5faf7-117">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="5faf7-117">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)