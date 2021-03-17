---
title: 清单文件中 EquivalentAddin 元素
description: 指定等效 COM 加载项或 XLL 的向后兼容性。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836835"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="18afd-103">EquivalentAddin 元素</span><span class="sxs-lookup"><span data-stu-id="18afd-103">EquivalentAddin element</span></span>

<span data-ttu-id="18afd-104">指定等效 COM 加载项或 XLL 的向后兼容性。</span><span class="sxs-lookup"><span data-stu-id="18afd-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="18afd-105">**外接程序类型：** 任务窗格、自定义函数</span><span class="sxs-lookup"><span data-stu-id="18afd-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="18afd-106">语法</span><span class="sxs-lookup"><span data-stu-id="18afd-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="18afd-107">包含于</span><span class="sxs-lookup"><span data-stu-id="18afd-107">Contained in</span></span>

[<span data-ttu-id="18afd-108">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="18afd-108">EquivalentAddins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="18afd-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="18afd-109">Must contain</span></span>

[<span data-ttu-id="18afd-110">类型</span><span class="sxs-lookup"><span data-stu-id="18afd-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="18afd-111">可以包含</span><span class="sxs-lookup"><span data-stu-id="18afd-111">Can contain</span></span>

<span data-ttu-id="18afd-112">[ProgId](progid.md) 
[FileName](filename.md)</span><span class="sxs-lookup"><span data-stu-id="18afd-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="18afd-113">备注</span><span class="sxs-lookup"><span data-stu-id="18afd-113">Remarks</span></span>

<span data-ttu-id="18afd-114">若要将 COM 加载项指定为等效加载项，请提供 和 `ProgId` `Type` 元素。</span><span class="sxs-lookup"><span data-stu-id="18afd-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="18afd-115">若要将 XLL 指定为等效的外接程序，请提供 和 `FileName` `Type` 元素。</span><span class="sxs-lookup"><span data-stu-id="18afd-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="18afd-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="18afd-116">See also</span></span>

- [<span data-ttu-id="18afd-117">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="18afd-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="18afd-118">让 Office 加载项与现有 COM 加载项兼容</span><span class="sxs-lookup"><span data-stu-id="18afd-118">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)