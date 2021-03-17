---
title: 清单文件中的类型元素
description: Type 元素指定等效加载项是 COM 加载项还是 XLL。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836807"
---
# <a name="type-element"></a><span data-ttu-id="d2d6d-103">Type 元素</span><span class="sxs-lookup"><span data-stu-id="d2d6d-103">Type element</span></span>

<span data-ttu-id="d2d6d-104">指定等效加载项是 COM 加载项还是 XLL。</span><span class="sxs-lookup"><span data-stu-id="d2d6d-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="d2d6d-105">**外接程序类型：** 任务窗格、自定义函数</span><span class="sxs-lookup"><span data-stu-id="d2d6d-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="d2d6d-106">语法</span><span class="sxs-lookup"><span data-stu-id="d2d6d-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="d2d6d-107">包含于</span><span class="sxs-lookup"><span data-stu-id="d2d6d-107">Contained in</span></span>

[<span data-ttu-id="d2d6d-108">EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="d2d6d-108">EquivalentAddin</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="d2d6d-109">外接程序类型值</span><span class="sxs-lookup"><span data-stu-id="d2d6d-109">Add-in type values</span></span>

<span data-ttu-id="d2d6d-110">必须为 元素指定下列值之 `Type` 一。</span><span class="sxs-lookup"><span data-stu-id="d2d6d-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="d2d6d-111">COM：指定等效加载项是 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="d2d6d-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="d2d6d-112">XLL：指定等效加载项是 Excel XLL。</span><span class="sxs-lookup"><span data-stu-id="d2d6d-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="d2d6d-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d2d6d-113">See also</span></span>

- [<span data-ttu-id="d2d6d-114">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="d2d6d-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="d2d6d-115">让 Office 加载项与现有 COM 加载项兼容</span><span class="sxs-lookup"><span data-stu-id="d2d6d-115">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)