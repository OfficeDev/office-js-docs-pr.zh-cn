---
title: 清单文件中的 Type 元素
description: Type 元素指定等效加载项是 COM 加载项还是 XLL。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720314"
---
# <a name="type-element"></a><span data-ttu-id="157d4-103">Type 元素</span><span class="sxs-lookup"><span data-stu-id="157d4-103">Type element</span></span>

<span data-ttu-id="157d4-104">指定等效的外接程序是 COM 加载项还是 XLL。</span><span class="sxs-lookup"><span data-stu-id="157d4-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="157d4-105">**外接类型：** 任务窗格，自定义函数</span><span class="sxs-lookup"><span data-stu-id="157d4-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="157d4-106">语法</span><span class="sxs-lookup"><span data-stu-id="157d4-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="157d4-107">包含于</span><span class="sxs-lookup"><span data-stu-id="157d4-107">Contained in</span></span>

[<span data-ttu-id="157d4-108">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="157d4-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="157d4-109">外接类型值</span><span class="sxs-lookup"><span data-stu-id="157d4-109">Add-in type values</span></span>

<span data-ttu-id="157d4-110">必须为`Type`元素指定下列值之一。</span><span class="sxs-lookup"><span data-stu-id="157d4-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="157d4-111">COM：指定等效的加载项是 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="157d4-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="157d4-112">XLL：指定等效的外接程序是 Excel XLL。</span><span class="sxs-lookup"><span data-stu-id="157d4-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="157d4-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="157d4-113">See also</span></span>

- [<span data-ttu-id="157d4-114">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="157d4-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="157d4-115">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="157d4-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)