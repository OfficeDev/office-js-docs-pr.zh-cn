---
title: 清单文件中的 Type 元素
description: Type 元素指定等效加载项是 COM 加载项还是 XLL。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604557"
---
# <a name="type-element"></a><span data-ttu-id="b617e-103">Type 元素</span><span class="sxs-lookup"><span data-stu-id="b617e-103">Type element</span></span>

<span data-ttu-id="b617e-104">指定等效的外接程序是 COM 加载项还是 XLL。</span><span class="sxs-lookup"><span data-stu-id="b617e-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="b617e-105">**外接类型：** 任务窗格，自定义函数</span><span class="sxs-lookup"><span data-stu-id="b617e-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="b617e-106">语法</span><span class="sxs-lookup"><span data-stu-id="b617e-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="b617e-107">包含于</span><span class="sxs-lookup"><span data-stu-id="b617e-107">Contained in</span></span>

[<span data-ttu-id="b617e-108">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="b617e-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="b617e-109">外接类型值</span><span class="sxs-lookup"><span data-stu-id="b617e-109">Add-in type values</span></span>

<span data-ttu-id="b617e-110">必须为元素指定下列值之一 `Type` 。</span><span class="sxs-lookup"><span data-stu-id="b617e-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="b617e-111">COM：指定等效的加载项是 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="b617e-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="b617e-112">XLL：指定等效的外接程序是 Excel XLL。</span><span class="sxs-lookup"><span data-stu-id="b617e-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="b617e-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b617e-113">See also</span></span>

- [<span data-ttu-id="b617e-114">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="b617e-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="b617e-115">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="b617e-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)