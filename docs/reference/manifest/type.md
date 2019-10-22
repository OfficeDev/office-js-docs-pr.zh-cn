---
title: 清单文件中的 Type 元素
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628226"
---
# <a name="type-element"></a><span data-ttu-id="6db15-102">Type 元素</span><span class="sxs-lookup"><span data-stu-id="6db15-102">Type element</span></span>

<span data-ttu-id="6db15-103">指定等效加载项是 COM 外接程序还是 XLL。</span><span class="sxs-lookup"><span data-stu-id="6db15-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="6db15-104">**外接类型：** 任务窗格，自定义函数</span><span class="sxs-lookup"><span data-stu-id="6db15-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="6db15-105">语法</span><span class="sxs-lookup"><span data-stu-id="6db15-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="6db15-106">包含于</span><span class="sxs-lookup"><span data-stu-id="6db15-106">Contained in</span></span>

[<span data-ttu-id="6db15-107">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="6db15-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="6db15-108">外接类型值</span><span class="sxs-lookup"><span data-stu-id="6db15-108">Add-in type values</span></span>

<span data-ttu-id="6db15-109">必须为`Type`元素指定下列值之一。</span><span class="sxs-lookup"><span data-stu-id="6db15-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="6db15-110">COM：指定等效的加载项是 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="6db15-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="6db15-111">XLL：指定等效的外接程序是 Excel XLL。</span><span class="sxs-lookup"><span data-stu-id="6db15-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="6db15-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6db15-112">See also</span></span>

- [<span data-ttu-id="6db15-113">让自定义功能与 XLL 用户定义的功能兼容</span><span class="sxs-lookup"><span data-stu-id="6db15-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="6db15-114">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="6db15-114">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)