---
title: 清单文件中的 Type 元素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356857"
---
# <a name="type-element"></a><span data-ttu-id="7f6de-102">Type 元素</span><span class="sxs-lookup"><span data-stu-id="7f6de-102">Type element</span></span>

<span data-ttu-id="7f6de-103">指定等效加载项是 COM 外接程序还是 XLL。</span><span class="sxs-lookup"><span data-stu-id="7f6de-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="7f6de-104">**外接类型:** 任务窗格, 自定义函数</span><span class="sxs-lookup"><span data-stu-id="7f6de-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="7f6de-105">语法</span><span class="sxs-lookup"><span data-stu-id="7f6de-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="7f6de-106">包含于</span><span class="sxs-lookup"><span data-stu-id="7f6de-106">Contained in</span></span>

[<span data-ttu-id="7f6de-107">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="7f6de-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="7f6de-108">外接类型值</span><span class="sxs-lookup"><span data-stu-id="7f6de-108">Add-in type values</span></span>

<span data-ttu-id="7f6de-109">必须为`Type`元素指定下列值之一。</span><span class="sxs-lookup"><span data-stu-id="7f6de-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="7f6de-110">com: 指定等效的加载项是 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="7f6de-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="7f6de-111">XLL: 指定等效的外接程序是 Excel XLL。</span><span class="sxs-lookup"><span data-stu-id="7f6de-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="7f6de-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7f6de-112">See also</span></span>

- [<span data-ttu-id="7f6de-113">使自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="7f6de-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="7f6de-114">使您的 Office 外接程序与现有的 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="7f6de-114">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)