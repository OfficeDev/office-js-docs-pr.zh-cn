---
title: 清单文件中的 Requirements 元素
description: "\"要求\" 元素指定要激活的 Office 外接程序所需的最低要求集和方法。"
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a3f41a763ec820a6c766e6a32b26e55ad34996f7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720447"
---
# <a name="requirements-element"></a><span data-ttu-id="79727-103">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="79727-103">Requirements element</span></span>

<span data-ttu-id="79727-104">指定 Office 外接程序需要激活的最小 Office JavaScript API 要求（[要求集](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)和/或方法）集。</span><span class="sxs-lookup"><span data-stu-id="79727-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="79727-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="79727-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="79727-106">语法</span><span class="sxs-lookup"><span data-stu-id="79727-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="79727-107">包含于</span><span class="sxs-lookup"><span data-stu-id="79727-107">Contained in</span></span>

[<span data-ttu-id="79727-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="79727-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="79727-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="79727-109">Can contain</span></span>

|<span data-ttu-id="79727-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="79727-110">**Element**</span></span>|<span data-ttu-id="79727-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="79727-111">**Content**</span></span>|<span data-ttu-id="79727-112">**Mail**</span><span class="sxs-lookup"><span data-stu-id="79727-112">**Mail**</span></span>|<span data-ttu-id="79727-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="79727-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="79727-114">Sets</span><span class="sxs-lookup"><span data-stu-id="79727-114">Sets</span></span>](sets.md)|<span data-ttu-id="79727-115">x</span><span class="sxs-lookup"><span data-stu-id="79727-115">x</span></span>|<span data-ttu-id="79727-116">x</span><span class="sxs-lookup"><span data-stu-id="79727-116">x</span></span>|<span data-ttu-id="79727-117">x</span><span class="sxs-lookup"><span data-stu-id="79727-117">x</span></span>|
|[<span data-ttu-id="79727-118">方法</span><span class="sxs-lookup"><span data-stu-id="79727-118">Methods</span></span>](methods.md)|<span data-ttu-id="79727-119">x</span><span class="sxs-lookup"><span data-stu-id="79727-119">x</span></span>||<span data-ttu-id="79727-120">x</span><span class="sxs-lookup"><span data-stu-id="79727-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="79727-121">注释</span><span class="sxs-lookup"><span data-stu-id="79727-121">Remarks</span></span>

<span data-ttu-id="79727-122">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="79727-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
