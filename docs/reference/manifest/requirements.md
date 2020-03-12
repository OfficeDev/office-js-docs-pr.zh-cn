---
title: 清单文件中的 Requirements 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 43c66118b9129c4c8ae395254ea82ef1cbcbaab1
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596457"
---
# <a name="requirements-element"></a><span data-ttu-id="392f1-102">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="392f1-102">Requirements element</span></span>

<span data-ttu-id="392f1-103">指定 Office 外接程序需要激活的最小 Office JavaScript API 要求（[要求集](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)和/或方法）集。</span><span class="sxs-lookup"><span data-stu-id="392f1-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="392f1-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="392f1-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="392f1-105">语法</span><span class="sxs-lookup"><span data-stu-id="392f1-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="392f1-106">包含于</span><span class="sxs-lookup"><span data-stu-id="392f1-106">Contained in</span></span>

[<span data-ttu-id="392f1-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="392f1-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="392f1-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="392f1-108">Can contain</span></span>

|<span data-ttu-id="392f1-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="392f1-109">**Element**</span></span>|<span data-ttu-id="392f1-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="392f1-110">**Content**</span></span>|<span data-ttu-id="392f1-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="392f1-111">**Mail**</span></span>|<span data-ttu-id="392f1-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="392f1-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="392f1-113">Sets</span><span class="sxs-lookup"><span data-stu-id="392f1-113">Sets</span></span>](sets.md)|<span data-ttu-id="392f1-114">x</span><span class="sxs-lookup"><span data-stu-id="392f1-114">x</span></span>|<span data-ttu-id="392f1-115">x</span><span class="sxs-lookup"><span data-stu-id="392f1-115">x</span></span>|<span data-ttu-id="392f1-116">x</span><span class="sxs-lookup"><span data-stu-id="392f1-116">x</span></span>|
|[<span data-ttu-id="392f1-117">方法</span><span class="sxs-lookup"><span data-stu-id="392f1-117">Methods</span></span>](methods.md)|<span data-ttu-id="392f1-118">x</span><span class="sxs-lookup"><span data-stu-id="392f1-118">x</span></span>||<span data-ttu-id="392f1-119">x</span><span class="sxs-lookup"><span data-stu-id="392f1-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="392f1-120">注释</span><span class="sxs-lookup"><span data-stu-id="392f1-120">Remarks</span></span>

<span data-ttu-id="392f1-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="392f1-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
