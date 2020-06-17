---
title: 清单文件中的 Requirements 元素
description: "\"要求\" 元素指定要激活的 Office 外接程序所需的最低要求集和方法。"
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611713"
---
# <a name="requirements-element"></a><span data-ttu-id="cd058-103">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="cd058-103">Requirements element</span></span>

<span data-ttu-id="cd058-104">指定 Office 外接程序需要激活的最小 Office JavaScript API 要求（[要求集](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)和/或方法）集。</span><span class="sxs-lookup"><span data-stu-id="cd058-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="cd058-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="cd058-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cd058-106">语法</span><span class="sxs-lookup"><span data-stu-id="cd058-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="cd058-107">包含于</span><span class="sxs-lookup"><span data-stu-id="cd058-107">Contained in</span></span>

[<span data-ttu-id="cd058-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cd058-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="cd058-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="cd058-109">Can contain</span></span>

|<span data-ttu-id="cd058-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="cd058-110">**Element**</span></span>|<span data-ttu-id="cd058-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="cd058-111">**Content**</span></span>|<span data-ttu-id="cd058-112">**Mail**</span><span class="sxs-lookup"><span data-stu-id="cd058-112">**Mail**</span></span>|<span data-ttu-id="cd058-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="cd058-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="cd058-114">Sets</span><span class="sxs-lookup"><span data-stu-id="cd058-114">Sets</span></span>](sets.md)|<span data-ttu-id="cd058-115">x</span><span class="sxs-lookup"><span data-stu-id="cd058-115">x</span></span>|<span data-ttu-id="cd058-116">x</span><span class="sxs-lookup"><span data-stu-id="cd058-116">x</span></span>|<span data-ttu-id="cd058-117">x</span><span class="sxs-lookup"><span data-stu-id="cd058-117">x</span></span>|
|[<span data-ttu-id="cd058-118">方法</span><span class="sxs-lookup"><span data-stu-id="cd058-118">Methods</span></span>](methods.md)|<span data-ttu-id="cd058-119">x</span><span class="sxs-lookup"><span data-stu-id="cd058-119">x</span></span>||<span data-ttu-id="cd058-120">x</span><span class="sxs-lookup"><span data-stu-id="cd058-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="cd058-121">注释</span><span class="sxs-lookup"><span data-stu-id="cd058-121">Remarks</span></span>

<span data-ttu-id="cd058-122">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="cd058-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
