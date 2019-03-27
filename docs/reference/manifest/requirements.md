---
title: 清单文件中的 Requirements 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870364"
---
# <a name="requirements-element"></a><span data-ttu-id="a39f7-102">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="a39f7-102">Requirements element</span></span>

<span data-ttu-id="a39f7-103">指定适用于 Office 的 JavaScript API 要求（[要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或方法）的最小集，Office 外接程序需要该集才能激活。</span><span class="sxs-lookup"><span data-stu-id="a39f7-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="a39f7-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="a39f7-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a39f7-105">语法</span><span class="sxs-lookup"><span data-stu-id="a39f7-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="a39f7-106">包含于</span><span class="sxs-lookup"><span data-stu-id="a39f7-106">Contained in</span></span>

[<span data-ttu-id="a39f7-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a39f7-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="a39f7-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="a39f7-108">Can contain</span></span>

|<span data-ttu-id="a39f7-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="a39f7-109">**Element**</span></span>|<span data-ttu-id="a39f7-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="a39f7-110">**Content**</span></span>|<span data-ttu-id="a39f7-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="a39f7-111">**Mail**</span></span>|<span data-ttu-id="a39f7-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="a39f7-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a39f7-113">Sets</span><span class="sxs-lookup"><span data-stu-id="a39f7-113">Sets</span></span>](sets.md)|<span data-ttu-id="a39f7-114">x</span><span class="sxs-lookup"><span data-stu-id="a39f7-114">x</span></span>|<span data-ttu-id="a39f7-115">x</span><span class="sxs-lookup"><span data-stu-id="a39f7-115">x</span></span>|<span data-ttu-id="a39f7-116">x</span><span class="sxs-lookup"><span data-stu-id="a39f7-116">x</span></span>|
|[<span data-ttu-id="a39f7-117">方法</span><span class="sxs-lookup"><span data-stu-id="a39f7-117">Methods</span></span>](methods.md)|<span data-ttu-id="a39f7-118">x</span><span class="sxs-lookup"><span data-stu-id="a39f7-118">x</span></span>||<span data-ttu-id="a39f7-119">x</span><span class="sxs-lookup"><span data-stu-id="a39f7-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="a39f7-120">注释</span><span class="sxs-lookup"><span data-stu-id="a39f7-120">Remarks</span></span>

<span data-ttu-id="a39f7-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="a39f7-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

