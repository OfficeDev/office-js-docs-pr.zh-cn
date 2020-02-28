---
title: 清单文件中的 Requirements 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3c4cb81ebd6a38ea311e8fcacfa6d5fcd3b26f68
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325246"
---
# <a name="requirements-element"></a><span data-ttu-id="39955-102">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="39955-102">Requirements element</span></span>

<span data-ttu-id="39955-103">指定 Office 外接程序需要激活的最小 Office JavaScript API 要求（[要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或方法）集。</span><span class="sxs-lookup"><span data-stu-id="39955-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="39955-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="39955-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="39955-105">语法</span><span class="sxs-lookup"><span data-stu-id="39955-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="39955-106">包含于</span><span class="sxs-lookup"><span data-stu-id="39955-106">Contained in</span></span>

[<span data-ttu-id="39955-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="39955-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="39955-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="39955-108">Can contain</span></span>

|<span data-ttu-id="39955-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="39955-109">**Element**</span></span>|<span data-ttu-id="39955-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="39955-110">**Content**</span></span>|<span data-ttu-id="39955-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="39955-111">**Mail**</span></span>|<span data-ttu-id="39955-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="39955-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="39955-113">Sets</span><span class="sxs-lookup"><span data-stu-id="39955-113">Sets</span></span>](sets.md)|<span data-ttu-id="39955-114">x</span><span class="sxs-lookup"><span data-stu-id="39955-114">x</span></span>|<span data-ttu-id="39955-115">x</span><span class="sxs-lookup"><span data-stu-id="39955-115">x</span></span>|<span data-ttu-id="39955-116">x</span><span class="sxs-lookup"><span data-stu-id="39955-116">x</span></span>|
|[<span data-ttu-id="39955-117">方法</span><span class="sxs-lookup"><span data-stu-id="39955-117">Methods</span></span>](methods.md)|<span data-ttu-id="39955-118">x</span><span class="sxs-lookup"><span data-stu-id="39955-118">x</span></span>||<span data-ttu-id="39955-119">x</span><span class="sxs-lookup"><span data-stu-id="39955-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="39955-120">注释</span><span class="sxs-lookup"><span data-stu-id="39955-120">Remarks</span></span>

<span data-ttu-id="39955-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="39955-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

