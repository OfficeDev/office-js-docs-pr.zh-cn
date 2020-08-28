---
title: 清单文件中的 Requirements 元素
description: "\"要求\" 元素指定要激活的 Office 外接程序所需的最低要求集和方法。"
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292270"
---
# <a name="requirements-element"></a><span data-ttu-id="35f2c-103">Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="35f2c-103">Requirements element</span></span>

<span data-ttu-id="35f2c-104">指定 office 外接程序) 需要激活的 Office JavaScript API 要求的最小集合 ([要求集](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) 和/或方法。</span><span class="sxs-lookup"><span data-stu-id="35f2c-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="35f2c-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="35f2c-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="35f2c-106">语法</span><span class="sxs-lookup"><span data-stu-id="35f2c-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="35f2c-107">包含于</span><span class="sxs-lookup"><span data-stu-id="35f2c-107">Contained in</span></span>

[<span data-ttu-id="35f2c-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="35f2c-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="35f2c-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="35f2c-109">Can contain</span></span>

|<span data-ttu-id="35f2c-110">元素</span><span class="sxs-lookup"><span data-stu-id="35f2c-110">Element</span></span>|<span data-ttu-id="35f2c-111">内容</span><span class="sxs-lookup"><span data-stu-id="35f2c-111">Content</span></span>|<span data-ttu-id="35f2c-112">邮件</span><span class="sxs-lookup"><span data-stu-id="35f2c-112">Mail</span></span>|<span data-ttu-id="35f2c-113">任务窗格</span><span class="sxs-lookup"><span data-stu-id="35f2c-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="35f2c-114">Sets</span><span class="sxs-lookup"><span data-stu-id="35f2c-114">Sets</span></span>](sets.md)|<span data-ttu-id="35f2c-115">x</span><span class="sxs-lookup"><span data-stu-id="35f2c-115">x</span></span>|<span data-ttu-id="35f2c-116">x</span><span class="sxs-lookup"><span data-stu-id="35f2c-116">x</span></span>|<span data-ttu-id="35f2c-117">x</span><span class="sxs-lookup"><span data-stu-id="35f2c-117">x</span></span>|
|[<span data-ttu-id="35f2c-118">方法</span><span class="sxs-lookup"><span data-stu-id="35f2c-118">Methods</span></span>](methods.md)|<span data-ttu-id="35f2c-119">x</span><span class="sxs-lookup"><span data-stu-id="35f2c-119">x</span></span>||<span data-ttu-id="35f2c-120">x</span><span class="sxs-lookup"><span data-stu-id="35f2c-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="35f2c-121">注释</span><span class="sxs-lookup"><span data-stu-id="35f2c-121">Remarks</span></span>

<span data-ttu-id="35f2c-122">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="35f2c-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
