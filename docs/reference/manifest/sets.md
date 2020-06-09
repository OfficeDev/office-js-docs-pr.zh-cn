---
title: 清单文件中的 Sets 元素
description: Set 元素指定 Office 外接程序在激活时所需的最小 Office JavaScript API 集。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8c1c97bfc2934ecf3cc20b472b29a03805603729
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608731"
---
# <a name="sets-element"></a><span data-ttu-id="7ac58-103">Sets 元素</span><span class="sxs-lookup"><span data-stu-id="7ac58-103">Sets element</span></span>

<span data-ttu-id="7ac58-104">指定 Office JavaScript API 的最小子集，Office 外接程序需要这些 API 才能激活。</span><span class="sxs-lookup"><span data-stu-id="7ac58-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="7ac58-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="7ac58-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7ac58-106">语法</span><span class="sxs-lookup"><span data-stu-id="7ac58-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="7ac58-107">包含于</span><span class="sxs-lookup"><span data-stu-id="7ac58-107">Contained in</span></span>

[<span data-ttu-id="7ac58-108">要求</span><span class="sxs-lookup"><span data-stu-id="7ac58-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="7ac58-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="7ac58-109">Can contain</span></span>

[<span data-ttu-id="7ac58-110">Set</span><span class="sxs-lookup"><span data-stu-id="7ac58-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="7ac58-111">属性</span><span class="sxs-lookup"><span data-stu-id="7ac58-111">Attributes</span></span>

|<span data-ttu-id="7ac58-112">**属性**</span><span class="sxs-lookup"><span data-stu-id="7ac58-112">**Attribute**</span></span>|<span data-ttu-id="7ac58-113">**类型**</span><span class="sxs-lookup"><span data-stu-id="7ac58-113">**Type**</span></span>|<span data-ttu-id="7ac58-114">**必需**</span><span class="sxs-lookup"><span data-stu-id="7ac58-114">**Required**</span></span>|<span data-ttu-id="7ac58-115">**描述**</span><span class="sxs-lookup"><span data-stu-id="7ac58-115">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7ac58-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="7ac58-116">DefaultMinVersion</span></span>|<span data-ttu-id="7ac58-117">字符串</span><span class="sxs-lookup"><span data-stu-id="7ac58-117">string</span></span>|<span data-ttu-id="7ac58-118">可选</span><span class="sxs-lookup"><span data-stu-id="7ac58-118">optional</span></span>|<span data-ttu-id="7ac58-119">指定所有子[集](set.md)元素的默认**MinVersion**属性值。</span><span class="sxs-lookup"><span data-stu-id="7ac58-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="7ac58-120">默认值为“1.1”。</span><span class="sxs-lookup"><span data-stu-id="7ac58-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="7ac58-121">注释</span><span class="sxs-lookup"><span data-stu-id="7ac58-121">Remarks</span></span>

<span data-ttu-id="7ac58-122">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="7ac58-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="7ac58-123">有关**Set**元素的**MinVersion**属性和**Sets**元素的**DefaultMinVersion**属性的详细信息，请参阅[在清单中设置需求元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="7ac58-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

