---
title: 清单文件中的 Sets 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 80f8a74b64186496ac1579b283b3e2976978328b
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596485"
---
# <a name="sets-element"></a><span data-ttu-id="3e51e-102">Sets 元素</span><span class="sxs-lookup"><span data-stu-id="3e51e-102">Sets element</span></span>

<span data-ttu-id="3e51e-103">指定 Office JavaScript API 的最小子集，Office 外接程序需要这些 API 才能激活。</span><span class="sxs-lookup"><span data-stu-id="3e51e-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="3e51e-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="3e51e-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3e51e-105">语法</span><span class="sxs-lookup"><span data-stu-id="3e51e-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="3e51e-106">包含于</span><span class="sxs-lookup"><span data-stu-id="3e51e-106">Contained in</span></span>

[<span data-ttu-id="3e51e-107">要求</span><span class="sxs-lookup"><span data-stu-id="3e51e-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="3e51e-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="3e51e-108">Can contain</span></span>

[<span data-ttu-id="3e51e-109">Set</span><span class="sxs-lookup"><span data-stu-id="3e51e-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="3e51e-110">属性</span><span class="sxs-lookup"><span data-stu-id="3e51e-110">Attributes</span></span>

|<span data-ttu-id="3e51e-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="3e51e-111">**Attribute**</span></span>|<span data-ttu-id="3e51e-112">**类型**</span><span class="sxs-lookup"><span data-stu-id="3e51e-112">**Type**</span></span>|<span data-ttu-id="3e51e-113">**必需**</span><span class="sxs-lookup"><span data-stu-id="3e51e-113">**Required**</span></span>|<span data-ttu-id="3e51e-114">**描述**</span><span class="sxs-lookup"><span data-stu-id="3e51e-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3e51e-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="3e51e-115">DefaultMinVersion</span></span>|<span data-ttu-id="3e51e-116">字符串</span><span class="sxs-lookup"><span data-stu-id="3e51e-116">string</span></span>|<span data-ttu-id="3e51e-117">可选</span><span class="sxs-lookup"><span data-stu-id="3e51e-117">optional</span></span>|<span data-ttu-id="3e51e-118">指定所有子[集](set.md)元素的默认**MinVersion**属性值。</span><span class="sxs-lookup"><span data-stu-id="3e51e-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="3e51e-119">默认值为“1.1”。</span><span class="sxs-lookup"><span data-stu-id="3e51e-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="3e51e-120">注释</span><span class="sxs-lookup"><span data-stu-id="3e51e-120">Remarks</span></span>

<span data-ttu-id="3e51e-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3e51e-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3e51e-122">有关**Set**元素的**MinVersion**属性和**Sets**元素的**DefaultMinVersion**属性的详细信息，请参阅[在清单中设置需求元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="3e51e-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

