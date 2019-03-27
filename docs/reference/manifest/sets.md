---
title: 清单文件中的 Sets 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 13777e54ec6bd2d97fa35609ebe194ed85ffa1b8
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871771"
---
# <a name="sets-element"></a><span data-ttu-id="1b08d-102">Sets 元素</span><span class="sxs-lookup"><span data-stu-id="1b08d-102">Sets element</span></span>

<span data-ttu-id="1b08d-103">指定适用于 Office 的 JavaScript API 的最小子集，Office 外接程序需要该子集才能激活。</span><span class="sxs-lookup"><span data-stu-id="1b08d-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="1b08d-104">**外接程序类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="1b08d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1b08d-105">语法</span><span class="sxs-lookup"><span data-stu-id="1b08d-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="1b08d-106">包含于</span><span class="sxs-lookup"><span data-stu-id="1b08d-106">Contained in</span></span>

[<span data-ttu-id="1b08d-107">要求</span><span class="sxs-lookup"><span data-stu-id="1b08d-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="1b08d-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="1b08d-108">Can contain</span></span>

[<span data-ttu-id="1b08d-109">Set</span><span class="sxs-lookup"><span data-stu-id="1b08d-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="1b08d-110">属性</span><span class="sxs-lookup"><span data-stu-id="1b08d-110">Attributes</span></span>

|<span data-ttu-id="1b08d-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="1b08d-111">**Attribute**</span></span>|<span data-ttu-id="1b08d-112">**类型**</span><span class="sxs-lookup"><span data-stu-id="1b08d-112">**Type**</span></span>|<span data-ttu-id="1b08d-113">**必需**</span><span class="sxs-lookup"><span data-stu-id="1b08d-113">**Required**</span></span>|<span data-ttu-id="1b08d-114">**描述**</span><span class="sxs-lookup"><span data-stu-id="1b08d-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1b08d-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="1b08d-115">DefaultMinVersion</span></span>|<span data-ttu-id="1b08d-116">字符串</span><span class="sxs-lookup"><span data-stu-id="1b08d-116">string</span></span>|<span data-ttu-id="1b08d-117">可选</span><span class="sxs-lookup"><span data-stu-id="1b08d-117">optional</span></span>|<span data-ttu-id="1b08d-p101">为所有子 **Set** 元素指定默认的 [MinVersion](set.md) 属性值。默认值为“1.1”。</span><span class="sxs-lookup"><span data-stu-id="1b08d-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="1b08d-120">注释</span><span class="sxs-lookup"><span data-stu-id="1b08d-120">Remarks</span></span>

<span data-ttu-id="1b08d-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="1b08d-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="1b08d-122">有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="1b08d-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

