---
title: 清单文件中的 Sets 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433654"
---
# <a name="sets-element"></a><span data-ttu-id="48f2a-102">Sets 元素</span><span class="sxs-lookup"><span data-stu-id="48f2a-102">Sets element</span></span>

<span data-ttu-id="48f2a-103">指定适用于 Office 的 JavaScript API 的最小子集，Office 外接程序需要该子集才能激活。</span><span class="sxs-lookup"><span data-stu-id="48f2a-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="48f2a-104">**外接程序类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="48f2a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="48f2a-105">语法</span><span class="sxs-lookup"><span data-stu-id="48f2a-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="48f2a-106">包含于</span><span class="sxs-lookup"><span data-stu-id="48f2a-106">Contained in</span></span>

[<span data-ttu-id="48f2a-107">要求</span><span class="sxs-lookup"><span data-stu-id="48f2a-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="48f2a-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="48f2a-108">Can contain</span></span>

[<span data-ttu-id="48f2a-109">Set</span><span class="sxs-lookup"><span data-stu-id="48f2a-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="48f2a-110">属性</span><span class="sxs-lookup"><span data-stu-id="48f2a-110">Attributes</span></span>

|<span data-ttu-id="48f2a-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="48f2a-111">**Attribute**</span></span>|<span data-ttu-id="48f2a-112">**类型**</span><span class="sxs-lookup"><span data-stu-id="48f2a-112">**Type**</span></span>|<span data-ttu-id="48f2a-113">**必需**</span><span class="sxs-lookup"><span data-stu-id="48f2a-113">**Required**</span></span>|<span data-ttu-id="48f2a-114">**说明**</span><span class="sxs-lookup"><span data-stu-id="48f2a-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="48f2a-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="48f2a-115">DefaultMinVersion</span></span>|<span data-ttu-id="48f2a-116">字符串</span><span class="sxs-lookup"><span data-stu-id="48f2a-116">string</span></span>|<span data-ttu-id="48f2a-117">可选</span><span class="sxs-lookup"><span data-stu-id="48f2a-117">optional</span></span>|<span data-ttu-id="48f2a-p101">为所有子 **Set** 元素指定默认的 [MinVersion](set.md) 属性值。默认值为“1.1”。</span><span class="sxs-lookup"><span data-stu-id="48f2a-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="48f2a-120">注释</span><span class="sxs-lookup"><span data-stu-id="48f2a-120">Remarks</span></span>

<span data-ttu-id="48f2a-121">有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="48f2a-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="48f2a-122">有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="48f2a-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

