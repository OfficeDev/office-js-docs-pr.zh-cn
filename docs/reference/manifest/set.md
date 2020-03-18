---
title: 清单文件中的 Set 元素
description: Set 元素指定 office 外接程序需要的 Office JavaScript API 要求集才能激活。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e9a70da0dc38c3aee077eb5e7f47cdf8e6dc2d32
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717913"
---
# <a name="set-element"></a><span data-ttu-id="49f4a-103">Set 元素</span><span class="sxs-lookup"><span data-stu-id="49f4a-103">Set element</span></span>

<span data-ttu-id="49f4a-104">指定 office 外接程序需要激活的 Office JavaScript API 中的要求集。</span><span class="sxs-lookup"><span data-stu-id="49f4a-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="49f4a-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="49f4a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="49f4a-106">语法</span><span class="sxs-lookup"><span data-stu-id="49f4a-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="49f4a-107">包含于</span><span class="sxs-lookup"><span data-stu-id="49f4a-107">Contained in</span></span>

[<span data-ttu-id="49f4a-108">Sets</span><span class="sxs-lookup"><span data-stu-id="49f4a-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="49f4a-109">属性</span><span class="sxs-lookup"><span data-stu-id="49f4a-109">Attributes</span></span>

|<span data-ttu-id="49f4a-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="49f4a-110">**Attribute**</span></span>|<span data-ttu-id="49f4a-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="49f4a-111">**Type**</span></span>|<span data-ttu-id="49f4a-112">**必需**</span><span class="sxs-lookup"><span data-stu-id="49f4a-112">**Required**</span></span>|<span data-ttu-id="49f4a-113">**说明**</span><span class="sxs-lookup"><span data-stu-id="49f4a-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="49f4a-114">名称</span><span class="sxs-lookup"><span data-stu-id="49f4a-114">Name</span></span>|<span data-ttu-id="49f4a-115">string</span><span class="sxs-lookup"><span data-stu-id="49f4a-115">string</span></span>|<span data-ttu-id="49f4a-116">必需</span><span class="sxs-lookup"><span data-stu-id="49f4a-116">required</span></span>|<span data-ttu-id="49f4a-117">[要求集](../../develop/office-versions-and-requirement-sets.md)名称。</span><span class="sxs-lookup"><span data-stu-id="49f4a-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="49f4a-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="49f4a-118">MinVersion</span></span>|<span data-ttu-id="49f4a-119">字符串</span><span class="sxs-lookup"><span data-stu-id="49f4a-119">string</span></span>|<span data-ttu-id="49f4a-120">可选</span><span class="sxs-lookup"><span data-stu-id="49f4a-120">optional</span></span>|<span data-ttu-id="49f4a-121">指定您的外接程序所需的 API 集的最低版本。</span><span class="sxs-lookup"><span data-stu-id="49f4a-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="49f4a-122">重写**DefaultMinVersion**的值（如果它在父[集](sets.md)元素中指定）。</span><span class="sxs-lookup"><span data-stu-id="49f4a-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="49f4a-123">注释</span><span class="sxs-lookup"><span data-stu-id="49f4a-123">Remarks</span></span>

<span data-ttu-id="49f4a-124">有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="49f4a-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="49f4a-125">有关**Set**元素的**MinVersion**属性和**Sets**元素的**DefaultMinVersion**属性的详细信息，请参阅[在清单中设置需求元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="49f4a-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="49f4a-126">对于邮件外接程序，则只能使用一个 `"Mailbox"` 要求集。</span><span class="sxs-lookup"><span data-stu-id="49f4a-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="49f4a-127">此要求集包含 Outlook 邮件外接程序支持的整个 API 子集，你必须在邮件外接程序清单中指定 `"Mailbox"` 要求集（针对内容和任务窗格外接程序，非可选）。</span><span class="sxs-lookup"><span data-stu-id="49f4a-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="49f4a-128">另外，您无法在邮件外接程序中声明对特定方法的支持。</span><span class="sxs-lookup"><span data-stu-id="49f4a-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
