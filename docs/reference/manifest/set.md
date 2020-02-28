---
title: 清单文件中的 Set 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324804"
---
# <a name="set-element"></a><span data-ttu-id="6068a-102">Set 元素</span><span class="sxs-lookup"><span data-stu-id="6068a-102">Set element</span></span>

<span data-ttu-id="6068a-103">指定 office 外接程序需要激活的 Office JavaScript API 中的要求集。</span><span class="sxs-lookup"><span data-stu-id="6068a-103">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="6068a-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="6068a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6068a-105">语法</span><span class="sxs-lookup"><span data-stu-id="6068a-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="6068a-106">包含于</span><span class="sxs-lookup"><span data-stu-id="6068a-106">Contained in</span></span>

[<span data-ttu-id="6068a-107">Sets</span><span class="sxs-lookup"><span data-stu-id="6068a-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="6068a-108">属性</span><span class="sxs-lookup"><span data-stu-id="6068a-108">Attributes</span></span>

|<span data-ttu-id="6068a-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="6068a-109">**Attribute**</span></span>|<span data-ttu-id="6068a-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="6068a-110">**Type**</span></span>|<span data-ttu-id="6068a-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="6068a-111">**Required**</span></span>|<span data-ttu-id="6068a-112">**说明**</span><span class="sxs-lookup"><span data-stu-id="6068a-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6068a-113">名称</span><span class="sxs-lookup"><span data-stu-id="6068a-113">Name</span></span>|<span data-ttu-id="6068a-114">string</span><span class="sxs-lookup"><span data-stu-id="6068a-114">string</span></span>|<span data-ttu-id="6068a-115">必需</span><span class="sxs-lookup"><span data-stu-id="6068a-115">required</span></span>|<span data-ttu-id="6068a-116">[要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。</span><span class="sxs-lookup"><span data-stu-id="6068a-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="6068a-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="6068a-117">MinVersion</span></span>|<span data-ttu-id="6068a-118">字符串</span><span class="sxs-lookup"><span data-stu-id="6068a-118">string</span></span>|<span data-ttu-id="6068a-119">可选</span><span class="sxs-lookup"><span data-stu-id="6068a-119">optional</span></span>|<span data-ttu-id="6068a-120">指定您的外接程序所需的 API 集的最低版本。</span><span class="sxs-lookup"><span data-stu-id="6068a-120">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="6068a-121">重写**DefaultMinVersion**的值（如果它在父[集](sets.md)元素中指定）。</span><span class="sxs-lookup"><span data-stu-id="6068a-121">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6068a-122">注释</span><span class="sxs-lookup"><span data-stu-id="6068a-122">Remarks</span></span>

<span data-ttu-id="6068a-123">有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="6068a-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="6068a-124">有关**Set**元素的**MinVersion**属性和**Sets**元素的**DefaultMinVersion**属性的详细信息，请参阅[在清单中设置需求元素](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="6068a-124">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="6068a-125">对于邮件外接程序，则只能使用一个 `"Mailbox"` 要求集。</span><span class="sxs-lookup"><span data-stu-id="6068a-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="6068a-126">此要求集包含 Outlook 邮件外接程序支持的整个 API 子集，你必须在邮件外接程序清单中指定 `"Mailbox"` 要求集（针对内容和任务窗格外接程序，非可选）。</span><span class="sxs-lookup"><span data-stu-id="6068a-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="6068a-127">另外，您无法在邮件外接程序中声明对特定方法的支持。</span><span class="sxs-lookup"><span data-stu-id="6068a-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
