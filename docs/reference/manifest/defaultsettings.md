---
title: 清单文件中的 DefaultSettings 元素
description: 指定内容或任务窗格外接程序的默认源位置和其他默认设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ace4f971d342f98d0aca5c21a7a48ceaf2563a2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611580"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="ad2ac-103">DefaultSettings 元素</span><span class="sxs-lookup"><span data-stu-id="ad2ac-103">DefaultSettings element</span></span>

<span data-ttu-id="ad2ac-104">指定内容或任务窗格外接程序的默认源位置和其他默认设置。</span><span class="sxs-lookup"><span data-stu-id="ad2ac-104">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="ad2ac-105">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="ad2ac-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="ad2ac-106">语法</span><span class="sxs-lookup"><span data-stu-id="ad2ac-106">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="ad2ac-107">包含于</span><span class="sxs-lookup"><span data-stu-id="ad2ac-107">Contained in</span></span>

[<span data-ttu-id="ad2ac-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="ad2ac-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="ad2ac-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="ad2ac-109">Can contain</span></span>

|<span data-ttu-id="ad2ac-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="ad2ac-110">**Element**</span></span>|<span data-ttu-id="ad2ac-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="ad2ac-111">**Content**</span></span>|<span data-ttu-id="ad2ac-112">**Mail**</span><span class="sxs-lookup"><span data-stu-id="ad2ac-112">**Mail**</span></span>|<span data-ttu-id="ad2ac-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="ad2ac-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="ad2ac-114">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ad2ac-114">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="ad2ac-115">x</span><span class="sxs-lookup"><span data-stu-id="ad2ac-115">x</span></span>||<span data-ttu-id="ad2ac-116">x</span><span class="sxs-lookup"><span data-stu-id="ad2ac-116">x</span></span>|
|[<span data-ttu-id="ad2ac-117">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="ad2ac-117">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="ad2ac-118">x</span><span class="sxs-lookup"><span data-stu-id="ad2ac-118">x</span></span>|||
|[<span data-ttu-id="ad2ac-119">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="ad2ac-119">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="ad2ac-120">x</span><span class="sxs-lookup"><span data-stu-id="ad2ac-120">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="ad2ac-121">注解</span><span class="sxs-lookup"><span data-stu-id="ad2ac-121">Remarks</span></span>

<span data-ttu-id="ad2ac-122">源位置和**DefaultSettings**元素中的其他设置仅适用于内容和任务窗格外接程序。对于邮件外接程序，您可以在[FormSettings](formsettings.md)元素中指定源文件和其他默认设置的默认位置。</span><span class="sxs-lookup"><span data-stu-id="ad2ac-122">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

