---
title: 清单文件中的 DefaultSettings 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 824c575b39a99c6028ffd603390d2b41ee0ad7dd
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324882"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="3c84f-102">DefaultSettings 元素</span><span class="sxs-lookup"><span data-stu-id="3c84f-102">DefaultSettings element</span></span>

<span data-ttu-id="3c84f-103">指定内容或任务窗格外接程序的默认源位置和其他默认设置。</span><span class="sxs-lookup"><span data-stu-id="3c84f-103">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="3c84f-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="3c84f-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="3c84f-105">语法</span><span class="sxs-lookup"><span data-stu-id="3c84f-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="3c84f-106">包含于</span><span class="sxs-lookup"><span data-stu-id="3c84f-106">Contained in</span></span>

[<span data-ttu-id="3c84f-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3c84f-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="3c84f-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="3c84f-108">Can contain</span></span>

|<span data-ttu-id="3c84f-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="3c84f-109">**Element**</span></span>|<span data-ttu-id="3c84f-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="3c84f-110">**Content**</span></span>|<span data-ttu-id="3c84f-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="3c84f-111">**Mail**</span></span>|<span data-ttu-id="3c84f-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="3c84f-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3c84f-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3c84f-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="3c84f-114">x</span><span class="sxs-lookup"><span data-stu-id="3c84f-114">x</span></span>||<span data-ttu-id="3c84f-115">x</span><span class="sxs-lookup"><span data-stu-id="3c84f-115">x</span></span>|
|[<span data-ttu-id="3c84f-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="3c84f-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="3c84f-117">x</span><span class="sxs-lookup"><span data-stu-id="3c84f-117">x</span></span>|||
|[<span data-ttu-id="3c84f-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="3c84f-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="3c84f-119">x</span><span class="sxs-lookup"><span data-stu-id="3c84f-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="3c84f-120">注解</span><span class="sxs-lookup"><span data-stu-id="3c84f-120">Remarks</span></span>

<span data-ttu-id="3c84f-121">源位置和**DefaultSettings**元素中的其他设置仅适用于内容和任务窗格外接程序。对于邮件外接程序，您可以在[FormSettings](formsettings.md)元素中指定源文件和其他默认设置的默认位置。</span><span class="sxs-lookup"><span data-stu-id="3c84f-121">The source location and other settings in the **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

