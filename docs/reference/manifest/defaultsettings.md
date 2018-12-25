---
title: 清单文件中的 DefaultSettings 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433752"
---
# <a name="defaultsettings-element"></a><span data-ttu-id="5ed05-102">DefaultSettings 元素</span><span class="sxs-lookup"><span data-stu-id="5ed05-102">DefaultSettings element</span></span>

<span data-ttu-id="5ed05-103">指定内容或任务窗格外接程序的默认源位置和其他默认设置。</span><span class="sxs-lookup"><span data-stu-id="5ed05-103">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="5ed05-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="5ed05-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="5ed05-105">语法</span><span class="sxs-lookup"><span data-stu-id="5ed05-105">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="5ed05-106">包含于</span><span class="sxs-lookup"><span data-stu-id="5ed05-106">Contained in</span></span>

[<span data-ttu-id="5ed05-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="5ed05-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="5ed05-108">可以包含</span><span class="sxs-lookup"><span data-stu-id="5ed05-108">Can contain</span></span>

|<span data-ttu-id="5ed05-109">**元素**</span><span class="sxs-lookup"><span data-stu-id="5ed05-109">**Element**</span></span>|<span data-ttu-id="5ed05-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="5ed05-110">**Content**</span></span>|<span data-ttu-id="5ed05-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="5ed05-111">**Mail**</span></span>|<span data-ttu-id="5ed05-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="5ed05-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5ed05-113">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5ed05-113">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="5ed05-114">x</span><span class="sxs-lookup"><span data-stu-id="5ed05-114">x</span></span>||<span data-ttu-id="5ed05-115">x</span><span class="sxs-lookup"><span data-stu-id="5ed05-115">x</span></span>|
|[<span data-ttu-id="5ed05-116">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="5ed05-116">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="5ed05-117">x</span><span class="sxs-lookup"><span data-stu-id="5ed05-117">x</span></span>|||
|[<span data-ttu-id="5ed05-118">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="5ed05-118">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="5ed05-119">x</span><span class="sxs-lookup"><span data-stu-id="5ed05-119">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="5ed05-120">注解</span><span class="sxs-lookup"><span data-stu-id="5ed05-120">Remarks</span></span>

<span data-ttu-id="5ed05-121">**DefaultSettings** 元素中的源位置和其他设置仅应用于内容和任务窗格外接程序。对于邮件外接程序，您在 [FormSettings](formsettings.md) 元素中指定源文件的默认位置和其他默认设置。</span><span class="sxs-lookup"><span data-stu-id="5ed05-121">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

