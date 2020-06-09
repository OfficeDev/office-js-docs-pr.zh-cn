---
title: 清单文件中的 RequestedHeight 元素
description: RequestedHeight 元素指定内容或邮件加载项的初始高度（以像素为单位）。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 44675918a4208683f442fe8a6e8f4f906f484571
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611727"
---
# <a name="requestedheight-element"></a><span data-ttu-id="0b0b6-103">RequestedHeight 元素</span><span class="sxs-lookup"><span data-stu-id="0b0b6-103">RequestedHeight element</span></span>

<span data-ttu-id="0b0b6-104">指定内容外接程序或邮件外接程序的初始高度（以像素为单位）。</span><span class="sxs-lookup"><span data-stu-id="0b0b6-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="0b0b6-105">**外接程序类型：** 内容、邮件</span><span class="sxs-lookup"><span data-stu-id="0b0b6-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0b0b6-106">语法</span><span class="sxs-lookup"><span data-stu-id="0b0b6-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="0b0b6-107">包含于</span><span class="sxs-lookup"><span data-stu-id="0b0b6-107">Contained in</span></span>

- <span data-ttu-id="0b0b6-108">[DefaultSettings](defaultsettings.md)（内容外接程序）：值可以在 32 至 1000 之间</span><span class="sxs-lookup"><span data-stu-id="0b0b6-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="0b0b6-109">[DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md) （邮件外接程序）：值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="0b0b6-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="0b0b6-110">[ExtensionPoint](extensionpoint.md) （上下文邮件外接程序），其值可以介于140和450之间的**DetectedEntity**扩展点，在32和450之间为[ **CustomPane**扩展点（已弃用）](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span><span class="sxs-lookup"><span data-stu-id="0b0b6-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>
