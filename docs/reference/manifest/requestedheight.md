---
title: 清单文件中的 RequestedHeight 元素
description: RequestedHeight 元素指定内容或邮件加载项的初始高度（以像素为单位）。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 853d12baf290167f3e6a635201e8b5d1d0e35a51
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720454"
---
# <a name="requestedheight-element"></a><span data-ttu-id="38098-103">RequestedHeight 元素</span><span class="sxs-lookup"><span data-stu-id="38098-103">RequestedHeight element</span></span>

<span data-ttu-id="38098-104">指定内容外接程序或邮件外接程序的初始高度（以像素为单位）。</span><span class="sxs-lookup"><span data-stu-id="38098-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="38098-105">**外接程序类型：** 内容、邮件</span><span class="sxs-lookup"><span data-stu-id="38098-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="38098-106">语法</span><span class="sxs-lookup"><span data-stu-id="38098-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="38098-107">包含于</span><span class="sxs-lookup"><span data-stu-id="38098-107">Contained in</span></span>

- <span data-ttu-id="38098-108">[DefaultSettings](defaultsettings.md)（内容外接程序）：值可以在 32 至 1000 之间</span><span class="sxs-lookup"><span data-stu-id="38098-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="38098-109">[DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md) （邮件外接程序）：值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="38098-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="38098-110">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）：如果是 **DetectedEntity** 扩展点，值可以在 140 至 450 之间，如果是 **CustomPane** 扩展点，值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="38098-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
