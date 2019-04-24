---
title: 清单文件中的 RequestedHeight 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e175d9012bb2f2a42fd466c35e5e28ade967d6f2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450525"
---
# <a name="requestedheight-element"></a><span data-ttu-id="1f463-102">RequestedHeight 元素</span><span class="sxs-lookup"><span data-stu-id="1f463-102">RequestedHeight element</span></span>

<span data-ttu-id="1f463-103">指定内容外接程序或邮件外接程序的初始高度（以像素为单位）。</span><span class="sxs-lookup"><span data-stu-id="1f463-103">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="1f463-104">**外接程序类型：** 内容、邮件</span><span class="sxs-lookup"><span data-stu-id="1f463-104">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1f463-105">语法</span><span class="sxs-lookup"><span data-stu-id="1f463-105">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="1f463-106">包含于</span><span class="sxs-lookup"><span data-stu-id="1f463-106">Contained in</span></span>

- <span data-ttu-id="1f463-107">[DefaultSettings](defaultsettings.md)（内容外接程序）：值可以在 32 至 1000 之间</span><span class="sxs-lookup"><span data-stu-id="1f463-107">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="1f463-108">[DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md) （邮件外接程序）：值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="1f463-108">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="1f463-109">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）：如果是 **DetectedEntity** 扩展点，值可以在 140 至 450 之间，如果是 **CustomPane** 扩展点，值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="1f463-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
