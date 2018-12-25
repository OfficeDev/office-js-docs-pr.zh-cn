---
title: 清单文件中的 RequestedHeight 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ea8c0403146f526b28eb20b8364fd210ac357baf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433472"
---
# <a name="requestedheight-element"></a><span data-ttu-id="244ea-102">RequestedHeight 元素</span><span class="sxs-lookup"><span data-stu-id="244ea-102">RequestedHeight element</span></span>

<span data-ttu-id="244ea-103">指定内容外接程序或邮件外接程序的初始高度（以像素为单位）。</span><span class="sxs-lookup"><span data-stu-id="244ea-103">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="244ea-104">**外接程序类型：** 内容、邮件</span><span class="sxs-lookup"><span data-stu-id="244ea-104">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="244ea-105">语法</span><span class="sxs-lookup"><span data-stu-id="244ea-105">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="244ea-106">包含于</span><span class="sxs-lookup"><span data-stu-id="244ea-106">Contained in</span></span>

- <span data-ttu-id="244ea-107">[DefaultSettings](defaultsettings.md)（内容外接程序）：值可以在 32 至 1000 之间</span><span class="sxs-lookup"><span data-stu-id="244ea-107">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="244ea-108">[DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md) （邮件外接程序）：值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="244ea-108">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="244ea-109">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）：如果是 **DetectedEntity** 扩展点，值可以在 140 至 450 之间，如果是 **CustomPane** 扩展点，值可以在 32 至 450 之间</span><span class="sxs-lookup"><span data-stu-id="244ea-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>