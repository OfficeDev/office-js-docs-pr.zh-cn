---
title: Office 外接程序中不支持的 Window 对象
description: 本文指定了一些在 Office 外接程序中无法工作的窗口运行时对象。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160501"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a><span data-ttu-id="eb5ea-103">Office 外接程序中不支持的 Window 对象</span><span class="sxs-lookup"><span data-stu-id="eb5ea-103">Window objects that are unsupported in Office Add-ins</span></span>

<span data-ttu-id="eb5ea-104">对于某些版本的 Windows 和 Office，外接程序在 Internet Explorer 11 运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-104">For some versions of Windows and Office, add-ins run in an Internet Explorer 11 runtime.</span></span> <span data-ttu-id="eb5ea-105">（有关详细信息，请参阅[Office 外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）`window`在 Internet Explorer 11 中不支持全局对象的某些属性或子属性。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-105">(For details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Some properties or subproperties of the global `window` object are not supported in Internet Explorer 11.</span></span> <span data-ttu-id="eb5ea-106">这些属性在外接程序中被禁用，以确保您的外接程序可以为所有用户提供一致的体验，而不管外接程序使用的是哪种浏览器。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-106">These properties are disabled in add-ins to ensure that your add-in provides a consistent experience to all users, regardless of which browser the add-in is using.</span></span> <span data-ttu-id="eb5ea-107">这还有助于 AngularJS 正确加载。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-107">This also helps AngularJS load properly.</span></span>

<span data-ttu-id="eb5ea-108">下面列出了已禁用的属性。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-108">The following is a list of the disabled properties.</span></span> <span data-ttu-id="eb5ea-109">列表是正在进行的工作。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-109">The list is a work in progress.</span></span> <span data-ttu-id="eb5ea-110">如果发现外接 `window` 程序中不起作用的其他属性，请使用下面的反馈工具告诉我们。</span><span class="sxs-lookup"><span data-stu-id="eb5ea-110">If you discover additional `window` properties that do not work in add-ins, please use the feedback tool below to tell us.</span></span>

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a><span data-ttu-id="eb5ea-111">另请参阅</span><span class="sxs-lookup"><span data-stu-id="eb5ea-111">See also</span></span>

- [<span data-ttu-id="eb5ea-112">Office 加载项使用的浏览器</span><span class="sxs-lookup"><span data-stu-id="eb5ea-112">Browsers used by Office Add-ins</span></span>](../concepts/browsers-used-by-office-web-add-ins.md)