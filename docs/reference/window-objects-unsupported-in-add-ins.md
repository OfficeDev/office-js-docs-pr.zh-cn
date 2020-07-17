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
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Office 外接程序中不支持的 Window 对象

对于某些版本的 Windows 和 Office，外接程序在 Internet Explorer 11 运行时中运行。 （有关详细信息，请参阅[Office 外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）`window`在 Internet Explorer 11 中不支持全局对象的某些属性或子属性。 这些属性在外接程序中被禁用，以确保您的外接程序可以为所有用户提供一致的体验，而不管外接程序使用的是哪种浏览器。 这还有助于 AngularJS 正确加载。

下面列出了已禁用的属性。 列表是正在进行的工作。 如果发现外接 `window` 程序中不起作用的其他属性，请使用下面的反馈工具告诉我们。

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>另请参阅

- [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)