---
title: 在加载项中不受支持Office窗口对象
description: 本文指定某些在加载项中不起作用的窗口运行时Office对象。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 654e8e311069a616e2d8859a4f63b19d299609982fa68449b5529df489816cbf
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097379"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>在加载项中不受支持Office窗口对象

对于某些版本的 Windows 和 Office，外接程序在 Internet Explorer 11 运行时中运行。  (有关详细信息，请参阅[Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) 某些属性或全局对象的子属性在 Internet Explorer `window` 11 中不受支持。 外接程序中禁用这些属性，以确保外接程序为所有用户提供一致的体验，无论外接程序使用的是哪个浏览器。 这还有助于正确加载 AngularJS。

以下是已禁用属性的列表。 该列表是一项进行中的工作。 如果发现其他属性在加载项中不起作用，请使用下面的 `window` 反馈工具告诉我们。

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>另请参阅

- [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)