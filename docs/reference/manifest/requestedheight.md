---
title: 清单文件中的 RequestedHeight 元素
description: RequestedHeight 元素指定内容 (或邮件) 的初始高度值（以像素为单位）。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 0e5f9de909d32622ac244ff4118c8a3192abf2ff0fe89ed81a6188ddcb265549
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092981"
---
# <a name="requestedheight-element"></a>RequestedHeight 元素

指定内容外接程序或邮件外接程序的初始高度（以像素为单位）。

**外接程序类型：** 内容、邮件

## <a name="syntax"></a>语法

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>包含于

- [DefaultSettings](defaultsettings.md)（内容外接程序）：值可以在 32 至 1000 之间
- [DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md) （邮件外接程序）：值可以在 32 至 450 之间
- [ExtensionPoint](extensionpoint.md) (上下文邮件外接程序) ，对于 **DetectedEntity** 扩展点，该值可以介于 140 和 450 之间，对于已弃用的 CustomPane 扩展点，该值介于 32 和 450 之间 ([  450)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)
