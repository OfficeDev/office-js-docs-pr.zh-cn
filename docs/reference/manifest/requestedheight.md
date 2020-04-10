---
title: 清单文件中的 RequestedHeight 元素
description: RequestedHeight 元素指定内容或邮件加载项的初始高度（以像素为单位）。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f4c3ca1ff39cc3150249fbc824b0db76f6b8a85
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215038"
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
- [ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）：如果是 **DetectedEntity** 扩展点，值可以在 140 至 450 之间，如果是 **CustomPane** 扩展点，值可以在 32 至 450 之间
