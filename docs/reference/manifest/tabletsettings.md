---
title: 清单文件中的 TabletSettings 元素
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 977fc2a781f3b93e4eb36041473c683196314adb
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120619"
---
# <a name="tabletsettings-element"></a>TabletSettings 元素

指定在平板电脑上使用邮件外接程序时应用的控制设置。

> [!IMPORTANT]
> 元素`TabletSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。 若要支持 Android 和 iOS 上的 Outlook，请参阅[适用于 Outlook Mobile 的外接程序](/outlook/add-ins/outlook-mobile-addins)。

**外接程序类型：** 邮件

## <a name="syntax"></a>语法

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>包含于

[Form](form.md)

