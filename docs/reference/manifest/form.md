---
title: 清单文件中的 Form 元素
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: d545d471e007f0077a8310b0b847bbbf99a8f7ac
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120647"
---
# <a name="form-element"></a>Form 元素

在特定设备（台式机、平板电脑或电话）上运行时邮件外接程序将使用的窗体的 UX 设置。

> [!IMPORTANT]
> `DesktopSettings`、 `TabletSettings`和`PhoneSettings`元素仅适用于 web 上的经典 Outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。

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

[FormSettings](formsettings.md)


## <a name="can-contain"></a>可以包含

|**Element**|
|:-----|
|[DesktopSettings](desktopsettings.md)|
|[TabletSettings](tabletsettings.md)|
|[PhoneSettings](phonesettings.md)|
