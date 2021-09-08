---
title: 清单文件中的 Form 元素
description: 在特定设备（台式机、平板电脑或电话）上运行时邮件外接程序将使用的窗体的 UX 设置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: c9cd1d9104fc51edc84149ef677c4308dfb1a9f5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937136"
---
# <a name="form-element"></a>Form 元素

在特定设备（台式机、平板电脑或电话）上运行时邮件外接程序将使用的窗体的 UX 设置。

> [!IMPORTANT]
> 、 和 元素仅在经典 Outlook 网页版 (中可用，通常连接到旧版本地 Exchange 服务器) 和 Outlook `DesktopSettings` `TabletSettings` `PhoneSettings` 2013 Windows。

**外接程序类型：** 邮件

## <a name="syntax"></a>语法

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
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
