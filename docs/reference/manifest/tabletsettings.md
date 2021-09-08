---
title: 清单文件中的 TabletSettings 元素
description: TabletSettings 元素指定在平板电脑上使用邮件外接程序时所应用的控制设置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936216"
---
# <a name="tabletsettings-element"></a>TabletSettings 元素

指定在平板电脑上使用邮件外接程序时应用的控制设置。

> [!IMPORTANT]
> 此元素仅在经典版本中可用Outlook 网页版 (通常连接到旧版本地 `TabletSettings` Exchange server) ，Outlook 2013 on Windows。 若要支持Outlook Android 和 iOS 上的加载项，请参阅适用于[Outlook Mobile 的外接程序](../../outlook/outlook-mobile-addins.md)。

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

[Form](form.md)
