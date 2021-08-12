---
title: 清单文件中的 PhoneSettings 元素
description: PhoneSettings 元素指定在手机上使用邮件外接程序时所应用到的源位置和控制设置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 883242dc290384f9f0b72736338bdd78a2d23ffeee6cf3aee46d5acd970654ab
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085724"
---
# <a name="phonesettings-element"></a>PhoneSettings 元素

指定在手机上使用邮件外接程序时应用的源位置和控制设置。

> [!IMPORTANT]
> 元素仅在经典版本中可用Outlook 网页版 (通常连接到旧版本地 Exchange server) 和 Outlook `PhoneSettings` 2013 Windows。 若要支持Outlook Android 和 iOS 上的加载项，请参阅适用于[Outlook Mobile 的外接程序](../../outlook/outlook-mobile-addins.md)。

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

