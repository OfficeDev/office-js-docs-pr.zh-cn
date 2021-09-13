---
title: 清单文件中的 PhoneSettings 元素
description: PhoneSettings 元素指定在手机上使用邮件外接程序时所应用到的源位置和控制设置。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 1e52827a20ee95397541f7c1d54c732ff8f96ba5
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152541"
---
# <a name="phonesettings-element"></a>PhoneSettings 元素

指定在手机上使用邮件外接程序时应用的源位置和控制设置。

> [!IMPORTANT]
> 此元素仅适用于经典Outlook 网页版 (通常连接到旧版本地 Exchange 服务器) Outlook `PhoneSettings` 2013 Windows。 若要支持Outlook Android 和 iOS 上的加载项，请参阅适用于[Outlook Mobile 的外接程序](../../outlook/outlook-mobile-addins.md)。

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

