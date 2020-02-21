---
title: 清单文件中的 PhoneSettings 元素
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 4614c86af865e5242657f47e21e6786545a616b6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165536"
---
# <a name="phonesettings-element"></a><span data-ttu-id="ff6d2-102">PhoneSettings 元素</span><span class="sxs-lookup"><span data-stu-id="ff6d2-102">PhoneSettings element</span></span>

<span data-ttu-id="ff6d2-103">指定在手机上使用邮件外接程序时应用的源位置和控制设置。</span><span class="sxs-lookup"><span data-stu-id="ff6d2-103">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ff6d2-104">元素`PhoneSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="ff6d2-104">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="ff6d2-105">若要支持 Android 和 iOS 上的 Outlook，请参阅[适用于 Outlook Mobile 的外接程序](../../outlook/outlook-mobile-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="ff6d2-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="ff6d2-106">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="ff6d2-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ff6d2-107">语法</span><span class="sxs-lookup"><span data-stu-id="ff6d2-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="ff6d2-108">包含于</span><span class="sxs-lookup"><span data-stu-id="ff6d2-108">Contained in</span></span>

[<span data-ttu-id="ff6d2-109">Form</span><span class="sxs-lookup"><span data-stu-id="ff6d2-109">Form</span></span>](form.md)

