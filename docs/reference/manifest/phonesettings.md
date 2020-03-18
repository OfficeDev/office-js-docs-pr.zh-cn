---
title: 清单文件中的 PhoneSettings 元素
description: PhoneSettings 元素指定在手机上使用邮件外接程序时应用的源位置和控制设置。
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 581a3ae71a58cd05aac52129a6f4395a60c20cef
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720474"
---
# <a name="phonesettings-element"></a><span data-ttu-id="65bdb-103">PhoneSettings 元素</span><span class="sxs-lookup"><span data-stu-id="65bdb-103">PhoneSettings element</span></span>

<span data-ttu-id="65bdb-104">指定在手机上使用邮件外接程序时应用的源位置和控制设置。</span><span class="sxs-lookup"><span data-stu-id="65bdb-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="65bdb-105">元素`PhoneSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="65bdb-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="65bdb-106">若要支持 Android 和 iOS 上的 Outlook，请参阅[适用于 Outlook Mobile 的外接程序](../../outlook/outlook-mobile-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="65bdb-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="65bdb-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="65bdb-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="65bdb-108">语法</span><span class="sxs-lookup"><span data-stu-id="65bdb-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="65bdb-109">包含于</span><span class="sxs-lookup"><span data-stu-id="65bdb-109">Contained in</span></span>

[<span data-ttu-id="65bdb-110">Form</span><span class="sxs-lookup"><span data-stu-id="65bdb-110">Form</span></span>](form.md)

