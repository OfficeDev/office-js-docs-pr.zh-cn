---
title: 清单文件中的 TabletSettings 元素
description: TabletSettings 元素指定在平板电脑上使用邮件外接程序时应用的控制设置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608696"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="d6403-103">TabletSettings 元素</span><span class="sxs-lookup"><span data-stu-id="d6403-103">TabletSettings element</span></span>

<span data-ttu-id="d6403-104">指定在平板电脑上使用邮件外接程序时应用的控制设置。</span><span class="sxs-lookup"><span data-stu-id="d6403-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d6403-105">`TabletSettings`元素仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="d6403-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="d6403-106">若要支持 Android 和 iOS 上的 Outlook，请参阅[适用于 Outlook Mobile 的外接程序](../../outlook/outlook-mobile-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="d6403-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="d6403-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="d6403-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d6403-108">语法</span><span class="sxs-lookup"><span data-stu-id="d6403-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="d6403-109">包含于</span><span class="sxs-lookup"><span data-stu-id="d6403-109">Contained in</span></span>

[<span data-ttu-id="d6403-110">Form</span><span class="sxs-lookup"><span data-stu-id="d6403-110">Form</span></span>](form.md)
