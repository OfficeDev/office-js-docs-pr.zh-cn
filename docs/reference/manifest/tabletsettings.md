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
# <a name="tabletsettings-element"></a><span data-ttu-id="445e7-102">TabletSettings 元素</span><span class="sxs-lookup"><span data-stu-id="445e7-102">TabletSettings element</span></span>

<span data-ttu-id="445e7-103">指定在平板电脑上使用邮件外接程序时应用的控制设置。</span><span class="sxs-lookup"><span data-stu-id="445e7-103">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="445e7-104">元素`TabletSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="445e7-104">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="445e7-105">若要支持 Android 和 iOS 上的 Outlook，请参阅[适用于 Outlook Mobile 的外接程序](/outlook/add-ins/outlook-mobile-addins)。</span><span class="sxs-lookup"><span data-stu-id="445e7-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).</span></span>

<span data-ttu-id="445e7-106">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="445e7-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="445e7-107">语法</span><span class="sxs-lookup"><span data-stu-id="445e7-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="445e7-108">包含于</span><span class="sxs-lookup"><span data-stu-id="445e7-108">Contained in</span></span>

[<span data-ttu-id="445e7-109">Form</span><span class="sxs-lookup"><span data-stu-id="445e7-109">Form</span></span>](form.md)

