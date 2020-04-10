---
title: 清单文件中的 DesktopSettings 元素
description: 指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 574e04ec577f831e17184cf4f801dae22441bca2
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215073"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="694d5-103">DesktopSettings 元素</span><span class="sxs-lookup"><span data-stu-id="694d5-103">DesktopSettings element</span></span>

<span data-ttu-id="694d5-104">指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。</span><span class="sxs-lookup"><span data-stu-id="694d5-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="694d5-105">元素`DesktopSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="694d5-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="694d5-106">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="694d5-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="694d5-107">语法</span><span class="sxs-lookup"><span data-stu-id="694d5-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="694d5-108">包含于</span><span class="sxs-lookup"><span data-stu-id="694d5-108">Contained in</span></span>

[<span data-ttu-id="694d5-109">Form</span><span class="sxs-lookup"><span data-stu-id="694d5-109">Form</span></span>](form.md)
