---
title: 清单文件中的 DesktopSettings 元素
description: 指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d48532482fc71fec2a96133ee8e813cae798613f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718354"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="7868b-103">DesktopSettings 元素</span><span class="sxs-lookup"><span data-stu-id="7868b-103">DesktopSettings element</span></span>

<span data-ttu-id="7868b-104">指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。</span><span class="sxs-lookup"><span data-stu-id="7868b-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7868b-105">元素`DesktopSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="7868b-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="7868b-106">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="7868b-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7868b-107">语法</span><span class="sxs-lookup"><span data-stu-id="7868b-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="7868b-108">包含于</span><span class="sxs-lookup"><span data-stu-id="7868b-108">Contained in</span></span>

[<span data-ttu-id="7868b-109">Form</span><span class="sxs-lookup"><span data-stu-id="7868b-109">Form</span></span>](form.md)
