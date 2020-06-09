---
title: 清单文件中的 DesktopSettings 元素
description: 指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 50201080d8be3c8943d16730c34a4bac236d7b90
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612274"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="7a7fc-103">DesktopSettings 元素</span><span class="sxs-lookup"><span data-stu-id="7a7fc-103">DesktopSettings element</span></span>

<span data-ttu-id="7a7fc-104">指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。</span><span class="sxs-lookup"><span data-stu-id="7a7fc-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7a7fc-105">`DesktopSettings`元素仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="7a7fc-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="7a7fc-106">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="7a7fc-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7a7fc-107">语法</span><span class="sxs-lookup"><span data-stu-id="7a7fc-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="7a7fc-108">包含于</span><span class="sxs-lookup"><span data-stu-id="7a7fc-108">Contained in</span></span>

[<span data-ttu-id="7a7fc-109">Form</span><span class="sxs-lookup"><span data-stu-id="7a7fc-109">Form</span></span>](form.md)
