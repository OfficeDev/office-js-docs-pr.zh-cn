---
title: 清单文件中的 DesktopSettings 元素
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6dfa69d407e267a1cbcfdeaad0bdf9cdf75c1465
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120640"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="b795e-102">DesktopSettings 元素</span><span class="sxs-lookup"><span data-stu-id="b795e-102">DesktopSettings element</span></span>

<span data-ttu-id="b795e-103">指定在台式计算机上使用邮件外接程序时应用的源位置和控制设置。</span><span class="sxs-lookup"><span data-stu-id="b795e-103">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b795e-104">元素`DesktopSettings`仅适用于 web 上的经典 outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="b795e-104">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="b795e-105">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="b795e-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b795e-106">语法</span><span class="sxs-lookup"><span data-stu-id="b795e-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b795e-107">包含于</span><span class="sxs-lookup"><span data-stu-id="b795e-107">Contained in</span></span>

[<span data-ttu-id="b795e-108">Form</span><span class="sxs-lookup"><span data-stu-id="b795e-108">Form</span></span>](form.md)
