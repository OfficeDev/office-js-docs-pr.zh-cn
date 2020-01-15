---
title: 清单文件中的 Form 元素
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: d545d471e007f0077a8310b0b847bbbf99a8f7ac
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120647"
---
# <a name="form-element"></a><span data-ttu-id="4373e-102">Form 元素</span><span class="sxs-lookup"><span data-stu-id="4373e-102">Form element</span></span>

<span data-ttu-id="4373e-103">在特定设备（台式机、平板电脑或电话）上运行时邮件外接程序将使用的窗体的 UX 设置。</span><span class="sxs-lookup"><span data-stu-id="4373e-103">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4373e-104">`DesktopSettings`、 `TabletSettings`和`PhoneSettings`元素仅适用于 web 上的经典 Outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="4373e-104">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="4373e-105">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="4373e-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4373e-106">语法</span><span class="sxs-lookup"><span data-stu-id="4373e-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="4373e-107">包含于</span><span class="sxs-lookup"><span data-stu-id="4373e-107">Contained in</span></span>

[<span data-ttu-id="4373e-108">FormSettings</span><span class="sxs-lookup"><span data-stu-id="4373e-108">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="4373e-109">可以包含</span><span class="sxs-lookup"><span data-stu-id="4373e-109">Can contain</span></span>

|<span data-ttu-id="4373e-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="4373e-110">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="4373e-111">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="4373e-111">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="4373e-112">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="4373e-112">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="4373e-113">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="4373e-113">PhoneSettings</span></span>](phonesettings.md)|
