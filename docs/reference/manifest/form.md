---
title: 清单文件中的 Form 元素
description: 在特定设备（台式机、平板电脑或电话）上运行时邮件外接程序将使用的窗体的 UX 设置。
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 9b1696b2fecf6b07ee2a3c0a31611d4f2ad1f291
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718207"
---
# <a name="form-element"></a><span data-ttu-id="1dc81-103">Form 元素</span><span class="sxs-lookup"><span data-stu-id="1dc81-103">Form element</span></span>

<span data-ttu-id="1dc81-104">在特定设备（台式机、平板电脑或电话）上运行时邮件外接程序将使用的窗体的 UX 设置。</span><span class="sxs-lookup"><span data-stu-id="1dc81-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1dc81-105">`DesktopSettings`、 `TabletSettings`和`PhoneSettings`元素仅适用于 web 上的经典 Outlook （通常连接到本地 Exchange server 的旧版本）和 Windows 上的 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="1dc81-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="1dc81-106">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="1dc81-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1dc81-107">语法</span><span class="sxs-lookup"><span data-stu-id="1dc81-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="1dc81-108">包含于</span><span class="sxs-lookup"><span data-stu-id="1dc81-108">Contained in</span></span>

[<span data-ttu-id="1dc81-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="1dc81-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="1dc81-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="1dc81-110">Can contain</span></span>

|<span data-ttu-id="1dc81-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="1dc81-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="1dc81-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="1dc81-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="1dc81-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="1dc81-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="1dc81-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="1dc81-114">PhoneSettings</span></span>](phonesettings.md)|
