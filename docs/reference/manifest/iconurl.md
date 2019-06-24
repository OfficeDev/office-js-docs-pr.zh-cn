---
title: 清单文件中的 IconUrl 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d4451409a457fa5522e27ab5efd203b9c37a2052
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127553"
---
# <a name="iconurl-element"></a><span data-ttu-id="c0659-102">IconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="c0659-102">IconUrl element</span></span>

<span data-ttu-id="c0659-103">指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="c0659-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="c0659-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="c0659-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c0659-105">语法</span><span class="sxs-lookup"><span data-stu-id="c0659-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="c0659-106">可以包含</span><span class="sxs-lookup"><span data-stu-id="c0659-106">Can contain</span></span>

[<span data-ttu-id="c0659-107">Override</span><span class="sxs-lookup"><span data-stu-id="c0659-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="c0659-108">属性</span><span class="sxs-lookup"><span data-stu-id="c0659-108">Attributes</span></span>

|<span data-ttu-id="c0659-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="c0659-109">**Attribute**</span></span>|<span data-ttu-id="c0659-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="c0659-110">**Type**</span></span>|<span data-ttu-id="c0659-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="c0659-111">**Required**</span></span>|<span data-ttu-id="c0659-112">**描述**</span><span class="sxs-lookup"><span data-stu-id="c0659-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c0659-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="c0659-113">DefaultValue</span></span>|<span data-ttu-id="c0659-114">字符串</span><span class="sxs-lookup"><span data-stu-id="c0659-114">string</span></span>|<span data-ttu-id="c0659-115">必需</span><span class="sxs-lookup"><span data-stu-id="c0659-115">required</span></span>|<span data-ttu-id="c0659-116">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="c0659-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c0659-117">注解</span><span class="sxs-lookup"><span data-stu-id="c0659-117">Remarks</span></span>

<span data-ttu-id="c0659-118">对于邮件外接程序, 该图标显示在 "**文件** > **管理外接程序**" ui (outlook) 中, 或**设置** > "**管理外接程序**" ui (outlook 网页版)。</span><span class="sxs-lookup"><span data-stu-id="c0659-118">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="c0659-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span><span class="sxs-lookup"><span data-stu-id="c0659-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="c0659-120">For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span><span class="sxs-lookup"><span data-stu-id="c0659-120">For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="c0659-121">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="c0659-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="c0659-122">对于内容和任务窗格应用程序，指定的图像必须是 32 x 32 像素。</span><span class="sxs-lookup"><span data-stu-id="c0659-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="c0659-123">对于邮件应用程序，推荐的图像分辨率是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="c0659-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="c0659-124">此外，还应指定用于使用 [HighResolutionIconUrl](highresolutioniconurl.md) 元素在高 DPI 屏幕上运行的 Office 主机应用程序的图标。</span><span class="sxs-lookup"><span data-stu-id="c0659-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="c0659-125">有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。</span><span class="sxs-lookup"><span data-stu-id="c0659-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="c0659-126">当前不支持在运行`IconUrl`时更改元素的值。</span><span class="sxs-lookup"><span data-stu-id="c0659-126">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>