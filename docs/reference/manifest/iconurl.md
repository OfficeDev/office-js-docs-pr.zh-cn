---
title: 清单文件中的 IconUrl 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 858f399ed36bfed60c3e091b26ac7400ff901179
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325260"
---
# <a name="iconurl-element"></a><span data-ttu-id="90c56-102">IconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="90c56-102">IconUrl element</span></span>

<span data-ttu-id="90c56-103">指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="90c56-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="90c56-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="90c56-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="90c56-105">语法</span><span class="sxs-lookup"><span data-stu-id="90c56-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="90c56-106">可以包含</span><span class="sxs-lookup"><span data-stu-id="90c56-106">Can contain</span></span>

[<span data-ttu-id="90c56-107">Override</span><span class="sxs-lookup"><span data-stu-id="90c56-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="90c56-108">属性</span><span class="sxs-lookup"><span data-stu-id="90c56-108">Attributes</span></span>

|<span data-ttu-id="90c56-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="90c56-109">**Attribute**</span></span>|<span data-ttu-id="90c56-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="90c56-110">**Type**</span></span>|<span data-ttu-id="90c56-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="90c56-111">**Required**</span></span>|<span data-ttu-id="90c56-112">**描述**</span><span class="sxs-lookup"><span data-stu-id="90c56-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="90c56-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="90c56-113">DefaultValue</span></span>|<span data-ttu-id="90c56-114">字符串</span><span class="sxs-lookup"><span data-stu-id="90c56-114">string</span></span>|<span data-ttu-id="90c56-115">必需</span><span class="sxs-lookup"><span data-stu-id="90c56-115">required</span></span>|<span data-ttu-id="90c56-116">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="90c56-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="90c56-117">注解</span><span class="sxs-lookup"><span data-stu-id="90c56-117">Remarks</span></span>

<span data-ttu-id="90c56-118">对于邮件外接程序，该图标显示在 "**文件** > **管理外接程序**" ui （outlook）中，或**设置** > "**管理外接程序**" ui （outlook 网页版）。</span><span class="sxs-lookup"><span data-stu-id="90c56-118">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="90c56-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span><span class="sxs-lookup"><span data-stu-id="90c56-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="90c56-120">对于所有外接程序类型，如果您将外接程序发布到 AppSource，则还会在[AppSource](https://appsource.microsoft.com)中使用该图标。</span><span class="sxs-lookup"><span data-stu-id="90c56-120">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="90c56-121">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="90c56-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="90c56-122">对于内容和任务窗格应用程序，指定的图像必须是 32 x 32 像素。</span><span class="sxs-lookup"><span data-stu-id="90c56-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="90c56-123">对于邮件应用程序，推荐的图像分辨率是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="90c56-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="90c56-124">此外，还应指定用于使用 [HighResolutionIconUrl](highresolutioniconurl.md) 元素在高 DPI 屏幕上运行的 Office 主机应用程序的图标。</span><span class="sxs-lookup"><span data-stu-id="90c56-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="90c56-125">有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。</span><span class="sxs-lookup"><span data-stu-id="90c56-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="90c56-126">当前不支持在运行`IconUrl`时更改元素的值。</span><span class="sxs-lookup"><span data-stu-id="90c56-126">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>