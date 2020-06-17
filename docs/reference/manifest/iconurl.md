---
title: 清单文件中的 IconUrl 元素
description: IconUrl 元素指定代表用户插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2ccfc2fc1d0a07f6d549f388bbb58e40e79a17d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611790"
---
# <a name="iconurl-element"></a><span data-ttu-id="3467b-103">IconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="3467b-103">IconUrl element</span></span>

<span data-ttu-id="3467b-104">指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="3467b-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="3467b-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="3467b-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3467b-106">语法</span><span class="sxs-lookup"><span data-stu-id="3467b-106">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="3467b-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="3467b-107">Can contain</span></span>

[<span data-ttu-id="3467b-108">Override</span><span class="sxs-lookup"><span data-stu-id="3467b-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="3467b-109">属性</span><span class="sxs-lookup"><span data-stu-id="3467b-109">Attributes</span></span>

|<span data-ttu-id="3467b-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="3467b-110">**Attribute**</span></span>|<span data-ttu-id="3467b-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="3467b-111">**Type**</span></span>|<span data-ttu-id="3467b-112">**必需**</span><span class="sxs-lookup"><span data-stu-id="3467b-112">**Required**</span></span>|<span data-ttu-id="3467b-113">**描述**</span><span class="sxs-lookup"><span data-stu-id="3467b-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3467b-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="3467b-114">DefaultValue</span></span>|<span data-ttu-id="3467b-115">字符串</span><span class="sxs-lookup"><span data-stu-id="3467b-115">string</span></span>|<span data-ttu-id="3467b-116">必需</span><span class="sxs-lookup"><span data-stu-id="3467b-116">required</span></span>|<span data-ttu-id="3467b-117">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="3467b-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="3467b-118">注解</span><span class="sxs-lookup"><span data-stu-id="3467b-118">Remarks</span></span>

<span data-ttu-id="3467b-119">对于邮件外接程序，该图标显示在 "**文件**  >  **管理外接程序**" ui （outlook）中，或**设置**"  >  **管理外接程序**" ui （outlook 网页版）。</span><span class="sxs-lookup"><span data-stu-id="3467b-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="3467b-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span><span class="sxs-lookup"><span data-stu-id="3467b-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="3467b-121">对于所有外接程序类型，如果您将外接程序发布到 AppSource，则还会在[AppSource](https://appsource.microsoft.com)中使用该图标。</span><span class="sxs-lookup"><span data-stu-id="3467b-121">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="3467b-122">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="3467b-122">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="3467b-123">对于内容和任务窗格应用程序，指定的图像必须是 32 x 32 像素。</span><span class="sxs-lookup"><span data-stu-id="3467b-123">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="3467b-124">对于邮件应用程序，推荐的图像分辨率是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="3467b-124">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="3467b-125">此外，还应指定用于使用 [HighResolutionIconUrl](highresolutioniconurl.md) 元素在高 DPI 屏幕上运行的 Office 主机应用程序的图标。</span><span class="sxs-lookup"><span data-stu-id="3467b-125">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="3467b-126">有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。</span><span class="sxs-lookup"><span data-stu-id="3467b-126">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="3467b-127">`IconUrl`当前不支持在运行时更改元素的值。</span><span class="sxs-lookup"><span data-stu-id="3467b-127">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>