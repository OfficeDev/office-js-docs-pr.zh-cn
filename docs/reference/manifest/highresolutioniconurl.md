---
title: 清单文件中的 HighResolutionIconUrl 元素
description: 指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 78a9296f38a688073e516fb78a77bb4cdac822c4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718137"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="4c85d-103">HighResolutionIconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="4c85d-103">HighResolutionIconUrl element</span></span>

<span data-ttu-id="4c85d-104">指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="4c85d-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="4c85d-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="4c85d-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4c85d-106">语法</span><span class="sxs-lookup"><span data-stu-id="4c85d-106">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="4c85d-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="4c85d-107">Can contain</span></span>

[<span data-ttu-id="4c85d-108">Override</span><span class="sxs-lookup"><span data-stu-id="4c85d-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="4c85d-109">属性</span><span class="sxs-lookup"><span data-stu-id="4c85d-109">Attributes</span></span>

|<span data-ttu-id="4c85d-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="4c85d-110">**Attribute**</span></span>|<span data-ttu-id="4c85d-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="4c85d-111">**Type**</span></span>|<span data-ttu-id="4c85d-112">**必需**</span><span class="sxs-lookup"><span data-stu-id="4c85d-112">**Required**</span></span>|<span data-ttu-id="4c85d-113">**描述**</span><span class="sxs-lookup"><span data-stu-id="4c85d-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4c85d-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="4c85d-114">DefaultValue</span></span>|<span data-ttu-id="4c85d-115">字符串 (URL)</span><span class="sxs-lookup"><span data-stu-id="4c85d-115">string (URL)</span></span>|<span data-ttu-id="4c85d-116">必需</span><span class="sxs-lookup"><span data-stu-id="4c85d-116">required</span></span>|<span data-ttu-id="4c85d-117">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="4c85d-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="4c85d-118">注解</span><span class="sxs-lookup"><span data-stu-id="4c85d-118">Remarks</span></span>

<span data-ttu-id="4c85d-119">对于邮件外接程序，该图标将显示在 "**文件** > **管理外接程序**" UI 中。</span><span class="sxs-lookup"><span data-stu-id="4c85d-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="4c85d-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span><span class="sxs-lookup"><span data-stu-id="4c85d-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="4c85d-121">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="4c85d-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="4c85d-122">对于内容和任务窗格应用程序，推荐的图像分辨率是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="4c85d-122">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="4c85d-123">对于邮件应用程序，图像必须是 128 x 128 像素。</span><span class="sxs-lookup"><span data-stu-id="4c85d-123">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="4c85d-124">有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。</span><span class="sxs-lookup"><span data-stu-id="4c85d-124">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
