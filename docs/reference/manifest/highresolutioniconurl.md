---
title: 清单文件中的 HighResolutionIconUrl 元素
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 41008be6b60d260bef78808af2b8dee1fbd0864a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325267"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="0c79a-102">HighResolutionIconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="0c79a-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="0c79a-103">指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="0c79a-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="0c79a-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="0c79a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0c79a-105">语法</span><span class="sxs-lookup"><span data-stu-id="0c79a-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="0c79a-106">可以包含</span><span class="sxs-lookup"><span data-stu-id="0c79a-106">Can contain</span></span>

[<span data-ttu-id="0c79a-107">Override</span><span class="sxs-lookup"><span data-stu-id="0c79a-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="0c79a-108">属性</span><span class="sxs-lookup"><span data-stu-id="0c79a-108">Attributes</span></span>

|<span data-ttu-id="0c79a-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="0c79a-109">**Attribute**</span></span>|<span data-ttu-id="0c79a-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="0c79a-110">**Type**</span></span>|<span data-ttu-id="0c79a-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="0c79a-111">**Required**</span></span>|<span data-ttu-id="0c79a-112">**描述**</span><span class="sxs-lookup"><span data-stu-id="0c79a-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0c79a-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="0c79a-113">DefaultValue</span></span>|<span data-ttu-id="0c79a-114">字符串 (URL)</span><span class="sxs-lookup"><span data-stu-id="0c79a-114">string (URL)</span></span>|<span data-ttu-id="0c79a-115">必需</span><span class="sxs-lookup"><span data-stu-id="0c79a-115">required</span></span>|<span data-ttu-id="0c79a-116">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="0c79a-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="0c79a-117">注解</span><span class="sxs-lookup"><span data-stu-id="0c79a-117">Remarks</span></span>

<span data-ttu-id="0c79a-118">对于邮件外接程序，该图标将显示在 "**文件** > **管理外接程序**" UI 中。</span><span class="sxs-lookup"><span data-stu-id="0c79a-118">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="0c79a-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span><span class="sxs-lookup"><span data-stu-id="0c79a-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="0c79a-120">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="0c79a-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="0c79a-121">对于内容和任务窗格应用程序，推荐的图像分辨率是 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="0c79a-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="0c79a-122">对于邮件应用程序，图像必须是 128 x 128 像素。</span><span class="sxs-lookup"><span data-stu-id="0c79a-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="0c79a-123">有关详细信息，请参阅[在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的_为你的应用创建一致的视觉标识_部分。</span><span class="sxs-lookup"><span data-stu-id="0c79a-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
