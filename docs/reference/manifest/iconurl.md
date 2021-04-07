---
title: 清单文件中的 IconUrl 元素
description: IconUrl 元素指定表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 68a449b40f6084d26140d59fec61967e163196df
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604636"
---
# <a name="iconurl-element"></a><span data-ttu-id="cb54a-103">IconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="cb54a-103">IconUrl element</span></span>

<span data-ttu-id="cb54a-104">指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="cb54a-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="cb54a-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="cb54a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cb54a-106">语法</span><span class="sxs-lookup"><span data-stu-id="cb54a-106">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="cb54a-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="cb54a-107">Can contain</span></span>

[<span data-ttu-id="cb54a-108">Override</span><span class="sxs-lookup"><span data-stu-id="cb54a-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="cb54a-109">属性</span><span class="sxs-lookup"><span data-stu-id="cb54a-109">Attributes</span></span>

|<span data-ttu-id="cb54a-110">属性</span><span class="sxs-lookup"><span data-stu-id="cb54a-110">Attribute</span></span>|<span data-ttu-id="cb54a-111">类型</span><span class="sxs-lookup"><span data-stu-id="cb54a-111">Type</span></span>|<span data-ttu-id="cb54a-112">必需</span><span class="sxs-lookup"><span data-stu-id="cb54a-112">Required</span></span>|<span data-ttu-id="cb54a-113">Description</span><span class="sxs-lookup"><span data-stu-id="cb54a-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cb54a-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="cb54a-114">DefaultValue</span></span>|<span data-ttu-id="cb54a-115">字符串</span><span class="sxs-lookup"><span data-stu-id="cb54a-115">string</span></span>|<span data-ttu-id="cb54a-116">必需</span><span class="sxs-lookup"><span data-stu-id="cb54a-116">required</span></span>|<span data-ttu-id="cb54a-117">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="cb54a-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="cb54a-118">注解</span><span class="sxs-lookup"><span data-stu-id="cb54a-118">Remarks</span></span>

<span data-ttu-id="cb54a-119">对于邮件外接程序，图标显示在"文件管理外接程序"UI  (Outlook) 或"设置""管理外接程序  >    >  "UI (Outlook 网页) 。</span><span class="sxs-lookup"><span data-stu-id="cb54a-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="cb54a-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span><span class="sxs-lookup"><span data-stu-id="cb54a-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="cb54a-121">对于所有加载项类型，如果你将加载项发布到 [AppSource，](https://appsource.microsoft.com)图标也会在 AppSource 中使用。</span><span class="sxs-lookup"><span data-stu-id="cb54a-121">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="cb54a-122">图像必须采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="cb54a-122">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="cb54a-123">对于内容和任务窗格应用程序，指定的图像必须是 32 x 32 像素。</span><span class="sxs-lookup"><span data-stu-id="cb54a-123">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="cb54a-124">对于邮件应用程序，图像分辨率必须为 64 x 64 像素。</span><span class="sxs-lookup"><span data-stu-id="cb54a-124">For mail apps, the image resolution must be 64 x 64 pixels.</span></span> <span data-ttu-id="cb54a-125">还应指定一个图标，以用于使用 [HighResolutionIconUrl](highresolutioniconurl.md) 元素在高 DPI 屏幕上运行的 Office 客户端应用程序。</span><span class="sxs-lookup"><span data-stu-id="cb54a-125">You should also specify an icon for use with Office client applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="cb54a-126">有关详细信息，请参阅 [在 AppSource 和 Office 中创建有效的应用一览](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)中的 _为你的应用创建一致的视觉标识_ 部分。</span><span class="sxs-lookup"><span data-stu-id="cb54a-126">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="cb54a-127">当前不支持 `IconUrl` 在运行时更改 元素的值。</span><span class="sxs-lookup"><span data-stu-id="cb54a-127">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>
