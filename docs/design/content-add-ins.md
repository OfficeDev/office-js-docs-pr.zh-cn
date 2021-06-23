---
title: 内容 Office 加载项
description: 内容加载项是指可以直接嵌入 Excel 或 PowerPoint 文档的图面，用户可以通过它访问界面控件，运行代码以修改文档或显示数据源中的数据。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 9f7ccd4cfaed5132debb7017caaf3b9da733850d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076355"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="0abbc-103">内容 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0abbc-103">Content Office Add-ins</span></span>

<span data-ttu-id="0abbc-104">内容加载项是指可以直接嵌入 Excel 或 PowerPoint 文档的图面。</span><span class="sxs-lookup"><span data-stu-id="0abbc-104">Content add-ins are surfaces that can be embedded directly into Excel or PowerPoint documents.</span></span> <span data-ttu-id="0abbc-105">用户可以通过内容加载项访问界面控件，运行代码以修改文档或显示数据源中的数据。</span><span class="sxs-lookup"><span data-stu-id="0abbc-105">Content add-ins give users access to interface controls that run code to modify documents or display data from a data source.</span></span> <span data-ttu-id="0abbc-106">在你要将功能直接嵌入文档时，请使用内容加载项。</span><span class="sxs-lookup"><span data-stu-id="0abbc-106">Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="0abbc-107">*图 1. 内容加载项的典型布局*</span><span class="sxs-lookup"><span data-stu-id="0abbc-107">*Figure 1. Typical layout for content add-ins*</span></span>

![应用程序内容加载项的典型布局Office布局。](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="0abbc-109">最佳做法</span><span class="sxs-lookup"><span data-stu-id="0abbc-109">Best practices</span></span>

- <span data-ttu-id="0abbc-110">在加载项顶部包括某些导航或命令元素，如命令栏或透视。</span><span class="sxs-lookup"><span data-stu-id="0abbc-110">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="0abbc-111">包括位于加载项底部的品牌元素，如品牌栏（仅适用于 Excel 和 PowerPoint 加载项）。</span><span class="sxs-lookup"><span data-stu-id="0abbc-111">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Excel and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="0abbc-112">变量</span><span class="sxs-lookup"><span data-stu-id="0abbc-112">Variants</span></span>

<span data-ttu-id="0abbc-113">用户指定桌面Excel PowerPoint Office外接程序Microsoft 365外接程序大小。</span><span class="sxs-lookup"><span data-stu-id="0abbc-113">Content add-in sizes for Excel and PowerPoint in Office desktop and Microsoft 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="0abbc-114">“个性”菜单</span><span class="sxs-lookup"><span data-stu-id="0abbc-114">Personality menu</span></span>

<span data-ttu-id="0abbc-p102">“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。</span><span class="sxs-lookup"><span data-stu-id="0abbc-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="0abbc-117">对于 Windows，个性菜单尺寸为 12x32 像素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="0abbc-117">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="0abbc-118">*图 2：Windows 上的个性菜单*</span><span class="sxs-lookup"><span data-stu-id="0abbc-118">*Figure 2. Personality menu on Windows*</span></span>

![桌面上的 12x32 像素Windows菜单。](../images/personality-menu-win.png)

<span data-ttu-id="0abbc-120">对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将占用空间增加至 34x32 像素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="0abbc-120">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="0abbc-121">*图 3：Mac 上的个性菜单*</span><span class="sxs-lookup"><span data-stu-id="0abbc-121">*Figure 3. Personality menu on Mac*</span></span>

![Mac 桌面上的 34x32 像素个性菜单。](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="0abbc-123">实现</span><span class="sxs-lookup"><span data-stu-id="0abbc-123">Implementation</span></span>

<span data-ttu-id="0abbc-124">有关实现内容加载项的示例，请参阅 GitHub 上的 [Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。</span><span class="sxs-lookup"><span data-stu-id="0abbc-124">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="0abbc-125">支持注意事项</span><span class="sxs-lookup"><span data-stu-id="0abbc-125">Support considerations</span></span>

- <span data-ttu-id="0abbc-126">检查你的加载项Office应用程序或平台的特定Office[工作](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="0abbc-126">Check to see if your Office Add-in will work on a [specific Office application or platform](../overview/office-add-in-availability.md).</span></span>
- <span data-ttu-id="0abbc-127">一些内容加载项可能会要求用户“信任”加载项对 Excel 或 PowerPoint 执行读取和写入操作。</span><span class="sxs-lookup"><span data-stu-id="0abbc-127">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="0abbc-128">可以在加载项清单中声明要拥有的[权限级别](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="0abbc-128">You can declare what [level of permissions](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) you want your user to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="0abbc-p104">Office 2013 版本及更高版本中的 Excel 和 PowerPoint 支持内容加载项。 如果在不支持 Office Web 加载项的 Office 版本中打开加载项，加载项会显示为图像。</span><span class="sxs-lookup"><span data-stu-id="0abbc-p104">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="0abbc-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0abbc-131">See also</span></span>

- [<span data-ttu-id="0abbc-132">Office 客户端应用程序和 Office 加载项的平台可用性</span><span class="sxs-lookup"><span data-stu-id="0abbc-132">Office client application and platform availability for Office Add-ins</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="0abbc-133">Office 加载项中的 Fabric Core</span><span class="sxs-lookup"><span data-stu-id="0abbc-133">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="0abbc-134">适用于 Office 加载项的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="0abbc-134">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="0abbc-135">在加载项中请求获取 API 使用权限</span><span class="sxs-lookup"><span data-stu-id="0abbc-135">Requesting permissions for API use in add-ins</span></span>](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
