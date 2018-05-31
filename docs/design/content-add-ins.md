---
title: 内容 Office 外接程序
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd0dcea7a3f37175a48946fc9dcd61d2b89f9c08
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437260"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="e50a9-102">内容 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e50a9-102">Content Office Add-ins</span></span>

<span data-ttu-id="e50a9-p101">内容外接程序这种图面可被直接嵌入 Word、Excel 或 PowerPoint 文档中。内容外接程序让用户访问运行代码以修改文档或显示数据源中数据的界面控件。在你要将功能直接嵌入文档时，请使用内容加载项。</span><span class="sxs-lookup"><span data-stu-id="e50a9-p101">Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="e50a9-106">*图 1：内容加载项的典型布局*</span><span class="sxs-lookup"><span data-stu-id="e50a9-106">*Figure 1. Typical layout for content add-ins*</span></span>

![显示内容加载项的典型布局的示例图像。](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="e50a9-108">最佳做法</span><span class="sxs-lookup"><span data-stu-id="e50a9-108">Best practices</span></span>

- <span data-ttu-id="e50a9-109">在外接程序顶部包括某些导航或命令元素，如命令栏或透视。</span><span class="sxs-lookup"><span data-stu-id="e50a9-109">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="e50a9-110">包括位于外接程序底部的品牌元素，如品牌栏（仅适用于 Word、Excel 和 PowerPoint 外接程序）。</span><span class="sxs-lookup"><span data-stu-id="e50a9-110">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="e50a9-111">变量</span><span class="sxs-lookup"><span data-stu-id="e50a9-111">Variants</span></span>

<span data-ttu-id="e50a9-112">Office 2016 桌面和 Office 365 中的 Word、Excel 和 PowerPoint 的内容外接程序大小由用户指定。</span><span class="sxs-lookup"><span data-stu-id="e50a9-112">Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="e50a9-113">“个性”菜单</span><span class="sxs-lookup"><span data-stu-id="e50a9-113">Personality menu</span></span>

<span data-ttu-id="e50a9-p102">“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。</span><span class="sxs-lookup"><span data-stu-id="e50a9-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="e50a9-116">对于 Windows，“个性”菜单尺寸为 12 x 32 像素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="e50a9-116">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="e50a9-117">*图 2：Windows 上的个性菜单*</span><span class="sxs-lookup"><span data-stu-id="e50a9-117">*Figure 2. Personality menu on Windows*</span></span> 

![显示 Windows 桌面上个性菜单的图像](../images/personality-menu-win.png)


<span data-ttu-id="e50a9-119">对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将占用空间增加至 34x32 像素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="e50a9-119">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="e50a9-120">*图 3：Mac 上的个性菜单*</span><span class="sxs-lookup"><span data-stu-id="e50a9-120">*Figure 3. Personality menu on Mac*</span></span>

![显示 Mac 桌面上个性菜单的图像](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="e50a9-122">实现</span><span class="sxs-lookup"><span data-stu-id="e50a9-122">Implementation</span></span>

<span data-ttu-id="e50a9-123">有关实现内容加载项的示例，请参阅 GitHub 上的 [Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。</span><span class="sxs-lookup"><span data-stu-id="e50a9-123">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="e50a9-124">支持注意事项</span><span class="sxs-lookup"><span data-stu-id="e50a9-124">Support considerations</span></span>
- <span data-ttu-id="e50a9-125">检查 Office 加载项是否适用于[特定 Office 主机平台](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)。</span><span class="sxs-lookup"><span data-stu-id="e50a9-125">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="e50a9-126">一些内容外接程序可能会要求用户“信任”外接程序，以便对 Excel 或 PowerPoint 执行读取和写入操作。</span><span class="sxs-lookup"><span data-stu-id="e50a9-126">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="e50a9-127">可以在外接程序清单中声明要让用户拥有的[权限级别](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="e50a9-127">You can declare what [level of permissions](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your use to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="e50a9-128">Office 2013 版本及更高版本中的 Excel 和 PowerPoint 支持内容外接程序。</span><span class="sxs-lookup"><span data-stu-id="e50a9-128">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="e50a9-129">如果在不支持 Office Web 加载项的 Office 版本中打开加载项，加载项会显示为图像。</span><span class="sxs-lookup"><span data-stu-id="e50a9-129">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="e50a9-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e50a9-130">See also</span></span>
- [<span data-ttu-id="e50a9-131">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="e50a9-131">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="e50a9-132">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="e50a9-132">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="e50a9-133">Office 加载项的用户体验设计模式</span><span class="sxs-lookup"><span data-stu-id="e50a9-133">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/ux-design-patterns)
- [<span data-ttu-id="e50a9-134">在内容加载项和任务窗格加载项中请求获取 API 使用权限</span><span class="sxs-lookup"><span data-stu-id="e50a9-134">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
