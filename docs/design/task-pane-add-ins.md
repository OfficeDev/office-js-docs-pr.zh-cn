---
title: Office 加载项中的任务窗格
description: 任务窗格允许用户访问界面控件，此类控件运行代码以修改文档或电子邮件，或显示数据源中的数据。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093754"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="93cbf-103">Office 加载项中的任务窗格</span><span class="sxs-lookup"><span data-stu-id="93cbf-103">Task panes in Office Add-ins</span></span>
 
<span data-ttu-id="93cbf-p101">任务窗格是接口图面，通常出现在 Word、PowerPoint、Excel 和 Outlook 中窗口的右侧。使用任务窗格，用户可以访问接口控件，以运行代码来修改文档或电子邮件，或显示数据源中的数据。如果不需要将功能直接嵌入文档，请使用任务窗格。</span><span class="sxs-lookup"><span data-stu-id="93cbf-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="93cbf-107">*图 1：典型任务窗格布局*</span><span class="sxs-lookup"><span data-stu-id="93cbf-107">*Figure 1. Typical task pane layout*</span></span>

![显示典型任务窗格布局的图像](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="93cbf-109">最佳做法</span><span class="sxs-lookup"><span data-stu-id="93cbf-109">Best practices</span></span>

|<span data-ttu-id="93cbf-110">**允许事项**</span><span class="sxs-lookup"><span data-stu-id="93cbf-110">**Do**</span></span>|<span data-ttu-id="93cbf-111">**禁止事项**</span><span class="sxs-lookup"><span data-stu-id="93cbf-111">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="93cbf-112">在标题中包括外接程序的名称。</span><span class="sxs-lookup"><span data-stu-id="93cbf-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="93cbf-113">请勿在标题中追加公司名称。</span><span class="sxs-lookup"><span data-stu-id="93cbf-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="93cbf-114">在标题中使用简短的描述性名称。</span><span class="sxs-lookup"><span data-stu-id="93cbf-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="93cbf-115">不要将字符串（例如 "外接程序"、"for Word" 或 "for Office"）追加到外接程序的标题。</span><span class="sxs-lookup"><span data-stu-id="93cbf-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="93cbf-116">在加载项顶部包括某些导航或命令元素，如命令栏或透视。</span><span class="sxs-lookup"><span data-stu-id="93cbf-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="93cbf-117">在外接程序底部包括品牌元素，如品牌栏，除非要在 Outlook 内使用外接程序。</span><span class="sxs-lookup"><span data-stu-id="93cbf-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||


## <a name="variants"></a><span data-ttu-id="93cbf-118">变量</span><span class="sxs-lookup"><span data-stu-id="93cbf-118">Variants</span></span>

<span data-ttu-id="93cbf-p102">以下图像显示了使用 Office 应用功能区的1366x768 分辨率的各种任务窗格大小。对于 Excel，需要额外的垂直空间来容纳编辑栏。</span><span class="sxs-lookup"><span data-stu-id="93cbf-p102">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="93cbf-121">*图 2：Office 2016 桌面任务窗格尺寸*</span><span class="sxs-lookup"><span data-stu-id="93cbf-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![显示尺寸为 1366x768 的桌面任务窗格的图像](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="93cbf-123">Excel - 320 x 455</span><span class="sxs-lookup"><span data-stu-id="93cbf-123">Excel - 320x455</span></span>
- <span data-ttu-id="93cbf-124">PowerPoint - 320 x 531</span><span class="sxs-lookup"><span data-stu-id="93cbf-124">PowerPoint - 320x531</span></span>
- <span data-ttu-id="93cbf-125">Word - 320 x 531</span><span class="sxs-lookup"><span data-stu-id="93cbf-125">Word - 320x531</span></span>
- <span data-ttu-id="93cbf-126">Outlook - 348x535</span><span class="sxs-lookup"><span data-stu-id="93cbf-126">Outlook - 348x535</span></span>

<br/>

<span data-ttu-id="93cbf-127">*图3。Office 任务窗格大小*</span><span class="sxs-lookup"><span data-stu-id="93cbf-127">*Figure 3. Office task pane sizes*</span></span>

![显示尺寸为 1366x768 的桌面任务窗格的图像](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="93cbf-129">Excel - 350 x 378</span><span class="sxs-lookup"><span data-stu-id="93cbf-129">Excel - 350x378</span></span>
- <span data-ttu-id="93cbf-130">PowerPoint - 348x391</span><span class="sxs-lookup"><span data-stu-id="93cbf-130">PowerPoint - 348x391</span></span>
- <span data-ttu-id="93cbf-131">Word - 329 x 445</span><span class="sxs-lookup"><span data-stu-id="93cbf-131">Word - 329x445</span></span>
- <span data-ttu-id="93cbf-132">Outlook（网页版）- 320x570</span><span class="sxs-lookup"><span data-stu-id="93cbf-132">Outlook (on the web) - 320x570</span></span>

## <a name="personality-menu"></a><span data-ttu-id="93cbf-133">“个性”菜单</span><span class="sxs-lookup"><span data-stu-id="93cbf-133">Personality menu</span></span>

<span data-ttu-id="93cbf-p103">“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。</span><span class="sxs-lookup"><span data-stu-id="93cbf-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="93cbf-136">对于 Windows，个性菜单尺寸为 12x32 像素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="93cbf-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="93cbf-137">*图 4：Windows 上的个性菜单*</span><span class="sxs-lookup"><span data-stu-id="93cbf-137">*Figure 4. Personality menu on Windows*</span></span>

![显示 Windows 桌面上个性菜单的图像](../images/personality-menu-win.png)

<span data-ttu-id="93cbf-139">对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将空间增加至 34x32 像素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="93cbf-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="93cbf-140">*图 5：Mac 上的个性菜单*</span><span class="sxs-lookup"><span data-stu-id="93cbf-140">*Figure 5. Personality menu on Mac*</span></span>

![显示 Mac 桌面上个性菜单的图像](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="93cbf-142">实现</span><span class="sxs-lookup"><span data-stu-id="93cbf-142">Implementation</span></span>

<span data-ttu-id="93cbf-143">有关实现任务窗格的示例，请参阅 GitHub 上的 [Excel 加载项 JS WoodGrove 支出趋势](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。</span><span class="sxs-lookup"><span data-stu-id="93cbf-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span> 


## <a name="see-also"></a><span data-ttu-id="93cbf-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="93cbf-144">See also</span></span>

- [<span data-ttu-id="93cbf-145">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="93cbf-145">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md) 
- [<span data-ttu-id="93cbf-146">适用于 Office 外接程序的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="93cbf-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)

