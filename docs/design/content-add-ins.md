---
title: ?? Office ????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd0dcea7a3f37175a48946fc9dcd61d2b89f9c08
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="content-office-add-ins"></a><span data-ttu-id="e310d-102">?? Office ???</span><span class="sxs-lookup"><span data-stu-id="e310d-102">Content Office Add-ins</span></span>

<span data-ttu-id="e310d-p101">???????????????? Word?Excel ? PowerPoint ??????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="e310d-p101">Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="e310d-106">*? 1???????????*</span><span class="sxs-lookup"><span data-stu-id="e310d-106">*Figure 1. Typical layout for content add-ins*</span></span>

![??????????????????](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="e310d-108">????</span><span class="sxs-lookup"><span data-stu-id="e310d-108">Best practices</span></span>

- <span data-ttu-id="e310d-109">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="e310d-109">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="e310d-110">????????????????????????? Word?Excel ? PowerPoint ??????</span><span class="sxs-lookup"><span data-stu-id="e310d-110">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="e310d-111">??</span><span class="sxs-lookup"><span data-stu-id="e310d-111">Variants</span></span>

<span data-ttu-id="e310d-112">Office 2016 ??? Office 365 ?? Word?Excel ? PowerPoint ???????????????</span><span class="sxs-lookup"><span data-stu-id="e310d-112">Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="e310d-113">??????</span><span class="sxs-lookup"><span data-stu-id="e310d-113">Personality menu</span></span>

<span data-ttu-id="e310d-p102">???????????????????????????????? Windows ? Mac ??????????????</span><span class="sxs-lookup"><span data-stu-id="e310d-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="e310d-116">?? Windows?????????? 12 x 32 ????????</span><span class="sxs-lookup"><span data-stu-id="e310d-116">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="e310d-117">*? 2?Windows ??????*</span><span class="sxs-lookup"><span data-stu-id="e310d-117">*Figure 2. Personality menu on Windows*</span></span> 

![?? Windows ??????????](../images/personality-menu-win.png)


<span data-ttu-id="e310d-119">?? Mac?????????? 26x26 ?????????? 8 ?????????? 6 ????????????? 34x32 ????????</span><span class="sxs-lookup"><span data-stu-id="e310d-119">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="e310d-120">*? 3?Mac ??????*</span><span class="sxs-lookup"><span data-stu-id="e310d-120">*Figure 3. Personality menu on Mac*</span></span>

![?? Mac ??????????](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="e310d-122">??</span><span class="sxs-lookup"><span data-stu-id="e310d-122">Implementation</span></span>

<span data-ttu-id="e310d-123">???????????????? GitHub ?? [Excel ????? Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)?</span><span class="sxs-lookup"><span data-stu-id="e310d-123">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="e310d-124">??????</span><span class="sxs-lookup"><span data-stu-id="e310d-124">Support considerations</span></span>
- <span data-ttu-id="e310d-125">?? Office ????????[?? Office ????](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)?</span><span class="sxs-lookup"><span data-stu-id="e310d-125">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="e310d-126">??????????????????????????? Excel ? PowerPoint ??????????</span><span class="sxs-lookup"><span data-stu-id="e310d-126">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="e310d-127">???????????????????[????](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)?</span><span class="sxs-lookup"><span data-stu-id="e310d-127">You can declare what [level of permissions](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your use to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="e310d-128">Office 2013 ????????? Excel ? PowerPoint ?????????</span><span class="sxs-lookup"><span data-stu-id="e310d-128">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="e310d-129">?????? Office Web ???? Office ???????????????????</span><span class="sxs-lookup"><span data-stu-id="e310d-129">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="e310d-130">????</span><span class="sxs-lookup"><span data-stu-id="e310d-130">See also</span></span>
- [<span data-ttu-id="e310d-131">Office ???????????</span><span class="sxs-lookup"><span data-stu-id="e310d-131">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="e310d-132">Office ????? Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="e310d-132">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="e310d-133">Office ????????????</span><span class="sxs-lookup"><span data-stu-id="e310d-133">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/ux-design-patterns)
- [<span data-ttu-id="e310d-134">??????????????????? API ????</span><span class="sxs-lookup"><span data-stu-id="e310d-134">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
