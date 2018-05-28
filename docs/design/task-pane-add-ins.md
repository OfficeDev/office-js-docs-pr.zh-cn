---
title: Office ?????????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d60af0a31b9f96be17aa55bda789d13c386d0a5b
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="0099c-102">Office ?????????</span><span class="sxs-lookup"><span data-stu-id="0099c-102">Task panes in Office Add-ins</span></span>
 
<span data-ttu-id="0099c-p101">??????????????? Word?PowerPoint?Excel ? Outlook ???????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="0099c-106">*? 1?????????*</span><span class="sxs-lookup"><span data-stu-id="0099c-106">*Figure 1. Typical task pane layout*</span></span>

![?????????????](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="0099c-108">????</span><span class="sxs-lookup"><span data-stu-id="0099c-108">Best practices</span></span>

|<span data-ttu-id="0099c-109">**????**</span><span class="sxs-lookup"><span data-stu-id="0099c-109">**Do**</span></span>|<span data-ttu-id="0099c-110">**????**</span><span class="sxs-lookup"><span data-stu-id="0099c-110">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="0099c-111">??????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-111">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="0099c-112">?????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-112">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="0099c-113">???????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-113">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="0099c-114">?????????????Add-in???For Word???for Office??????</span><span class="sxs-lookup"><span data-stu-id="0099c-114">Don't append strings such as ?Add-in,? ?For Word,? or ?for Office? to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="0099c-115">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-115">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="0099c-116">??????????????????????? Outlook ????????</span><span class="sxs-lookup"><span data-stu-id="0099c-116">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||


## <a name="variants"></a><span data-ttu-id="0099c-117">??</span><span class="sxs-lookup"><span data-stu-id="0099c-117">Variants</span></span>

<span data-ttu-id="0099c-p102">???????? 1366x768 ? Office ??????????????? Excel?????????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-p102">The following images show the various task pane sizes with the Office ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="0099c-120">*? 2?Office 2016 ????????*</span><span class="sxs-lookup"><span data-stu-id="0099c-120">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![????? 1366x768 ??????????](../images/add-in-taskpane-sizes-desktop.png)

- <span data-ttu-id="0099c-122">Excel - 320 x 455</span><span class="sxs-lookup"><span data-stu-id="0099c-122">Excel - 320x455</span></span>
- <span data-ttu-id="0099c-123">PowerPoint - 320 x 531</span><span class="sxs-lookup"><span data-stu-id="0099c-123">PowerPoint - 320x531</span></span>
- <span data-ttu-id="0099c-124">Word - 320 x 531</span><span class="sxs-lookup"><span data-stu-id="0099c-124">Word - 320x531</span></span>
- <span data-ttu-id="0099c-125">Outlook - 348x535</span><span class="sxs-lookup"><span data-stu-id="0099c-125">Outlook - 348x535</span></span>

<br/>

<span data-ttu-id="0099c-126">*? 3?Office 365 ??????*</span><span class="sxs-lookup"><span data-stu-id="0099c-126">*Figure 3. Office 365 task pane sizes*</span></span>

![????? 1366x768 ??????????](../images/add-in-taskpane-sizes-online.png)

- <span data-ttu-id="0099c-128">Excel - 350 x 378</span><span class="sxs-lookup"><span data-stu-id="0099c-128">Excel - 350x378</span></span>
- <span data-ttu-id="0099c-129">PowerPoint - 348x391</span><span class="sxs-lookup"><span data-stu-id="0099c-129">PowerPoint - 348x391</span></span>
- <span data-ttu-id="0099c-130">Word - 329 x 445</span><span class="sxs-lookup"><span data-stu-id="0099c-130">Word - 329x445</span></span>
- <span data-ttu-id="0099c-131">Outlook Web ?? - 320x570</span><span class="sxs-lookup"><span data-stu-id="0099c-131">Outlook Web App - 320x570</span></span>

## <a name="personality-menu"></a><span data-ttu-id="0099c-132">??????</span><span class="sxs-lookup"><span data-stu-id="0099c-132">Personality menu</span></span>

<span data-ttu-id="0099c-p103">???????????????????????????????? Windows ? Mac ??????????????</span><span class="sxs-lookup"><span data-stu-id="0099c-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="0099c-135">?? Windows???????? 12x32 ????????</span><span class="sxs-lookup"><span data-stu-id="0099c-135">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="0099c-136">*? 4?Windows ??????*</span><span class="sxs-lookup"><span data-stu-id="0099c-136">*Figure 4. Personality menu on Windows*</span></span>

![?? Windows ??????????](../images/personality-menu-win.png)

<span data-ttu-id="0099c-138">?? Mac?????????? 26x26 ?????????? 8 ?????????? 6 ??????????? 34x32 ????????</span><span class="sxs-lookup"><span data-stu-id="0099c-138">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="0099c-139">*? 5?Mac ??????*</span><span class="sxs-lookup"><span data-stu-id="0099c-139">*Figure 5. Personality menu on Mac*</span></span>

![?? Mac ??????????](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="0099c-141">??</span><span class="sxs-lookup"><span data-stu-id="0099c-141">Implementation</span></span>

<span data-ttu-id="0099c-142">??????????????? GitHub ?? [Excel ??? JS WoodGrove ????](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)?</span><span class="sxs-lookup"><span data-stu-id="0099c-142">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span> 


## <a name="see-also"></a><span data-ttu-id="0099c-143">????</span><span class="sxs-lookup"><span data-stu-id="0099c-143">See also</span></span>

- [<span data-ttu-id="0099c-144">Office ????? Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="0099c-144">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md) 
- [<span data-ttu-id="0099c-145">??? Office ????? UX ????</span><span class="sxs-lookup"><span data-stu-id="0099c-145">UX design patterns for Office Add-ins</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)


