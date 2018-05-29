---
title: ??? Office ????? UX ??????
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: c8ec23db5e7c4c571babff94bdc617b78340d965
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="ux-design-pattern-templates-for-office-add-ins"></a><span data-ttu-id="97b3e-102">Office ??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-102">UX design pattern templates for Office Add-ins</span></span>

<span data-ttu-id="97b3e-103">[??? Office ????? UX ??????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "??? Office ????? UX ??????")(#???-office-?????-ux-??????) ?? HTML?JavaScript ? CSS ??????????????????? UX?</span><span class="sxs-lookup"><span data-stu-id="97b3e-103">The [UX design patterns for Office Add-ins project](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "UX design patterns for Office Add-ins project") includes HTML, JavaScript, and CSS files that you can use to create the UX for your add-in.</span></span>   

<span data-ttu-id="97b3e-104">UX ?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-104">Use the UX design patterns project to:</span></span>

* <span data-ttu-id="97b3e-105">????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-105">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="97b3e-106">?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-106">Apply design best practices.</span></span>
* <span data-ttu-id="97b3e-107">???[Office UI Fabric](https://dev.office.com/fabric#/get-started)???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-107">Incorporate [Office UI Fabric](https://dev.office.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="97b3e-108">?????????? Office UI ????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-108">Build add-ins that visually integrate with the default Office UI.</span></span>  

## <a name="using-the-ux-design-patterns"></a><span data-ttu-id="97b3e-109">?? UX ????</span><span class="sxs-lookup"><span data-stu-id="97b3e-109">Using the UX design patterns</span></span>

<span data-ttu-id="97b3e-110">???? [Office ?????????](https://aka.ms/addins_toolkit) ? [???????](https://aka.ms/fabric-toolkit) ???????????? Office ?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-110">You can use the [Office Add-ins Design Toolkit](https://aka.ms/addins_toolkit) together with the [Fabric Design Toolkit](https://aka.ms/fabric-toolkit) as a guide when you design your own Office Add-in.</span></span> <span data-ttu-id="97b3e-111">??????? [???](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) ????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-111">You can also add the [source code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) directly to your project.</span></span>

<span data-ttu-id="97b3e-112">?????????????? UI ???</span><span class="sxs-lookup"><span data-stu-id="97b3e-112">To use the specifications to build a mock-up of your own add-in UI:</span></span>

1. <span data-ttu-id="97b3e-113">???????????????? UI?</span><span class="sxs-lookup"><span data-stu-id="97b3e-113">Download design assets files and begin designing your own UI:</span></span>
    * [<span data-ttu-id="97b3e-114">Office ?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-114">Office Add-ins Design Toolkit</span></span>](https://aka.ms/addins_toolkit)
    * [<span data-ttu-id="97b3e-115">???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-115">Fabric Design Toolkit</span></span>](https://aka.ms/fabric-toolkit)

2. <span data-ttu-id="97b3e-116">???????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-116">Refer to the following articles for guidance:</span></span>
    * <span data-ttu-id="97b3e-117">[?? Office ????](add-in-design.md) ?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-117">Best practices for [Designing your Office Add-ins](add-in-design.md)</span></span>
    * [<span data-ttu-id="97b3e-118">Office UI Fabric ???</span><span class="sxs-lookup"><span data-stu-id="97b3e-118">Office UI Fabric Toolkits</span></span>](https://developer.microsoft.com/en-us/fabric#/resources)

> [!NOTE]
> <span data-ttu-id="97b3e-119">????????????? UX ???????? UX ????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-119">Some UX patterns in the Add-ins Design Toolkit do not match the UX design patterns detailed below.</span></span> <span data-ttu-id="97b3e-120">????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-120">We're planning to release updated documentation that will align with the toolkit.</span></span>

<span data-ttu-id="97b3e-121">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-121">To add the source code:</span></span>

1. <span data-ttu-id="97b3e-122">?? [??? Office ????? UX ?????????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "??? Office ????? UX ??????")?</span><span class="sxs-lookup"><span data-stu-id="97b3e-122">Clone the [UX design patterns for Office Add-ins project repo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "UX design patterns for Office Add-ins project").</span></span>
2. <span data-ttu-id="97b3e-123">? [assets ???](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets)(#assets-???) ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-123">Copy the [assets folder](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets), and the code folder for the individual pattern you choose to your add-in project.</span></span>  
3. <span data-ttu-id="97b3e-p103">?????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p103">Incorporate the individual pattern into your add-in. For example:</span></span>
    - <span data-ttu-id="97b3e-126">???????????????? URL?</span><span class="sxs-lookup"><span data-stu-id="97b3e-126">Edit the source location or add-in command URL in the manifest.</span></span>
    - <span data-ttu-id="97b3e-127">? UX ??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-127">Use the UX design pattern as a template for other pages.</span></span>
    - <span data-ttu-id="97b3e-128">??????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-128">Link to or from the UX design pattern.</span></span>

> [!NOTE]
> <span data-ttu-id="97b3e-129">??????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-129">Some UX pattern specifications do not match the source code.</span></span> <span data-ttu-id="97b3e-130">????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-130">We're working hard to bring all assets into alignment.</span></span> <span data-ttu-id="97b3e-131">??????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-131">Also notice that some specifications are presented as archived.</span></span> <span data-ttu-id="97b3e-132">?????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-132">We're assessing these archived specifications for value to the platform.</span></span> <span data-ttu-id="97b3e-133">????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-133">Each pattern aims to represent a unique template and pattern of interaction.</span></span> <span data-ttu-id="97b3e-134">?????????????? Office Fabric UI ???</span><span class="sxs-lookup"><span data-stu-id="97b3e-134">The patterns should not overlap with each other and should be well differentiated from Office Fabric UI components.</span></span>


## <a name="types-of-ux-design-patterns"></a><span data-ttu-id="97b3e-135">UX ???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-135">Types of UX design patterns</span></span>
### <a name="generic-pages"></a><span data-ttu-id="97b3e-136">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-136">Generic pages</span></span>

<span data-ttu-id="97b3e-p105">???????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p105">Generic page templates can be applied to any page in your add-in and don't have a special purpose. An example of a special purpose page, would be any of the first-run patterns. The following list describes the generic pages available:</span></span>

* <span data-ttu-id="97b3e-140">**???** - ??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-140">**Landing page** - A standard add-in page, for example the page a user lands on after a first-run experience or sign-in process.</span></span> 
    * <span data-ttu-id="97b3e-141">??????????? [Office ????](add-in-design-language.md)????</span><span class="sxs-lookup"><span data-stu-id="97b3e-141">Learn about guidelines for adopting the [Office design language](add-in-design-language.md) in your add-in.</span></span>
    * [<span data-ttu-id="97b3e-142">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-142">Landing page code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page)
* <span data-ttu-id="97b3e-143">**?????????** - ???????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-143">**Brand image in brand bar** - The landing page with an image in the footer that represents your brand.</span></span> 
    * [<span data-ttu-id="97b3e-144">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-144">Brand bar specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/brand-bar.md)
    * [<span data-ttu-id="97b3e-145">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-145">Brand bar code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar)

<table>
 <tr><th><span data-ttu-id="97b3e-146">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-146">Landing</span></span></th><th><span data-ttu-id="97b3e-147">???</span><span class="sxs-lookup"><span data-stu-id="97b3e-147">Brand Bar</span></span></th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../images/landing-pages.png" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../images/word-brand-bar.png" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>
 
### <a name="first-run-experience"></a><span data-ttu-id="97b3e-148">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-148">First-run experience</span></span>

<span data-ttu-id="97b3e-p106">????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p106">A first-run experience is the experience a user has when they open your add-in for the first time. The following first-run design pattern templates are available:</span></span> 

* <span data-ttu-id="97b3e-151">**????** - ??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-151">**Steps to start** - Provides users with an ordered list of steps to perform to get started using your add-in.</span></span> 
    * <span data-ttu-id="97b3e-152">[??????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-152">[Steps to start specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-153">??????????????[???????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)??</span><span class="sxs-lookup"><span data-stu-id="97b3e-153">As we assess its value, see [First-Run Value specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md).)</span></span>  
    * [<span data-ttu-id="97b3e-154">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-154">Steps to start code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step)
* <span data-ttu-id="97b3e-155">**?** - ??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-155">**Value** - Communicates your add-in's value proposition.</span></span>
    * [<span data-ttu-id="97b3e-156">???</span><span class="sxs-lookup"><span data-stu-id="97b3e-156">Value specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)
    * [<span data-ttu-id="97b3e-157">???</span><span class="sxs-lookup"><span data-stu-id="97b3e-157">Value code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)
* <span data-ttu-id="97b3e-158">**??** - ??????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-158">**Video** - Shows users a video before they start using your add-in.</span></span>
    * [<span data-ttu-id="97b3e-159">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-159">Video specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/video-placemat.md)
    * [<span data-ttu-id="97b3e-160">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-160">Video code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)
* <span data-ttu-id="97b3e-161">**??** - ??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-161">**Walkthrough** - Takes users through a series of features or information before they start using the add-in.</span></span>
    * <span data-ttu-id="97b3e-162">[????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md)????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-162">[Carousel specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md) (Note that this UX design pattern has been renamed to "Carousel."</span></span> <span data-ttu-id="97b3e-163">???????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-163">Former specifications refered to it as a "Paging Panel."</span></span> <span data-ttu-id="97b3e-164">??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-164">Code assets refer to it as a "First-run Walkthrough."</span></span> 
    * [<span data-ttu-id="97b3e-165">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-165">Walkthrough code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough)

<span data-ttu-id="97b3e-166">[AppSource](https://docs.microsoft.com/en-us/office/dev/store/use-the-seller-dashboard-to-submit-to-the-office-store) ??????????????????????????? UI?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-166">[AppSource](https://docs.microsoft.com/en-us/office/dev/store/use-the-seller-dashboard-to-submit-to-the-office-store) has a system that manages trial versions of an add-in, but if you want to control the UI of the trial experience for your add-in, use the following patterns:</span></span>

* <span data-ttu-id="97b3e-167">**???** - ??????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-167">**Trial** - Shows users how to get started with a trial version of your add-in.</span></span>
    * <span data-ttu-id="97b3e-168">[?????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-168">[Trial specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-169">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-169">As we assess its value, refer to this PDF.)</span></span>
    * [<span data-ttu-id="97b3e-170">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-170">Trial code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat)
* <span data-ttu-id="97b3e-p110">**?????** - ?????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p110">**Trial feature** - Advises users that the feature they are trying to use is not available in the trial version of the add-in. Alternatively, if your add-in is free but it includes a feature that requires a subscription, consider using this pattern. You might also use this pattern to provide a downgraded experience after a trial has ended.</span></span>
    * <span data-ttu-id="97b3e-174">[???????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-174">[Trial feature specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-175">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-175">As we assess its value, refer to this PDF.)</span></span>
    * [<span data-ttu-id="97b3e-176">???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-176">Trial feature code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature)

> [!IMPORTANT]
> <span data-ttu-id="97b3e-177">?????????????????? AppSource ????????????????????????????????****???</span><span class="sxs-lookup"><span data-stu-id="97b3e-177">If you decide to manage your own trial, and not use AppSource to manage the trial, make sure to include the **Additional purchase may be required** tag in the testing notes in the seller dashboard.</span></span>

<span data-ttu-id="97b3e-p112">???????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p112">Consider whether showing users the first-run experience once or many times is important to your scenario. For example, if users use your add-in periodically, they might forget how to use it, and it might be helpful to see the first-run experience more than once.</span></span> 

 <table>
 <tr><th><span data-ttu-id="97b3e-180">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-180">Steps to Start</span></span></th><th><span data-ttu-id="97b3e-181">?</span><span class="sxs-lookup"><span data-stu-id="97b3e-181">Value</span></span></th><th><span data-ttu-id="97b3e-182">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-182">Video</span></span></th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../images/instruction-steps.png" alt="instruction steps" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../images/value-placemats.png" alt="value placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../images/video-placemats.png" alt="video placemat" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th><span data-ttu-id="97b3e-183">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-183">Walkthrough first page</span></span></th><th><span data-ttu-id="97b3e-184">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-184">Trial</span></span></th><th><span data-ttu-id="97b3e-185">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-185">Trial feature</span></span></th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../images/walkthrough01.png" alt="walkthrough 1" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../images/trial-placemats.png" alt="trial placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../images/trial-placemats-feature.png" alt="trial placemat feature" style="width: 250px;"/></A></td></tr>
 </table> 

### <a name="navigation"></a><span data-ttu-id="97b3e-186">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-186">Navigation</span></span>

<span data-ttu-id="97b3e-p113">?????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p113">Users need to navigate between the different pages of your add-in. The following navigation templates show different options you can use to organize pages and commands in your add-in.</span></span>

* <span data-ttu-id="97b3e-p114">**????????????** - ??????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p114">**Back button and Next page** - Shows a task pane with Back and Next page buttons. Use this pattern to ensure users follow an ordered series of steps.</span></span>
    * [<span data-ttu-id="97b3e-191">??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-191">Back Button and Next Page specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/back-button.md)
    * [<span data-ttu-id="97b3e-192">??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-192">Back Button and Next Page code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button) 
* <span data-ttu-id="97b3e-193">**??** - ?????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-193">**Navigation** - Shows a menu, commonly referred to as the hamburger menu, with page menu items in a task pane.</span></span> 
    * [<span data-ttu-id="97b3e-194">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-194">Navigation specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/contextual-menu.md)
    * [<span data-ttu-id="97b3e-195">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-195">Navigation code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation) 
* <span data-ttu-id="97b3e-p115">**????????** - ???????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p115">**Navigation with commands** - Shows the hamburger menu with command (or action) buttons in a task pane. Use this pattern when you want to provide navigation and command options together.</span></span> 
    * [<span data-ttu-id="97b3e-198">?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-198">Navigation with commands specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/command-bar.md)
    * [<span data-ttu-id="97b3e-199">??????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-199">Navigation with commands code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands)
* <span data-ttu-id="97b3e-p116">**??** - ?????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p116">**Pivot** - Shows Pivot navigation inside of a task pane. Use pivot navigation to allow users to navigate between different content.</span></span>
    * [<span data-ttu-id="97b3e-202">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-202">Pivot specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/pivot.md)
    * [<span data-ttu-id="97b3e-203">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-203">Pivot code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot)
* <span data-ttu-id="97b3e-p117">**????** - ??????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p117">**Tab bar** - Shows navigation using buttons with vertically stacked text and icons. Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>
    * [<span data-ttu-id="97b3e-206">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-206">Tab bar specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/tab-bar.md)
    * [<span data-ttu-id="97b3e-207">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-207">Tab bar code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar) 

<table>
<tr><th><span data-ttu-id="97b3e-208">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-208">Back button</span></span></th><th><span data-ttu-id="97b3e-209">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-209">Navigation</span></span></th><th><span data-ttu-id="97b3e-210">????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-210">Navigation with commands</span></span></th></tr>
<tr>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
        <img src="../images/back-button.png" alt="back button" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
        <img src="../images/navigation.png" alt="navigation" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
        <img src="../images/navigation-commands.png" alt="navigation with commands" style="width: 250px;"/></A>
    </td>
</tr>
 </table>

<table>
<tr><th><span data-ttu-id="97b3e-211">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-211">Pivot</span></span></th><th><span data-ttu-id="97b3e-212">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-212">Tab bar</span></span></th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../images/pivot.png" alt="pivot navigation" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../images/tab-bar.png" alt="tab bar" style="width: 250px;"/></A></td>
</tr>
 </table>

### <a name="notifications"></a><span data-ttu-id="97b3e-213">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-213">Notifications</span></span>

<span data-ttu-id="97b3e-p118">??????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p118">Your add-in can notify users of events, such as errors, or of progress in a variety of ways. The following notification templates are available:</span></span> 

* <span data-ttu-id="97b3e-p119">**??????** - ???????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p119">**Embedded dialog box** - Shows a dialog box inside the task pane that provides information and, optionally, an interactive experience, using buttons or other controls. Consider using one to prompt a user to confirm an action. Use the Embedded dialog pattern when you want to keep the user experience in the task pane.</span></span>
    * [<span data-ttu-id="97b3e-219">????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-219">Embedded dialog box specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/embedded-dialog.md)
    * [<span data-ttu-id="97b3e-220">????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-220">Embedded dialog box code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog)
* <span data-ttu-id="97b3e-p120">**????** - ??????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p120">**Inline message** - Indicates error, success, or information, and can appear at a specified location in the task pane. For example, if a user enters an incorrectly formatted email address in a text box, an error message appears just below the text box.</span></span> 
    * <span data-ttu-id="97b3e-223">[??????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-223">[Inline message specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-224">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-224">As we assess its value, refer to this PDF.)</span></span>
    * [<span data-ttu-id="97b3e-225">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-225">Inline message code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message)
* <span data-ttu-id="97b3e-p122">**????** - ????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p122">**Message banner** - Provides information and, optionally, a simple call to action, in a banner that can be collapsed to a single line, expanded to multiple lines, or dismissed. Use message banners to report a service update or a helpful tip when the add-in starts.</span></span> 
    * <span data-ttu-id="97b3e-228">[??????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-228">[Message banner specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-229">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-229">As we assess its value, refer to this PDF.)</span></span>
    * [<span data-ttu-id="97b3e-230">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-230">Message banner code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner)
* <span data-ttu-id="97b3e-p124">**???** - ????????????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p124">**Progress bar** - Indicates the progress of a long-running, synchronous process, such as a configuration task that must complete before the user can take any further action. It is a separate interstitial page that also reinforces the add-in brand. Use a progress bar when the process can send periodic measures of how far along it is back to the add-in.</span></span>
    * [<span data-ttu-id="97b3e-234">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-234">Progress bar specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/progress-indicator.md)
    * [<span data-ttu-id="97b3e-235">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-235">Progress bar code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar)
* <span data-ttu-id="97b3e-p125">**???** - ?????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p125">**Spinner** - Indicates that a long-running, synchronous process is underway, but provides no indication of how far along it is. It is a separate interstitial page that also reinforces the add-in brand. Use a spinner when the add-in cannot know reliably how far along a process is.</span></span> 
    * [<span data-ttu-id="97b3e-239">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-239">Spinner specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/spinner.md)
    * [<span data-ttu-id="97b3e-240">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-240">Spinner code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner)
* <span data-ttu-id="97b3e-p126">**Toast** - ???????????????????????????????toast ??????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p126">**Toast** - Provides a brief message that fades away after a few seconds. Because the user might not see the message, use toast only for nonessential information. It is a good choice for notifying users of an event in a remote system, such as the receipt of an email.</span></span>
    * [<span data-ttu-id="97b3e-244">Toast ??</span><span class="sxs-lookup"><span data-stu-id="97b3e-244">Toast specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/toast.md)
    * [<span data-ttu-id="97b3e-245">Toast ??</span><span class="sxs-lookup"><span data-stu-id="97b3e-245">Toast code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast)

 <table>
 <tr><th><span data-ttu-id="97b3e-246">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-246">Embedded dialog</span></span></th><th><span data-ttu-id="97b3e-247">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-247">Inline message</span></span></th><th><span data-ttu-id="97b3e-248">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-248">Message banner</span></span></th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../images/embedded-dialogs.png" alt="embedded dialog" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../images/inline-messages.png" alt="inline message" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../images/message-banners.png" alt="message banner" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th><span data-ttu-id="97b3e-249">???</span><span class="sxs-lookup"><span data-stu-id="97b3e-249">Progress bar</span></span></th><th><span data-ttu-id="97b3e-250">???</span><span class="sxs-lookup"><span data-stu-id="97b3e-250">Spinner</span></span></th><th><span data-ttu-id="97b3e-251">Toast</span><span class="sxs-lookup"><span data-stu-id="97b3e-251">Toast</span></span></th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../images/progress-bars.png" alt="progress bar" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../images/logo-spinner.png" alt="spinner" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../images/toast-header.png" alt="toast" style="width: 250px;"/></A></td></tr>
 </table>
 


### <a name="general-components"></a><span data-ttu-id="97b3e-252">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-252">General components</span></span>

<span data-ttu-id="97b3e-253">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-253">The following are general components that you can use in your add-ins in a variety of scenarios.</span></span>  

#### <a name="client-dialog-boxes"></a><span data-ttu-id="97b3e-254">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-254">Client dialog boxes</span></span>

<span data-ttu-id="97b3e-p127">????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p127">Client dialog boxes provide another way for users to work with your add-in outside of a task pane. The following dialog box templates are available:</span></span>

* <span data-ttu-id="97b3e-p128">**Typeramp ???** - ?????????????Typeramp ????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p128">**Typeramp dialog box** - Shows a dialog box with textual content. Use the typeramp dialog to display elaborative information to users.</span></span> 
    * <span data-ttu-id="97b3e-259">?????? [Office ????????](dialog-boxes.md)????? [Office ???????](add-in-design-language.md#typography)???</span><span class="sxs-lookup"><span data-stu-id="97b3e-259">Learn about designing [dialog boxes in Office Add-ins](dialog-boxes.md). Also follow our guidelines for [Typography in Office Add-ins](add-in-design-language.md#typography).</span></span>
    * [<span data-ttu-id="97b3e-260">Typeramp ?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-260">Typeramp dialog box code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp)
* <span data-ttu-id="97b3e-261">**?????** - ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-261">**Alert dialog box** - Shows an alert box with important information, such as errors or notifications, to users.</span></span>  
    * <span data-ttu-id="97b3e-262">[???????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-262">[Alert dialog box specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-263">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-263">As we assess its value, refer to this PDF.)</span></span>
    * [<span data-ttu-id="97b3e-264">???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-264">Alert dialog box code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert)
* <span data-ttu-id="97b3e-p130">**?????** - ????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p130">**Navigation dialog box** - Shows a dialog box with navigation. Use the navigation dialog box to allow users to navigate between different content.</span></span> 
    * <span data-ttu-id="97b3e-267">?????? [Office ????????](dialog-boxes.md)?????[??? Office ?????? Office UI Fabric ????](pivot.md)?</span><span class="sxs-lookup"><span data-stu-id="97b3e-267">Learn about designing [dialog boxes in Office Add-ins](dialog-boxes.md). Also learn about using Office UI Fabric [Pivot components in Office Add-ins](pivot.md).</span></span>
    * [<span data-ttu-id="97b3e-268">???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-268">Navigation dialog box code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)

<table>
 <tr><th><span data-ttu-id="97b3e-269">Typeramp ???</span><span class="sxs-lookup"><span data-stu-id="97b3e-269">Typeramp dialog</span></span></th><th><span data-ttu-id="97b3e-270">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-270">Alert dialog</span></span></th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../images/typeramp-dialog.png" alt="typeramp dialog" width="400"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../images/alert-dialog.png" alt="alert dialog" width="400"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th><span data-ttu-id="97b3e-271">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-271">Navigation dialog</span></span></th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../images/navigation-dialog.png" alt="navigation dialog" width="450"/></A></td></tr>
</tr>
 </table>


#### <a name="feedback-and-ratings"></a><span data-ttu-id="97b3e-272">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-272">Feedback and ratings</span></span>

<span data-ttu-id="97b3e-p131">??????????????????????? AppSource ??????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p131">To improve the visibility and adoption of your add-in, it is helpful to provide users with the ability to rate and review your add-in in AppSource. This pattern shows two methods for presenting feedback and ratings from within the add-in:</span></span>

- <span data-ttu-id="97b3e-275">??????? - ?????????????????????????????**????**????</span><span class="sxs-lookup"><span data-stu-id="97b3e-275">User-initiated feedback - A user chooses to send feedback by using either the navigation menu (for example, using the **Send Feedback** link) or an icon on the footer.</span></span>
- <span data-ttu-id="97b3e-276">??????? - ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-276">System-initiated feedback - After the add-in runs three times, the user is prompted to provide feedback via a Message Banner.</span></span>

<span data-ttu-id="97b3e-277">???????????????????????? AppSource ???</span><span class="sxs-lookup"><span data-stu-id="97b3e-277">Either method opens a dialog box that contains the AppSource page for the add-in.</span></span>

* <span data-ttu-id="97b3e-278">[???????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-278">[Feedback and ratings specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-279">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-279">As we assess its value, refer to this PDF.)</span></span>
* [<span data-ttu-id="97b3e-280">???????</span><span class="sxs-lookup"><span data-stu-id="97b3e-280">Feedback and ratings code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store)

> [!IMPORTANT]
> <span data-ttu-id="97b3e-p133">?????? AppSource ???????? URL ????????? AppSource ?? URL?</span><span class="sxs-lookup"><span data-stu-id="97b3e-p133">This pattern currently points to the AppSource home page. Be sure to update this URL to the URL of your add-in's page in AppSource.</span></span>


 <table>
 <tr><th><span data-ttu-id="97b3e-283">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-283">Feedback and ratings</span></span></th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../images/feedback-rating.png" alt="Feedback and Ratings" style="width: 350px;"/></A></td></tr>
</tr>
 </table>

#### <a name="settings-and-privacy"></a><span data-ttu-id="97b3e-284">?????</span><span class="sxs-lookup"><span data-stu-id="97b3e-284">Settings and privacy</span></span>

<span data-ttu-id="97b3e-p134">????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p134">Add-ins may need a Settings page that allows users to configure settings that control the behavior of the add-in. Also, you may want to provide users with the privacy policies your add-in adheres to.</span></span> 

* <span data-ttu-id="97b3e-p135">**??** - ??????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-p135">**Settings** - Shows a task pane with configuration components that controls the behavior of the add-in. A settings page provides options for the user to choose.</span></span>
    * [<span data-ttu-id="97b3e-289">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-289">Settings specification</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/settings.md)
    * [<span data-ttu-id="97b3e-290">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-290">Settings code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)
* <span data-ttu-id="97b3e-291">**????** - ????????????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-291">**Privacy policy** - Shows task pane with important information about privacy policies.</span></span> 
    * <span data-ttu-id="97b3e-292">[??????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf)??????????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-292">[Privacy Policy specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf) (This UX design pattern has been archived.</span></span> <span data-ttu-id="97b3e-293">???????????????? PDF??</span><span class="sxs-lookup"><span data-stu-id="97b3e-293">As we assess its value, refer to this PDF.)</span></span>
    * [<span data-ttu-id="97b3e-294">??????</span><span class="sxs-lookup"><span data-stu-id="97b3e-294">Privacy Policy code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)

<table>
 <tr><th><span data-ttu-id="97b3e-295">??</span><span class="sxs-lookup"><span data-stu-id="97b3e-295">Settings</span></span></th><th><span data-ttu-id="97b3e-296">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-296">Privacy Policy</span></span></th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/privacy-policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## <a name="see-also"></a><span data-ttu-id="97b3e-297">????</span><span class="sxs-lookup"><span data-stu-id="97b3e-297">See also</span></span>

* [<span data-ttu-id="97b3e-298">Office ?????????</span><span class="sxs-lookup"><span data-stu-id="97b3e-298">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="97b3e-299">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="97b3e-299">Office UI Fabric</span></span>](http://dev.office.com/fabric/)
