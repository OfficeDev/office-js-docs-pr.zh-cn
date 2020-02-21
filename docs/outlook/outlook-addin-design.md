---
title: Outlook 外接程序设计
description: 可帮助你设计和构建引人入胜的外接程序的准则，通过遵循这些准则，你可以将自己的最佳的应用引入到 Windows、Web、iOS、Mac 和 Android 上的 Outlook 中。
ms.date: 06/24/2019
localization_priority: Priority
ms.openlocfilehash: efedeb32643bff12e167931ac4da80fdcc2c277f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165910"
---
# <a name="outlook-add-in-design-guidelines"></a><span data-ttu-id="2d7fa-103">Outlook 外接程序设计准则</span><span class="sxs-lookup"><span data-stu-id="2d7fa-103">Outlook add-in design guidelines</span></span>

<span data-ttu-id="2d7fa-p101">外接程序是一种可供合作伙伴在我们的核心功能集之外进一步扩展 Outlook 功能的绝佳方式。通过外接程序，用户无需离开收件箱即可访问第三方体验、任务和内容。安装后，Outlook 外接程序将在所有平台和设备上可用。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p101">Add-ins are a great way for partners to extend the functionality of Outlook beyond our core feature set. Add-ins enable users to access third-party experiences, tasks, and content without needing to leave their inbox. Once installed, Outlook add-ins are available on every platform and device.</span></span>  

<span data-ttu-id="2d7fa-107">以下高级指南将有助于设计和生成引人注目的加载项，可将应用的最佳功能直接引入 Windows、Web、iOS、Mac 和 Android 设备上的 Outlook。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-107">The following high-level guidelines will help you design and build a compelling add-in, which brings the best of your app right into Outlook&mdash;on Windows, Web, iOS, Mac, and Android.</span></span>

## <a name="principles"></a><span data-ttu-id="2d7fa-108">原则</span><span class="sxs-lookup"><span data-stu-id="2d7fa-108">Principles</span></span>

1. <span data-ttu-id="2d7fa-109">**重点关注几个关键任务；并将其做好**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-109">**Focus on a few key tasks; do them well**</span></span>

   <span data-ttu-id="2d7fa-p102">设计一流的加载项易于使用、目标明确并且可为用户带来实际价值。由于加载项将在 Outlook 内部运行，因此这一原则额外重要。Outlook 是生产力应用，人们使用此应用来完成工作。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p102">The best designed add-ins are simple to use, focused, and provide real value to users. Because your add-in will run inside of Outlook, there is additional emphasis placed on this principle. Outlook is a productivity app&mdash;it's where people go to get things done.</span></span>

   <span data-ttu-id="2d7fa-p103">你将成为我们体验的扩展测试人员，请务必确保启用方案就像是在 Outlook 内部进行操作一样自然恰当。认真考虑你的哪些常用用例通过与这些方案挂钩可以从我们的电子邮件和日历体验中获益最大。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p103">You will be an extension of our experience and it is important to make sure the scenarios you enable feel like a natural fit inside of Outlook. Think carefully about which of your common use cases will benefit the most from having hooks to them from within our email and calendaring experiences.</span></span>

   <span data-ttu-id="2d7fa-p104">外接程序不应尝试执行应用所执行的一切操作。重点应放在 Outlook 内容的上下文中使用最频繁的恰当操作。考虑操作调用并明确任务窗格打开时用户应执行什么操作。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p104">An add-in should not attempt to do everything your app does. The focus should be on the most frequently used, and appropriate, actions in the context of Outlook content. Think about your call to action and make it clear what the user should do when your task pane opens.</span></span>

2. <span data-ttu-id="2d7fa-118">**使其尽可能类似于本机模式**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-118">**Make it feel as native as possible**</span></span>

   <span data-ttu-id="2d7fa-p105">应使用正在运行 Outlook 的平台上的本机模式设计外接程序。若要实现这一点，务必尊重并实现各个平台规定的交互和外观准则。Outlook 具有自己的准则，同样也必须考虑这些准则。设计良好的外接程序将恰当地融合体验、平台和 Outlook。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p105">Your add-in should be designed using patterns native to the platform that Outlook is running on. To achieve this, be sure to respect and implement the interaction and visual guidelines set forth by each platform. Outlook has its own guidelines and those are also important to consider. A well-designed add-in will be an appropriate blend of your experience, the platform, and Outlook.</span></span>

   <span data-ttu-id="2d7fa-p106">这就是说，加载项在 iOS 版 Outlook 与在 Android 版 Outlook 上运行时的外观必须不同。我们建议不妨使用 [Framework7](https://framework7.io/) 作为样式设置选项。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p106">This does mean that your add-in will have to visually be different when it runs in Outlook on iOS versus Android. We recommend taking a look at [Framework7](https://framework7.io/) as one option to help you with styling.</span></span>

3. <span data-ttu-id="2d7fa-125">**确保使用体验令人满意，并正确设置详细信息**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-125">**Make it enjoyable to use and get the details right**</span></span>

   <span data-ttu-id="2d7fa-p107">人们喜欢使用实用且外观吸引人的产品。在已仔细考虑每个交互和外观细节的情况下精心构建体验有助于确保加载项成功。完成任务的必要步骤必须清楚并相互关联。理想情况下，操作不应超过一次或两次单击。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p107">People enjoy using products that are both functionally and visually appealing. You can help ensure the success of your add-in by crafting an experience where you've carefully considered every interaction and visual detail. The necessary steps to complete a task must be clear and relevant. Ideally, no action should be further than a click or two away.</span></span> 
   
   <span data-ttu-id="2d7fa-130">尽量不要使用户脱离上下文来完成操作。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-130">Try not to take a user out of context to complete an action.</span></span> <span data-ttu-id="2d7fa-131">用户应可以轻松进入和退出加载项并可轻松返回至用户之前正在执行的操作。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-131">A user should easily be able to get in and out of your add-in and back to whatever she was doing before.</span></span> <span data-ttu-id="2d7fa-132">不应对加载项花费大量的时间，它主要用于增强核心功能。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-132">An add-in is not meant to be a destination to spend a lot of time in&mdash;it is an enhancement to our core functionality.</span></span> <span data-ttu-id="2d7fa-133">如果处理得当，加载项将有助于实现使用户更高效的目标。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-133">If done properly, your add-in will help us deliver on the goal of making people more productive.</span></span>

4. <span data-ttu-id="2d7fa-134">**明智地进行品牌打造**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-134">**Brand wisely**</span></span>

   <span data-ttu-id="2d7fa-135">我们非常重视品牌打造，同时我们知道向用户提供唯一体验至关重要。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-135">We value great branding, and we know it is important to provide users with your unique experience.</span></span> <span data-ttu-id="2d7fa-136">但是我们认为确保加载项成功的最佳方式是生成巧妙整合品牌元素的直观体验，而非显示重复或突兀的品牌元素，它们只会分散用户无阻碍进入系统的注意力。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-136">But we feel the best way to ensure your add-in's success is to build an intuitive experience that subtly incorporates elements of your brand versus displaying persistent or obtrusive brand elements that only distract a user from moving through your system in an unencumbered manner.</span></span> 
    
   <span data-ttu-id="2d7fa-137">有效地整合品牌的良好方式是使用品牌颜色、图标和声音（假定这些与首选的平台模式或辅助功能要求不冲突）。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-137">A good way to incorporate your brand in a meaningful way is through the use of your brand colors, icons, and voice&mdash;assuming these don't conflict with the preferred platform patterns or accessibility requirements.</span></span> <span data-ttu-id="2d7fa-138">努力将重点集中在内容和任务完成方面，而非品牌关注。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-138">Strive to keep the focus on content and task completion, not brand attention.</span></span> 
    
   > [!NOTE]
   >  <span data-ttu-id="2d7fa-139">iOS 或 Android 上的加载项中不应显示广告。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-139">Ads should not be shown within add-ins on iOS or Android.</span></span>

## <a name="design-patterns"></a><span data-ttu-id="2d7fa-140">设计模式</span><span class="sxs-lookup"><span data-stu-id="2d7fa-140">Design patterns</span></span>

> [!NOTE]
> <span data-ttu-id="2d7fa-141">上述准则适用于所有端点/平台，但以下模式和示例特定于 iOS 平台上的移动外接程序。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-141">While the above principles apply to all endpoints/platforms, the following patterns and examples are specific to mobile add-ins on the iOS platform.</span></span>

<span data-ttu-id="2d7fa-p111">我们提供了包含适用于 Outlook Mobile 环境的 iOS 移动模式的[模板](../design/ux-design-pattern-templates.md)，以帮助创建设计良好的外接程序。利用这些特定模式有助于确保外接程序如同在 iOS 平台和 Outlook Mobile 本机自带一般。下面详细介绍了这些模式。虽不全面，但这只是构建库的开始，在我们发现合作伙伴希望纳入其外接程序的其他范例时我们将继续构建此库。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p111">To help you create a well-designed add-in, we have [templates](../design/ux-design-pattern-templates.md) that contain iOS mobile patterns that work within the Outlook Mobile environment. Leveraging these specific patterns will help ensure your add-in feels native to both the iOS platform and Outlook Mobile. These patterns are also detailed below. While not exhaustive, this is the start of a library that we will continue to build upon as we uncover additional paradigms partners wish to include in their add-ins.</span></span>  

### <a name="overview"></a><span data-ttu-id="2d7fa-146">概述</span><span class="sxs-lookup"><span data-stu-id="2d7fa-146">Overview</span></span>

<span data-ttu-id="2d7fa-147">典型的外接程序由下列组件组成。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-147">A typical add-in is made up of the following components.</span></span>

![iOS 上的任务窗格的基本 UX 模式关系图](../images/outlook-mobile-design-overview.png)

![Android 上的任务窗格的基本 UX 模式关系图](../images/outlook-mobile-design-overview-android.jpg)

### <a name="loading"></a><span data-ttu-id="2d7fa-150">加载</span><span class="sxs-lookup"><span data-stu-id="2d7fa-150">Loading</span></span>

<span data-ttu-id="2d7fa-p112">用户点击外接程序后，UX 应尽快显示。如果出现任何延迟，则使用进度栏或活动指示器。时间量可确定时应使用进度栏，时间量不可确定时应使用活动指示器。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p112">When a user taps on your add-in, the UX should display as quickly as possible. If there is any delay, use a progress bar or activity indicator. A progress bar should be used when the amount of time is determinable and an activity indicator should be used when the amount of time is indeterminable.</span></span>

<span data-ttu-id="2d7fa-154">**iOS 上的加载页示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-154">**An example of loading pages on iOS**</span></span>

![iOS 上的进度栏和活动指示器示例](../images/outlook-mobile-design-loading.png)

<span data-ttu-id="2d7fa-156">**Android 上的加载页示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-156">**An example of loading pages on Android**</span></span>

![Android 上的进度栏和活动指示器示例](../images/outlook-mobile-design-loading-android.jpg)


### <a name="sign-insign-up"></a><span data-ttu-id="2d7fa-158">登录/注册</span><span class="sxs-lookup"><span data-stu-id="2d7fa-158">Sign in/Sign up</span></span>

<span data-ttu-id="2d7fa-159">使登录（和注册）流程简单易用。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-159">Make your sign in (and sign up) flow straightforward and simple to use.</span></span>

<span data-ttu-id="2d7fa-160">**iOS 上的登录和注册页示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-160">**An example sign in and sign up page on iOS**</span></span>

![iOS 上的登录和注册页示例](../images/outlook-mobile-design-signin.png)

<span data-ttu-id="2d7fa-162">**Android 上的登录页示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-162">**An example sign in page on Android**</span></span>

![Android 上的登录页示例](../images/outlook-mobile-design-signin-android.png)

### <a name="brand-bar"></a><span data-ttu-id="2d7fa-164">品牌栏</span><span class="sxs-lookup"><span data-stu-id="2d7fa-164">Brand bar</span></span>

<span data-ttu-id="2d7fa-p113">外接程序的第一个屏幕应包含品牌元素。品牌栏用于进行识别，同时也有助于为用户设置上下文。由于导航栏包含公司/品牌的名称，因此没有必要在后续页面上重复品牌栏。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p113">The first screen of your add-in should include your branding element. Designed for recognition, the brand bar also helps set context for the user. Because the navigation bar contains the name of your company/brand, it's unnecessary to repeat the brand bar on subsequent pages.</span></span>

<span data-ttu-id="2d7fa-168">**iOS 上的品牌打造示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-168">**An example of branding on iOS**</span></span>

![iOS 上的品牌栏示例](../images/outlook-mobile-design-branding.png)

<span data-ttu-id="2d7fa-170">**Android 上的品牌打造示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-170">**An example of branding on Android**</span></span>

![Android 上的品牌栏示例](../images/outlook-mobile-design-branding-android.png)

### <a name="margins"></a><span data-ttu-id="2d7fa-172">边距</span><span class="sxs-lookup"><span data-stu-id="2d7fa-172">Margins</span></span>

<span data-ttu-id="2d7fa-173">移动电话边距每侧应设置为 15px（屏幕的 8%），与 Outlook iOS 一致；每侧应设置为 16px 以与 Outlook Android 一致。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-173">Mobile margins should be set to 15px (8% of screen) for each side, to align with Outlook iOS and 16px for each side to align with Outlook Android.</span></span>

![iOS 上的边距示例](../images/outlook-mobile-design-margins.png)

### <a name="typography"></a><span data-ttu-id="2d7fa-175">版式</span><span class="sxs-lookup"><span data-stu-id="2d7fa-175">Typography</span></span>

<span data-ttu-id="2d7fa-176">版式使用与 Outlook iOS 对齐并尽量简单以保证易于浏览。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-176">Typography usage is aligned to Outlook iOS and is kept simple for scannability.</span></span>

<span data-ttu-id="2d7fa-177">**iOS 上的版式**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-177">**Typography on iOS**</span></span>

![适用于 iOS 的版式示例](../images/outlook-mobile-design-typography.png)

<span data-ttu-id="2d7fa-179">**Android 上的版式**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-179">**Typography on Android**</span></span>

![适用于 Android 的版式示例](../images/outlook-mobile-design-typography-android.png)

### <a name="color-palette"></a><span data-ttu-id="2d7fa-181">调色板</span><span class="sxs-lookup"><span data-stu-id="2d7fa-181">Color palette</span></span>

<span data-ttu-id="2d7fa-p114">颜色使用在 Outlook iOS 中比较微妙。我们要求颜色使用本地化到操作和错误状态，以保证一致，仅品牌栏使用唯一的颜色。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p114">Color usage is subtle in Outlook iOS.  To align, we ask that usage of color is localized to actions and error states, with only the brand bar using a unique color.</span></span>

![适用于 iOS 的调色板](../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a><span data-ttu-id="2d7fa-185">单元格</span><span class="sxs-lookup"><span data-stu-id="2d7fa-185">Cells</span></span>

<span data-ttu-id="2d7fa-186">由于导航栏不能用于标记页面，因此使用节标题标记页面。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-186">Since the navigation bar cannot be used to label a page, use section titles to label pages.</span></span>

<span data-ttu-id="2d7fa-187">**iOS 上的单元格示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-187">**Examples of cells on iOS**</span></span>

![适用于 iOS 的单元格类型](../images/outlook-mobile-design-cell-types.png)
* * *
![适用于 iOS 的单元格“待办事项”](../images/outlook-mobile-design-cell-dos.png)
* * *
![适用于 iOS 的单元格“禁止事项”](../images/outlook-mobile-design-cell-donts.png)
* * *
![适用于 iOS 的单元格和输入](../images/outlook-mobile-design-cell-input.png)

<span data-ttu-id="2d7fa-192">**Android 上的单元格示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-192">**Examples of cells on Android**</span></span>

![适用于 Android 的单元格类型](../images/outlook-mobile-design-cell-type-android.png)
* * *
![适用于 Android 的单元格“待办事项”](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![适用于 Android 的单元格“禁止事项”](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![适用于 Android 的单元格和输入第 1 部分](../images/outlook-mobile-design-cell-input-1-android.png)

![适用于 Android 的单元格和输入第 2 部分](../images/outlook-mobile-design-cell-input-2-android.png)

### <a name="actions"></a><span data-ttu-id="2d7fa-198">操作</span><span class="sxs-lookup"><span data-stu-id="2d7fa-198">Actions</span></span>

<span data-ttu-id="2d7fa-199">即使应用要处理大量操作，也要考虑想要外接程序执行的最重要的操作，并重点关注这些操作。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-199">Even if your app handles a multitude of actions, think about the most important ones you want your add-in to perform, and concentrate on those.</span></span>

<span data-ttu-id="2d7fa-200">**iOS 上的操作示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-200">**Examples of actions on iOS**</span></span>

![iOS 中的操作和单元格](../images/outlook-mobile-design-action-cells.png)
* * *
![适用于 iOS 的操作“待办事项”](../images/outlook-mobile-design-action-dos.png)

<span data-ttu-id="2d7fa-203">**Android 上的操作示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-203">**Examples of actions on Android**</span></span>

![Android 中的操作和单元格](../images/outlook-mobile-design-action-cells-android.png)
* * *
![适用于 Android 的操作“待办事项”](../images/outlook-mobile-design-action-dos-android.png)

### <a name="buttons"></a><span data-ttu-id="2d7fa-206">按钮</span><span class="sxs-lookup"><span data-stu-id="2d7fa-206">Buttons</span></span>

<span data-ttu-id="2d7fa-207">存在以下其他 UX 元素时使用按钮（相对于操作，其中操作是屏幕上的最后一个元素）。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-207">Buttons are used when there are other UX elements below (vs. actions, where the action is the last element on the screen).</span></span>

<span data-ttu-id="2d7fa-208">**iOS 上的按钮示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-208">**Examples of buttons on iOS**</span></span>

![适用于 iOS 的按钮示例](../images/outlook-mobile-design-buttons.png)

<span data-ttu-id="2d7fa-210">**Android 上的按钮示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-210">**Examples of buttons on Android**</span></span>

![适用于 Android 的按钮示例](../images/outlook-mobile-design-buttons-android.png)

### <a name="tabs"></a><span data-ttu-id="2d7fa-212">选项卡</span><span class="sxs-lookup"><span data-stu-id="2d7fa-212">Tabs</span></span>

<span data-ttu-id="2d7fa-213">选项卡可帮助组织内容。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-213">Tabs can aid in content organization.</span></span>

<span data-ttu-id="2d7fa-214">**iOS 上的选项卡示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-214">**Examples of tabs on iOS**</span></span>

![适用于 iOS 的选项卡示例](../images/outlook-mobile-design-tabs.png)

<span data-ttu-id="2d7fa-216">**Android 上的选项卡示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-216">**Examples of tabs on Android**</span></span>

![适用于 Android 的选项卡示例](../images/outlook-mobile-design-tabs-android.png)

### <a name="icons"></a><span data-ttu-id="2d7fa-218">图标</span><span class="sxs-lookup"><span data-stu-id="2d7fa-218">Icons</span></span>

<span data-ttu-id="2d7fa-p115">图标应尽可能遵循当前 Outlook iOS 的设计。使用标准大小和颜色。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-p115">Icons should follow the current Outlook iOS design when possible. Use our standard size and color.</span></span>

<span data-ttu-id="2d7fa-221">**iOS 上的图标示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-221">**Examples of icons on iOS**</span></span>

![适用于 iOS 的图标示例](../images/outlook-mobile-design-icons.png)

<span data-ttu-id="2d7fa-223">**Android 上的图标示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-223">**Examples of icons on Android**</span></span>

![适用于 Android 的图标示例](../images/outlook-mobile-design-icons-android.jpg)

## <a name="end-to-end-examples"></a><span data-ttu-id="2d7fa-225">端到端示例</span><span class="sxs-lookup"><span data-stu-id="2d7fa-225">End-to-end examples</span></span>

<span data-ttu-id="2d7fa-226">为了推动 v1 Outlook Mobile 外接程序的启动，我们已与正在生成外接程序的合作伙伴密切合作。作为展示其外接程序在 Outlook Mobile 上的潜力的方式，我们的设计人员使用我们的准则和模式将每个外接程序的端到端流组合在一起。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-226">For our v1 Outlook Mobile Add-ins launch, we worked closely with our partners who were building add-ins. As a way to showcase the potential of their add-ins on Outlook Mobile, our designer put together end-to-end flows for each add-in, leveraging our guidelines and patterns.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2d7fa-227">这些示例旨在强调同时处理加载项的交互和可视化设计的理想方法，可能与加载项发布版本中的准确功能集不匹配。</span><span class="sxs-lookup"><span data-stu-id="2d7fa-227">These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.</span></span> 

### <a name="giphy"></a><span data-ttu-id="2d7fa-228">GIPHY</span><span class="sxs-lookup"><span data-stu-id="2d7fa-228">GIPHY</span></span>

<span data-ttu-id="2d7fa-229">**iOS 上的 GIPHY 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-229">**An example of GIPHY on iOS**</span></span>

![适用于 iOS 上的 GIPHY 加载项的端到端设计](../images/outlook-mobile-design-giphy.png)

<span data-ttu-id="2d7fa-231">**Android 上的 GIPHY 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-231">**An example of GIPHY on Android**</span></span>

![适用于 Android 上的 GIPHY 加载项的端到端设计](../images/outlook-mobile-design-giphy-android.png)

### <a name="nimble"></a><span data-ttu-id="2d7fa-233">Nimble</span><span class="sxs-lookup"><span data-stu-id="2d7fa-233">Nimble</span></span>

<span data-ttu-id="2d7fa-234">**iOS 上的 Nimble 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-234">**An example of Nimble on iOS**</span></span>

![适用于 iOS 上的 Nimble 加载项的端到端设计](../images/outlook-mobile-design-nimble.png)

<span data-ttu-id="2d7fa-236">**Android 上的 Nimble 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-236">**An example of Nimble on Android**</span></span>

![适用于 Android 上的 Nimble 加载项的端到端设计](../images/outlook-mobile-design-nimble-android.png)

### <a name="trello"></a><span data-ttu-id="2d7fa-238">Trello</span><span class="sxs-lookup"><span data-stu-id="2d7fa-238">Trello</span></span>

<span data-ttu-id="2d7fa-239">**iOS 上的 Trello 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-239">**An example of Trello on iOS**</span></span>

![适用于 iOS 上的 Trello 加载项的端到端设计第 1 部分](../images/outlook-mobile-design-trello-1.png)
* * *
![适用于 iOS 上的 Trello 加载项的端到端设计第 2 部分](../images/outlook-mobile-design-trello-2.png)
* * *
![适用于 iOS 上的 Trello 加载项的端到端设计第 3 部分](../images/outlook-mobile-design-trello-3.png)

<span data-ttu-id="2d7fa-243">**Android 上的 Trello 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-243">**An example of Trello on Android**</span></span>

![适用于 Android 上的 Trello 加载项的端到端设计第 1 部分](../images/outlook-mobile-design-trello-1-android.png)
* * *
![适用于 Android 上的 Trello 加载项的端到端设计第 2 部分](../images/outlook-mobile-design-trello-2-android.png)

### <a name="dynamics-crm"></a><span data-ttu-id="2d7fa-246">Dynamics CRM</span><span class="sxs-lookup"><span data-stu-id="2d7fa-246">Dynamics CRM</span></span>

<span data-ttu-id="2d7fa-247">**iOS 上的 Dynamics CRM 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-247">**An example of Dynamics CRM on iOS**</span></span>

![适用于 iOS 上的 Dynamics CRM 加载项的端到端设计](../images/outlook-mobile-design-crm.png)

<span data-ttu-id="2d7fa-249">**Android 上的 Dynamics CRM 示例**</span><span class="sxs-lookup"><span data-stu-id="2d7fa-249">**An example of Dynamics CRM on Android**</span></span>

![适用于 Android 上的 Dynamics CRM 加载项的端到端设计](../images/outlook-mobile-design-crm-android.png)
