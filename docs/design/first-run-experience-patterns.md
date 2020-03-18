---
title: Office 外接程序的首次运行体验模式
description: 了解在 Office 外接程序中设计首次运行体验的最佳实践。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 69634db19c261a8ec475f48e235cd564d91fbfa9
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718669"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="462e8-103">首次运行体验模式</span><span class="sxs-lookup"><span data-stu-id="462e8-103">First-run experience patterns</span></span>

<span data-ttu-id="462e8-104">首次运行体验模式 (FRE) 是对外接程序的用户介绍。</span><span class="sxs-lookup"><span data-stu-id="462e8-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="462e8-105">用户首次打开外接程序时，将会显示 FRE，其中提供有外接程序的函数、功能和/或权益相关的见解。</span><span class="sxs-lookup"><span data-stu-id="462e8-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="462e8-106">此体验有助于塑造用户对外接程序的印象，并提高用户继续使用你的外接程序的可能性。</span><span class="sxs-lookup"><span data-stu-id="462e8-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="462e8-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="462e8-107">Best practices</span></span>


<span data-ttu-id="462e8-108">创建首次运行体验时，请按照以下最佳做法：</span><span class="sxs-lookup"><span data-stu-id="462e8-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="462e8-109">允许事项</span><span class="sxs-lookup"><span data-stu-id="462e8-109">Do</span></span>|<span data-ttu-id="462e8-110">禁止事项</span><span class="sxs-lookup"><span data-stu-id="462e8-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="462e8-111">提供了外接程序中的主要操作的简要介绍。</span><span class="sxs-lookup"><span data-stu-id="462e8-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="462e8-112">不包括与入门无关的信息和标注。</span><span class="sxs-lookup"><span data-stu-id="462e8-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="462e8-113">让用户有机会完成可以积极影响其外接程序使用的操作。</span><span class="sxs-lookup"><span data-stu-id="462e8-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="462e8-114">不要期望用户可以一次性学完全部内容。</span><span class="sxs-lookup"><span data-stu-id="462e8-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="462e8-115">重点关注可提供最大价值的操作。</span><span class="sxs-lookup"><span data-stu-id="462e8-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="462e8-116">创建用户期望完成的富有吸引力的体验。</span><span class="sxs-lookup"><span data-stu-id="462e8-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="462e8-117">不要强制用户单击使用首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="462e8-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="462e8-118">为用户提供可绕过首次运行体验的选项。</span><span class="sxs-lookup"><span data-stu-id="462e8-118">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="462e8-119">向用户显示首次运行体验一次还是定期显示对你的方案来说非常重要。</span><span class="sxs-lookup"><span data-stu-id="462e8-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="462e8-120">例如，如果只是定期使用外接程序，则用户可能不太熟悉外接程序，因此，再次使用首次运行体验可能会有用处。</span><span class="sxs-lookup"><span data-stu-id="462e8-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="462e8-121">根据需要应用以下模式，以创建或提升外接程序的首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="462e8-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="462e8-122">旋转式传送</span><span class="sxs-lookup"><span data-stu-id="462e8-122">Carousel</span></span>


<span data-ttu-id="462e8-123">旋转式传送让用户能够在开始使用外接程序之前浏览一系列功能或信息页面。</span><span class="sxs-lookup"><span data-stu-id="462e8-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="462e8-124">*图 1：允许用户跳过旋转式传送流的开始页面。*
![首次运行 - 旋转式传送 - 桌面任务窗格规范](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="462e8-124">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="462e8-125">*图 2：最小化向用户显示的旋转式传送屏幕的数量，仅提供其所需的信息，以有效传递信息*
![首次运行 - 旋转式传送 - 桌面任务窗格规范](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="462e8-125">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="462e8-126">*图 3：提供明确的行动号召，以退出首次运行体验。*
![首次运行 - 旋转式传送 - 桌面任务窗格规范](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="462e8-126">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="462e8-127">值占位图片</span><span class="sxs-lookup"><span data-stu-id="462e8-127">Value Placemat</span></span>

<span data-ttu-id="462e8-128">值占位通过徽标占位、明确的价值主张、功能亮点或汇总和行动号召传递外接程序的价值主张。</span><span class="sxs-lookup"><span data-stu-id="462e8-128">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="462e8-129">![首次运行 - 值占位图片 - 桌面任务窗格规范](../images/add-in-FRE-value.png)
*具有徽标、明确价值主张、功能汇总和行动号召的值占位图片。*</span><span class="sxs-lookup"><span data-stu-id="462e8-129">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="462e8-130">视频占位图片</span><span class="sxs-lookup"><span data-stu-id="462e8-130">Video Placemat</span></span>

<span data-ttu-id="462e8-131">视频占位图片可以在用户开始使用外接程序之前向其显示视频。</span><span class="sxs-lookup"><span data-stu-id="462e8-131">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="462e8-132">*图 1：首次运行占位图片 - 该屏幕包含视频中的一个静止图像以及一个播放按钮和明确的行动号召按钮。*![视频占位图片 - 桌面任务窗格规范](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="462e8-132">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="462e8-133">*图 2：视频播放器 - 在对话窗口内向用户展示视频。*
![视频占位图片 - 桌面任务窗格规范](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="462e8-133">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
