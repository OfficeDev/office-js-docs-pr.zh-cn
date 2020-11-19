---
title: Office 外接程序的首次运行体验模式
description: 了解在 Office 外接程序中设计首次运行体验的最佳实践。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 00785df2cfd2f41b41917ea720c154e24b72f779
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132065"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="cc9e0-103">首次运行体验模式</span><span class="sxs-lookup"><span data-stu-id="cc9e0-103">First-run experience patterns</span></span>

<span data-ttu-id="cc9e0-104">首次运行体验模式 (FRE) 是对外接程序的用户介绍。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="cc9e0-105">用户首次打开外接程序时，将会显示 FRE，其中提供有外接程序的函数、功能和/或权益相关的见解。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="cc9e0-106">此体验有助于塑造用户对外接程序的印象，并提高用户继续使用你的外接程序的可能性。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="cc9e0-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="cc9e0-107">Best practices</span></span>

<span data-ttu-id="cc9e0-108">创建首次运行体验时，请按照以下最佳做法：</span><span class="sxs-lookup"><span data-stu-id="cc9e0-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="cc9e0-109">允许事项</span><span class="sxs-lookup"><span data-stu-id="cc9e0-109">Do</span></span>|<span data-ttu-id="cc9e0-110">禁止事项</span><span class="sxs-lookup"><span data-stu-id="cc9e0-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="cc9e0-111">提供了外接程序中的主要操作的简要介绍。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="cc9e0-112">不包括与入门无关的信息和标注。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="cc9e0-113">让用户有机会完成可以积极影响其外接程序使用的操作。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="cc9e0-114">不要期望用户可以一次性学完全部内容。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="cc9e0-115">重点关注可提供最大价值的操作。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="cc9e0-116">创建用户期望完成的富有吸引力的体验。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="cc9e0-117">不要强制用户单击使用首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="cc9e0-118">为用户提供可绕过首次运行体验的选项。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-118">Give users an option to bypass the first-run experience.</span></span> |

<span data-ttu-id="cc9e0-119">向用户显示首次运行体验一次还是定期显示对你的方案来说非常重要。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="cc9e0-120">例如，如果只是定期使用外接程序，则用户可能不太熟悉外接程序，因此，再次使用首次运行体验可能会有用处。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>

<span data-ttu-id="cc9e0-121">根据需要应用以下模式，以创建或提升外接程序的首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>

## <a name="carousel"></a><span data-ttu-id="cc9e0-122">旋转式传送</span><span class="sxs-lookup"><span data-stu-id="cc9e0-122">Carousel</span></span>

<span data-ttu-id="cc9e0-123">旋转式传送让用户能够在开始使用外接程序之前浏览一系列功能或信息页面。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="cc9e0-124">*图1。允许用户提前或跳过轮播流的起始页*</span><span class="sxs-lookup"><span data-stu-id="cc9e0-124">*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*</span></span>

![图示在 Office 桌面应用程序任务窗格的首次运行体验中显示轮播的步骤1。](../images/add-in-FRE-step-1.png)

<span data-ttu-id="cc9e0-127">*图2。将轮播屏幕的数量最小化，以有效地传递邮件所需的屏幕数量*</span><span class="sxs-lookup"><span data-stu-id="cc9e0-127">*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*</span></span>

![图示在 Office 桌面应用程序任务窗格的首次运行体验中显示轮播的步骤2。](../images/add-in-FRE-step-2.png)

<span data-ttu-id="cc9e0-130">*图3。提供对操作的明确调用，以退出首次运行体验*</span><span class="sxs-lookup"><span data-stu-id="cc9e0-130">*Figure 3. Provide a clear call to action to exit the first-run-experience*</span></span>

![图示在 Office 桌面应用程序任务窗格的首次运行体验中显示轮播的步骤3。](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a><span data-ttu-id="cc9e0-133">值占位图片</span><span class="sxs-lookup"><span data-stu-id="cc9e0-133">Value Placemat</span></span>

<span data-ttu-id="cc9e0-134">值占位通过徽标占位、明确的价值主张、功能亮点或汇总和行动号召传递外接程序的价值主张。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-134">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>

<span data-ttu-id="cc9e0-135">*图4。值占位图片，带有徽标、清除价值主张、功能摘要和行动要求*</span><span class="sxs-lookup"><span data-stu-id="cc9e0-135">*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*</span></span>

![图中显示了在 Office 桌面应用程序任务窗格的首次运行体验中占位图片的值。](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a><span data-ttu-id="cc9e0-138">视频占位图片</span><span class="sxs-lookup"><span data-stu-id="cc9e0-138">Video Placemat</span></span>

<span data-ttu-id="cc9e0-139">视频占位图片可以在用户开始使用外接程序之前向其显示视频。</span><span class="sxs-lookup"><span data-stu-id="cc9e0-139">The video placemat shows users a video before they start using your add-in.</span></span>

<span data-ttu-id="cc9e0-140">*图5。第一次运行视频占位图片-屏幕包含视频中的静止图像和 "播放" 按钮，并清除 "操作-操作" 按钮*</span><span class="sxs-lookup"><span data-stu-id="cc9e0-140">*Figure 5. First run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*</span></span>

![在 Office 桌面应用程序任务窗格的首次运行体验中显示视频占位图片的图示](../images/add-in-FRE-video.png)

<span data-ttu-id="cc9e0-142">*图6。视频播放器-在对话框窗口中显示有视频的用户*</span><span class="sxs-lookup"><span data-stu-id="cc9e0-142">*Figure 6. Video player - Users presented with a video within a dialog window*</span></span>

![在背景中显示带有 Office 桌面应用程序和外接程序任务窗格的对话框窗口中的视频的插图](../images/add-in-FRE-video-dialog.png)
