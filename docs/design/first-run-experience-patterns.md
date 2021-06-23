---
title: Office 外接程序的首次运行体验模式
description: 了解在加载项中设计首次运行体验Office最佳做法。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: d020a281aca10805ba8fd1176403f3788f6d716c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076341"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="dfc54-103">首次运行体验模式</span><span class="sxs-lookup"><span data-stu-id="dfc54-103">First-run experience patterns</span></span>

<span data-ttu-id="dfc54-104">首次运行体验模式 (FRE) 是对外接程序的用户介绍。</span><span class="sxs-lookup"><span data-stu-id="dfc54-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="dfc54-105">用户首次打开外接程序时，将会显示 FRE，其中提供有外接程序的函数、功能和/或权益相关的见解。</span><span class="sxs-lookup"><span data-stu-id="dfc54-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="dfc54-106">此体验有助于塑造用户对外接程序的印象，并提高用户继续使用你的外接程序的可能性。</span><span class="sxs-lookup"><span data-stu-id="dfc54-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="dfc54-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="dfc54-107">Best practices</span></span>

<span data-ttu-id="dfc54-108">创建首次运行体验时，请按照以下最佳做法：</span><span class="sxs-lookup"><span data-stu-id="dfc54-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="dfc54-109">允许事项</span><span class="sxs-lookup"><span data-stu-id="dfc54-109">Do</span></span>|<span data-ttu-id="dfc54-110">禁止事项</span><span class="sxs-lookup"><span data-stu-id="dfc54-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="dfc54-111">提供了外接程序中的主要操作的简要介绍。</span><span class="sxs-lookup"><span data-stu-id="dfc54-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="dfc54-112">不包括与入门无关的信息和标注。</span><span class="sxs-lookup"><span data-stu-id="dfc54-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="dfc54-113">让用户有机会完成可以积极影响其外接程序使用的操作。</span><span class="sxs-lookup"><span data-stu-id="dfc54-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="dfc54-114">不要期望用户可以一次性学完全部内容。</span><span class="sxs-lookup"><span data-stu-id="dfc54-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="dfc54-115">重点关注可提供最大价值的操作。</span><span class="sxs-lookup"><span data-stu-id="dfc54-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="dfc54-116">创建用户期望完成的富有吸引力的体验。</span><span class="sxs-lookup"><span data-stu-id="dfc54-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="dfc54-117">不要强制用户单击使用首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="dfc54-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="dfc54-118">为用户提供可绕过首次运行体验的选项。</span><span class="sxs-lookup"><span data-stu-id="dfc54-118">Give users an option to bypass the first-run experience.</span></span> |

<span data-ttu-id="dfc54-119">向用户显示首次运行体验一次还是定期显示对你的方案来说非常重要。</span><span class="sxs-lookup"><span data-stu-id="dfc54-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="dfc54-120">例如，如果只是定期使用外接程序，则用户可能不太熟悉外接程序，因此，再次使用首次运行体验可能会有用处。</span><span class="sxs-lookup"><span data-stu-id="dfc54-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>

<span data-ttu-id="dfc54-121">根据需要应用以下模式，以创建或提升外接程序的首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="dfc54-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>

## <a name="carousel"></a><span data-ttu-id="dfc54-122">旋转式传送</span><span class="sxs-lookup"><span data-stu-id="dfc54-122">Carousel</span></span>

<span data-ttu-id="dfc54-123">旋转式传送让用户能够在开始使用外接程序之前浏览一系列功能或信息页面。</span><span class="sxs-lookup"><span data-stu-id="dfc54-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="dfc54-124">*图 1.允许用户前进或跳过木马流的开始页面*</span><span class="sxs-lookup"><span data-stu-id="dfc54-124">*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*</span></span>

![插图显示桌面应用程序任务窗格首次运行体验中的Office 1。](../images/add-in-FRE-step-1.png)

<span data-ttu-id="dfc54-127">*图 2.将可播放的屏幕数最小化为仅有效传达消息所需的内容*</span><span class="sxs-lookup"><span data-stu-id="dfc54-127">*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*</span></span>

![插图显示桌面应用程序任务窗格首次运行体验中的轮Office的第 2 步。](../images/add-in-FRE-step-2.png)

<span data-ttu-id="dfc54-130">*图 3.提供明确的行动号召以退出首次运行体验*</span><span class="sxs-lookup"><span data-stu-id="dfc54-130">*Figure 3. Provide a clear call to action to exit the first-run-experience*</span></span>

![插图显示桌面应用程序任务窗格首次运行体验中的Office 3。](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a><span data-ttu-id="dfc54-133">值占位图片</span><span class="sxs-lookup"><span data-stu-id="dfc54-133">Value Placemat</span></span>

<span data-ttu-id="dfc54-134">值占位通过徽标占位、明确的价值主张、功能亮点或汇总和行动号召传递外接程序的价值主张。</span><span class="sxs-lookup"><span data-stu-id="dfc54-134">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>

<span data-ttu-id="dfc54-135">*图 4.具有徽标、清晰价值主张、功能摘要和行动号召的值位置图片*</span><span class="sxs-lookup"><span data-stu-id="dfc54-135">*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*</span></span>

![显示桌面应用程序任务窗格首次运行体验中的Office插图。](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a><span data-ttu-id="dfc54-138">视频占位图片</span><span class="sxs-lookup"><span data-stu-id="dfc54-138">Video Placemat</span></span>

<span data-ttu-id="dfc54-139">视频占位图片可以在用户开始使用外接程序之前向其显示视频。</span><span class="sxs-lookup"><span data-stu-id="dfc54-139">The video placemat shows users a video before they start using your add-in.</span></span>

<span data-ttu-id="dfc54-140">*图 5.首次运行视频放置图片 - 屏幕包含视频中的静止图像，其中包含播放按钮和清除"调用操作"按钮*</span><span class="sxs-lookup"><span data-stu-id="dfc54-140">*Figure 5. First run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*</span></span>

![插图显示桌面应用程序任务窗格首次运行体验中的Office图片。](../images/add-in-FRE-video.png)

<span data-ttu-id="dfc54-142">*图 6.视频播放器 - 在对话框窗口中向用户呈现视频*</span><span class="sxs-lookup"><span data-stu-id="dfc54-142">*Figure 6. Video player - Users presented with a video within a dialog window*</span></span>

![插图显示对话框窗口中的一个视频Office桌面应用程序和外接程序任务窗格在后台显示。](../images/add-in-FRE-video-dialog.png)
