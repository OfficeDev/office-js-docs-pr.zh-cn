---
title: 在 Office 加载项中使用动作
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d3be2454b36fe1003c0697f0bca3c29d743e5330
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871533"
---
# <a name="using-motion-in-office-add-ins"></a><span data-ttu-id="844b5-102">在 Office 加载项中使用动作</span><span class="sxs-lookup"><span data-stu-id="844b5-102">Using motion in Office Add-ins</span></span>

<span data-ttu-id="844b5-p101">设计 Office 加载项时，可以使用动作来提升用户体验。 UI 元素、控件和组件通常都有需要使用转换、动作或动画的交互行为。 UI 界面元素之间运动的共同特征定义设计语言的动画方面。</span><span class="sxs-lookup"><span data-stu-id="844b5-p101">When you design an Office Add-in, you can use motion to enhance the user experience. UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation. Common characteristics of motion across UI elements define the animation aspects of a design language.</span></span> 

<span data-ttu-id="844b5-p102">Office 的重点是工作效率，因此 Office 动画语言支持帮助客户完成工作的目标。 力求在高性能响应、可靠编排和细节带来的喜悦之间实现平衡。 Office 中嵌入的加载项不超出现有动画语言范围。 鉴于此，在应用动作时，请务必注意遵循以下几项指南。</span><span class="sxs-lookup"><span data-stu-id="844b5-p102">Because Office is focused on productivity, the Office animation language supports the goal of helping customers get things done. It strikes a balance between performant response, reliable choreography, and detailed delight. Add-ins embedded in Office sit within this existing animation language. Given this context, it is important to consider the following guidelines when applying motion.</span></span> 


## <a name="create-motion-with-a-purpose"></a><span data-ttu-id="844b5-110">创建有明确用途的动作</span><span class="sxs-lookup"><span data-stu-id="844b5-110">Create motion with a purpose</span></span>

<span data-ttu-id="844b5-p103">动作应具有明确用途，让用户感受到更有价值。 选择动画时，请考虑内容的基调和用途。 关键消息的处理方式不同于探索导航。</span><span class="sxs-lookup"><span data-stu-id="844b5-p103">Motion should have a purpose that communicates additional value to the user. Consider the tone and purpose of your content when choosing animations. Handle critical messages differently than exploratory navigations.</span></span>

<span data-ttu-id="844b5-p104">加载项中使用的标准元素可以纳入动作，不仅有助于用户集中注意力、呈现元素之间的关系，还有助于验证用户操作。 将元素编排为加强层次结构和心理模型。</span><span class="sxs-lookup"><span data-stu-id="844b5-p104">Standard elements used in an add-in can incorporate motion to help focus the user, show how elements relate to each other, and validate user actions. Choreograph elements to reinforce hierarchy and mental models.</span></span>

### <a name="best-practices"></a><span data-ttu-id="844b5-116">最佳做法</span><span class="sxs-lookup"><span data-stu-id="844b5-116">Best practices</span></span>

|<span data-ttu-id="844b5-117">允许事项</span><span class="sxs-lookup"><span data-stu-id="844b5-117">Do</span></span>|<span data-ttu-id="844b5-118">禁止事项</span><span class="sxs-lookup"><span data-stu-id="844b5-118">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="844b5-p105">确定加载项中应包含动作的关键元素。 加载项中的常见动画元素包括面板、叠加层、模式、工具提示、菜单和教导标注。</span><span class="sxs-lookup"><span data-stu-id="844b5-p105">Identify key elements in the add-in that should have motion. Commonly animated elements in an add-in are panels, overlays, modals, tool tips, menus, and teaching call outs.</span></span>| <span data-ttu-id="844b5-p106">不得为每个元素都添加动画效果，否则用户会感到不知所措。 避免应用多个动作，以试图让用户一次关注多个元素。</span><span class="sxs-lookup"><span data-stu-id="844b5-p106">Don't overwhelm the user by animating every element. Avoid applying multiple motions that attempt to lead or focus the user on many elements at once.</span></span> |
|<span data-ttu-id="844b5-p107">应使用行为符合预期的简单精细动作。请考虑触发元素的起源。使用动作可以在操作和生成的 UI 之间创建关联。</span><span class="sxs-lookup"><span data-stu-id="844b5-p107">Use simple, subtle motion that behaves in expected ways. Consider the origin of your triggering element. Use motion to create a link between the action and the resulting UI.</span></span> | <span data-ttu-id="844b5-p108">不得创建有等待时间的动作。 加载项中的动作不得妨碍任务完成。</span><span class="sxs-lookup"><span data-stu-id="844b5-p108">Don't create wait time for a motion. Motion in add-ins should not hinder task completion.</span></span>|

![左 gif 显示打开后移动元素最少的面板，右 gif 显示打开后包含许多移动元素的面板](../images/add-in-motion-purpose.gif)

## <a name="use-expected-motions"></a><span data-ttu-id="844b5-129">使用符合预期的动作</span><span class="sxs-lookup"><span data-stu-id="844b5-129">Use expected motions</span></span>

<span data-ttu-id="844b5-130">建议使用 [Office UI Fabric](https://developer.microsoft.com/fabric) 直观连接到 Office 平台，还建议使用 [Fabric 动画](https://developer.microsoft.com/fabric#/styles/animations)创建与 Fabric 动作语言一致的动作。</span><span class="sxs-lookup"><span data-stu-id="844b5-130">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric) to create a visual connection with the Office platform, and we also encourage the use of [Fabric Animations](https://developer.microsoft.com/fabric#/styles/animations) to create motions that align with the Fabric motion language.</span></span> 

<span data-ttu-id="844b5-p109">它可用于在 Office 中无缝集成。它有助于创建更侧重用户感受（而不是外观）的体验。动画 CSS 类提供方向、进入/退出和持续时间（强化 Office 心理模型），并为客户提供了解如何与加载项交互的机会。</span><span class="sxs-lookup"><span data-stu-id="844b5-p109">Use it to fit seamlessly in Office. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in.</span></span>

### <a name="best-practices"></a><span data-ttu-id="844b5-134">最佳做法</span><span class="sxs-lookup"><span data-stu-id="844b5-134">Best practices</span></span>

|<span data-ttu-id="844b5-135">允许事项</span><span class="sxs-lookup"><span data-stu-id="844b5-135">Do</span></span>|<span data-ttu-id="844b5-136">禁止事项</span><span class="sxs-lookup"><span data-stu-id="844b5-136">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="844b5-137">应使用与 Fabric 行为一致的动作。</span><span class="sxs-lookup"><span data-stu-id="844b5-137">Use motion that aligns with behaviors in Fabric.</span></span>| <span data-ttu-id="844b5-138">不得创建干扰 Office 中常见动作模式或与其冲突的动作。</span><span class="sxs-lookup"><span data-stu-id="844b5-138">Don't create motions that interfere or conflict with common motion patterns in Office.</span></span>
|<span data-ttu-id="844b5-139">确保在 like 元素中有一致的运动应用程序。</span><span class="sxs-lookup"><span data-stu-id="844b5-139">Ensure that there is a consistent application of motion across like elements.</span></span>| <span data-ttu-id="844b5-140">不得使用不同动作为同一组件或对象添加动画效果。</span><span class="sxs-lookup"><span data-stu-id="844b5-140">Don't use different motions to animate the same component or object.</span></span>|
|<span data-ttu-id="844b5-p110">应确保动画方向的使用一致。 例如，从右侧打开的面板应向右侧关闭。</span><span class="sxs-lookup"><span data-stu-id="844b5-p110">Create consistency with use of direction in animation. For example, a panel that opens from the right should close to the right.</span></span>|<span data-ttu-id="844b5-143">不得使用多个方向为元素添加动画效果。</span><span class="sxs-lookup"><span data-stu-id="844b5-143">Don't animate an element using multiple directions.</span></span>

![左 gif 显示模式以预期方式打开，右 gif 显示模式以异常方式打开](../images/add-in-motion-expected.gif)

## <a name="avoid-out-of-character-motion-for-an-element"></a><span data-ttu-id="844b5-145">避免对元素使用不符合预期的动作</span><span class="sxs-lookup"><span data-stu-id="844b5-145">Avoid out of character motion for an element</span></span>

<span data-ttu-id="844b5-p111">实现动作时，请考虑 HTML 画布（任务窗格、对话框或内容加载项）的尺寸。 避免在受限空间中重载。 一个或多个移动元素应与 Office 协调一致。 加载项动作应可靠、流畅且高性能。 动作旨在提供指示和指导，而不是降低工作效率。</span><span class="sxs-lookup"><span data-stu-id="844b5-p111">Consider the size of the HTML canvas (task pane, dialog box, or content add-in) when implementing motion. Avoid overloading in constrained spaces. Moving element(s) should be in tune with Office. The character of add-in motion should be performant, reliable, and fluid. Instead of impeding productivity, aim to inform and direct.</span></span>

### <a name="best-practices"></a><span data-ttu-id="844b5-151">最佳做法</span><span class="sxs-lookup"><span data-stu-id="844b5-151">Best practices</span></span>

|<span data-ttu-id="844b5-152">允许事项</span><span class="sxs-lookup"><span data-stu-id="844b5-152">Do</span></span>|<span data-ttu-id="844b5-153">禁止事项</span><span class="sxs-lookup"><span data-stu-id="844b5-153">Don't</span></span>|
|:-----|:-----|
| <span data-ttu-id="844b5-154">应使用[建议的动作持续时间](https://developer.microsoft.com/fabric#/styles/animations)。</span><span class="sxs-lookup"><span data-stu-id="844b5-154">Use [recommended motion durations](https://developer.microsoft.com/fabric#/styles/animations).</span></span> | <span data-ttu-id="844b5-p112">不得使用夸张的动画。 避免打造会分散客户注意力的花哨体验。</span><span class="sxs-lookup"><span data-stu-id="844b5-p112">Don't use exaggerated animations. Avoid creating experiences that embellish and distract your customers.</span></span>
| <span data-ttu-id="844b5-157">请遵循[建议的缓和曲线](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion)。</span><span class="sxs-lookup"><span data-stu-id="844b5-157">Follow [recommended easing curves](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span></span>  |<span data-ttu-id="844b5-p113">不得用不连贯的方式移动元素。 避免占先、退回、橡皮筋或其他模拟自然世界物理学的效果。</span><span class="sxs-lookup"><span data-stu-id="844b5-p113">Don't move elements in a jerky or disjointed manner. Avoid anticipations, bounces, rubberband, or other effects that emulate natural world physics.</span></span>|

![左 gif 显示使用缓和淡化效果加载磁贴，右 gif 显示使用退回效果加载磁贴](../images/add-in-motion-character.gif)

## <a name="see-also"></a><span data-ttu-id="844b5-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="844b5-161">See also</span></span>

* [<span data-ttu-id="844b5-162">Fabric 动画指南</span><span class="sxs-lookup"><span data-stu-id="844b5-162">Fabric animation guidelines</span></span>](https://developer.microsoft.com/fabric#/styles/animations)
* [<span data-ttu-id="844b5-163">通用 Windows 平台应用动作</span><span class="sxs-lookup"><span data-stu-id="844b5-163">Motion for Universal Windows Platform apps</span></span>](/windows/uwp/design/motion)
