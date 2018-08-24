# <a name="using-motion-in-office-add-ins"></a><span data-ttu-id="8cb56-101">在 Office 加载项中使用动作</span><span class="sxs-lookup"><span data-stu-id="8cb56-101">Using motion in Office add-ins</span></span>

<span data-ttu-id="8cb56-102">设计 Office 加载项时，可以使用动作来提升用户体验。</span><span class="sxs-lookup"><span data-stu-id="8cb56-102">When you design an Office Add-in, you can use motion to enhance the user experience.</span></span> <span data-ttu-id="8cb56-103">UI 元素、控件和组件通常都有需要使用转换、动作或动画的交互行为。</span><span class="sxs-lookup"><span data-stu-id="8cb56-103">UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation.</span></span> <span data-ttu-id="8cb56-104">UI 界面元素之间运动的共同特征定义设计语言的动画方面。</span><span class="sxs-lookup"><span data-stu-id="8cb56-104">Common characteristics of motion across UI elements define the animation aspects of a design language.</span></span> 

<span data-ttu-id="8cb56-105">Office 的重点是工作效率，因此 Office 动画语言支持帮助客户完成工作的目标。</span><span class="sxs-lookup"><span data-stu-id="8cb56-105">Because Office is focused on productivity, the Office animation language supports the goal of helping customers get things done.</span></span> <span data-ttu-id="8cb56-106">力求在高性能响应、可靠编排和细节带来的喜悦之间实现平衡。</span><span class="sxs-lookup"><span data-stu-id="8cb56-106">It strikes a balance between performant response, reliable choreography, and detailed delight.</span></span> <span data-ttu-id="8cb56-107">Office 中嵌入的加载项不超出现有动画语言范围。</span><span class="sxs-lookup"><span data-stu-id="8cb56-107">Add-ins embedded in Office sit within this existing animation language.</span></span> <span data-ttu-id="8cb56-108">鉴于此，在应用动作时，请务必注意遵循以下几项指南。</span><span class="sxs-lookup"><span data-stu-id="8cb56-108">Given this context, it is important to consider the following guidelines when applying motion.</span></span> 


## <a name="create-motion-with-a-purpose"></a><span data-ttu-id="8cb56-109">创建有明确用途的动作</span><span class="sxs-lookup"><span data-stu-id="8cb56-109">Create motion with a purpose</span></span>

<span data-ttu-id="8cb56-110">动作应具有明确用途，让用户感受到更有价值。</span><span class="sxs-lookup"><span data-stu-id="8cb56-110">Motion should have a purpose that communicates additional value to the user.</span></span> <span data-ttu-id="8cb56-111">选择动画时，请考虑内容的基调和用途。</span><span class="sxs-lookup"><span data-stu-id="8cb56-111">Consider the tone and purpose of your content when choosing animations.</span></span> <span data-ttu-id="8cb56-112">关键消息的处理方式不同于探索导航。</span><span class="sxs-lookup"><span data-stu-id="8cb56-112">Handle critical messages differently than exploratory navigations.</span></span>

<span data-ttu-id="8cb56-113">加载项中使用的标准元素可以纳入动作，不仅有助于用户集中注意力、呈现元素之间的关系，还有助于验证用户操作。</span><span class="sxs-lookup"><span data-stu-id="8cb56-113">Standard elements used in an add-in can incorporate motion to help focus the user, show how elements relate to each other, and validate user actions.</span></span> <span data-ttu-id="8cb56-114">将元素编排为加强层次结构和心理模型。</span><span class="sxs-lookup"><span data-stu-id="8cb56-114">Choreograph elements to reinforce hierarchy and mental models.</span></span>



### <a name="best-practices"></a><span data-ttu-id="8cb56-115">最佳做法</span><span class="sxs-lookup"><span data-stu-id="8cb56-115">Best practices</span></span>

|<span data-ttu-id="8cb56-116">允许事项</span><span class="sxs-lookup"><span data-stu-id="8cb56-116">Do</span></span>|<span data-ttu-id="8cb56-117">禁止事项</span><span class="sxs-lookup"><span data-stu-id="8cb56-117">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="8cb56-118">确定加载项中应包含动作的关键元素。</span><span class="sxs-lookup"><span data-stu-id="8cb56-118">Identify key elements in the add-in that should have motion.</span></span> <span data-ttu-id="8cb56-119">加载项中的常见动画元素包括面板、叠加层、模式、工具提示、菜单和教导标注。</span><span class="sxs-lookup"><span data-stu-id="8cb56-119">Commonly animated elements in an add-in are panels, overlays, modals, tool tips, menus, and teaching call outs.</span></span>| <span data-ttu-id="8cb56-120">不得为每个元素都添加动画效果，否则用户会感到不知所措。</span><span class="sxs-lookup"><span data-stu-id="8cb56-120">Don't overwhelm the user by animating every element.</span></span> <span data-ttu-id="8cb56-121">避免应用多个动作，以试图让用户一次关注多个元素。</span><span class="sxs-lookup"><span data-stu-id="8cb56-121">Avoid applying multiple motions that attempt to lead or focus the user on many elements at once.</span></span> |
|<span data-ttu-id="8cb56-p107">应使用行为符合预期的简单精细动作。请考虑触发元素的起源。使用动作可以在操作和生成的 UI 之间创建关联。</span><span class="sxs-lookup"><span data-stu-id="8cb56-p107">Use simple, subtle motion that behaves in expected ways. Consider the origin of your triggering element. Use motion to create a link between the action and the resulting UI.</span></span> | <span data-ttu-id="8cb56-125">不得创建有等待时间的动作。</span><span class="sxs-lookup"><span data-stu-id="8cb56-125">Don't create wait time for a motion.</span></span> <span data-ttu-id="8cb56-126">加载项中的动作不得妨碍任务完成。</span><span class="sxs-lookup"><span data-stu-id="8cb56-126">Motion in add-ins should not hinder task completion.</span></span>|

![左 gif 显示打开后移动元素最少的面板，右 gif 显示打开后包含许多移动元素的面板](../images/add-in-motion-purpose.gif)



## <a name="use-expected-motions"></a><span data-ttu-id="8cb56-128">使用符合预期的动作</span><span class="sxs-lookup"><span data-stu-id="8cb56-128">Use expected motions</span></span>
<span data-ttu-id="8cb56-129">建议使用 [Office UI Fabric](https://developer.microsoft.com/fabric) 直观连接到 Office 平台，还建议使用 [Fabric 动画](https://developer.microsoft.com/fabric#/styles/animations)创建与 Fabric 动作语言一致的动作。</span><span class="sxs-lookup"><span data-stu-id="8cb56-129">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric) to create a visual connection with the Office platform, and we also encourage the use of [Fabric Animations](https://developer.microsoft.com/fabric#/styles/animations) to create motions that align with the Fabric motion language.</span></span> 

<span data-ttu-id="8cb56-p109">它可用于在 Office 中无缝集成。它有助于创建更侧重用户感受（而不是外观）的体验。动画 CSS 类提供方向、进入/退出和持续时间（强化 Office 心理模型），并为客户提供了解如何与加载项交互的机会。</span><span class="sxs-lookup"><span data-stu-id="8cb56-p109">Use it to fit seamlessly in Office. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in.</span></span>

### <a name="best-practices"></a><span data-ttu-id="8cb56-133">最佳做法</span><span class="sxs-lookup"><span data-stu-id="8cb56-133">Best practices</span></span>


|<span data-ttu-id="8cb56-134">允许事项</span><span class="sxs-lookup"><span data-stu-id="8cb56-134">Do</span></span>|<span data-ttu-id="8cb56-135">禁止事项</span><span class="sxs-lookup"><span data-stu-id="8cb56-135">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="8cb56-136">应使用与 Fabric 行为一致的动作。</span><span class="sxs-lookup"><span data-stu-id="8cb56-136">Use motion that aligns with behaviors in Fabric.</span></span>| <span data-ttu-id="8cb56-137">不得创建干扰 Office 中常见动作模式或与其冲突的动作。</span><span class="sxs-lookup"><span data-stu-id="8cb56-137">Don't create motions that interfere or conflict with common motion patterns in Office.</span></span> 
|<span data-ttu-id="8cb56-138">应确保动作的跨元素应用一致。</span><span class="sxs-lookup"><span data-stu-id="8cb56-138">Ensure that there is a consistent application of motion acoss like elements.</span></span>| <span data-ttu-id="8cb56-139">不得使用不同动作为同一组件或对象添加动画效果。</span><span class="sxs-lookup"><span data-stu-id="8cb56-139">Don't use different motions to animate the same component or object.</span></span>|
|<span data-ttu-id="8cb56-140">应确保动画方向的使用一致。</span><span class="sxs-lookup"><span data-stu-id="8cb56-140">Create consistency with use of direction in animation.</span></span> <span data-ttu-id="8cb56-141">例如，从右侧打开的面板应向右侧关闭。</span><span class="sxs-lookup"><span data-stu-id="8cb56-141">For example, a panel that opens from the right should close to the right.</span></span>|<span data-ttu-id="8cb56-142">不得使用多个方向为元素添加动画效果。</span><span class="sxs-lookup"><span data-stu-id="8cb56-142">Don't animate an element using multiple directions.</span></span>

![左 gif 显示模式以预期方式打开，右 gif 显示模式以异常方式打开](../images/add-in-motion-expected.gif)

## <a name="avoid-out-of-character-motion-for-an-element"></a><span data-ttu-id="8cb56-144">避免对元素使用不符合预期的动作</span><span class="sxs-lookup"><span data-stu-id="8cb56-144">Avoid out of character motion for an element</span></span>

<span data-ttu-id="8cb56-145">实现动作时，请考虑 HTML 画布（任务窗格、对话框或内容加载项）的尺寸。</span><span class="sxs-lookup"><span data-stu-id="8cb56-145">Consider the size of the HTML canvas (task pane, dialog box, or content add-in) when implementing motion.</span></span> <span data-ttu-id="8cb56-146">避免在受限空间中重载。</span><span class="sxs-lookup"><span data-stu-id="8cb56-146">Avoid overloading in constrained spaces.</span></span> <span data-ttu-id="8cb56-147">一个或多个移动元素应与 Office 协调一致。</span><span class="sxs-lookup"><span data-stu-id="8cb56-147">Moving element(s) should be in tune with Office.</span></span> <span data-ttu-id="8cb56-148">加载项动作应可靠、流畅且高性能。</span><span class="sxs-lookup"><span data-stu-id="8cb56-148">The character of add-in motion should be performant, reliable, and fluid.</span></span> <span data-ttu-id="8cb56-149">动作旨在提供指示和指导，而不是降低工作效率。</span><span class="sxs-lookup"><span data-stu-id="8cb56-149">Instead of impeding productivity, aim to inform and direct.</span></span>

### <a name="best-practices"></a><span data-ttu-id="8cb56-150">最佳做法</span><span class="sxs-lookup"><span data-stu-id="8cb56-150">Best practices</span></span>

|<span data-ttu-id="8cb56-151">允许事项</span><span class="sxs-lookup"><span data-stu-id="8cb56-151">Do</span></span>|<span data-ttu-id="8cb56-152">禁止事项</span><span class="sxs-lookup"><span data-stu-id="8cb56-152">Don't</span></span>|
|:-----|:-----|
| <span data-ttu-id="8cb56-153">应使用[建议的动作持续时间](https://developer.microsoft.com/fabric#/styles/animations)。</span><span class="sxs-lookup"><span data-stu-id="8cb56-153">Use [recommended motion durations](https://developer.microsoft.com/fabric#/styles/animations).</span></span> | <span data-ttu-id="8cb56-154">不得使用夸张的动画。</span><span class="sxs-lookup"><span data-stu-id="8cb56-154">Don't use exaggerated animations.</span></span> <span data-ttu-id="8cb56-155">避免打造会分散客户注意力的花哨体验。</span><span class="sxs-lookup"><span data-stu-id="8cb56-155">Avoid creating experiences that embellish and distract your customers.</span></span>
| <span data-ttu-id="8cb56-156">遵循[建议的缓和曲线](https://docs.microsoft.com/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion)。</span><span class="sxs-lookup"><span data-stu-id="8cb56-156">Follow recommended easing curves in the [Microsoft Motion Guide](https://docs.microsoft.com/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span></span>  |<span data-ttu-id="8cb56-157">不得用不连贯的方式移动元素。</span><span class="sxs-lookup"><span data-stu-id="8cb56-157">Don't move elements in a jerky or disjointed manner.</span></span> <span data-ttu-id="8cb56-158">避免占先、退回、橡皮筋或其他模拟自然世界物理学的效果。</span><span class="sxs-lookup"><span data-stu-id="8cb56-158">Avoid anticipations, bounces, rubberband, or other effects that emulate natural world physics.</span></span>|

![左 gif 显示使用缓和淡化效果加载磁贴，右 gif 显示使用退回效果加载磁贴](../images/add-in-motion-character.gif)

## <a name="see-also"></a><span data-ttu-id="8cb56-160">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8cb56-160">See also</span></span>

* [<span data-ttu-id="8cb56-161">Fabric 动画指南</span><span class="sxs-lookup"><span data-stu-id="8cb56-161">Fabric animation guidelines</span></span>](https://developer.microsoft.com/fabric#/styles/animations)
* [<span data-ttu-id="8cb56-162">通用 Windows 平台应用动作</span><span class="sxs-lookup"><span data-stu-id="8cb56-162">Motion for Universal Windows Platform apps</span></span>](https://docs.microsoft.com/windows/uwp/design/motion)

