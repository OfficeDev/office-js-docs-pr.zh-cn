---
title: Office 加载项开发最佳做法
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: efb637cfd31f95688d6b0db0e17b24fefe51c2fc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944646"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="420d6-102">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="420d6-102">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="420d6-p101">有效的外接程序提供独特且极具吸引力的功能，采用具有视觉吸引力的方式扩展 Office 应用程序。若要创建出色的外接程序，需为用户提供极具吸引力的首次使用体验、设计一流的 UI 体验和优化外接程序的性能。将本文中描述的最佳实践应用于创建有助于您的用户快速有效地完成其任务的外接程序。</span><span class="sxs-lookup"><span data-stu-id="420d6-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

> [!NOTE]
> <span data-ttu-id="420d6-p102">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="420d6-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="provide-clear-value"></a><span data-ttu-id="420d6-108">提供明确值</span><span class="sxs-lookup"><span data-stu-id="420d6-108">Provide clear value</span></span>

- <span data-ttu-id="420d6-p103">创建可帮助用户快速、高效地完成任务的外接程序。专注于对 Office 应用程序有用的方案。例如：</span><span class="sxs-lookup"><span data-stu-id="420d6-p103">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
 - <span data-ttu-id="420d6-112">使核心创作任务更快、更简单，且中断更少。</span><span class="sxs-lookup"><span data-stu-id="420d6-112">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
 - <span data-ttu-id="420d6-113">在 Office 内启用新方案。</span><span class="sxs-lookup"><span data-stu-id="420d6-113">Enable new scenarios within Office.</span></span>
 - <span data-ttu-id="420d6-114">在 Office 主机内嵌入补充服务。</span><span class="sxs-lookup"><span data-stu-id="420d6-114">Embed complementary services within Office hosts.</span></span>
 - <span data-ttu-id="420d6-115">改善 Office 体验来提高工作效率。</span><span class="sxs-lookup"><span data-stu-id="420d6-115">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="420d6-116">通过[创建极具吸引力的首次运行体验](#create-an-engaging-first-run-experience)，确保用户能够快速明确加载项的价值。</span><span class="sxs-lookup"><span data-stu-id="420d6-116">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="420d6-p104">创建[有效的 AppSource 一览](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)。在标题和说明中明确介绍加载项的优势。请勿依赖品牌来传达加载项的用途。</span><span class="sxs-lookup"><span data-stu-id="420d6-p104">Create an [effective AppSource listing](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>


## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="420d6-120">创建极具吸引力的首次运行体验</span><span class="sxs-lookup"><span data-stu-id="420d6-120">Create an engaging first-run experience</span></span>

- <span data-ttu-id="420d6-p105">要用具有高可用性和直观性的首次体验吸引新用户。请注意，用户从商店下载外接程序之后，仍可决定是使用还是放弃该外接程序。</span><span class="sxs-lookup"><span data-stu-id="420d6-p105">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="420d6-p106">明确用户与您的外接程序交互所需执行的步骤。使用视频、泡沫垫、分页面板或其他资源来吸引用户。</span><span class="sxs-lookup"><span data-stu-id="420d6-p106">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="420d6-125">在启动时强调您的外接程序的价值主张，而不只是让用户登录。</span><span class="sxs-lookup"><span data-stu-id="420d6-125">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="420d6-126">提供用以指导用户的教学 UI，并使您的 UI 富有个性化。</span><span class="sxs-lookup"><span data-stu-id="420d6-126">Provide teaching UI to guide users and make your UI personal.</span></span>

   ![显示没有入门步骤的外接程序旁边具有入门步骤的外接程序任务窗格的屏幕截图](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="420d6-128">如果内容外接程序绑定到用户文档中的数据，请将那些用于向用户显示要使用的数据格式的示例数据或模板包含在内。</span><span class="sxs-lookup"><span data-stu-id="420d6-128">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

   ![显示没有数据的内容外接程序旁边具有数据的内容外接程序的屏幕截图](../images/add-in-title.png)

- <span data-ttu-id="420d6-p107">提供[免费试用版](https://docs.microsoft.com/office/dev/store/decide-on-a-pricing-model#office-store-pricing-options)。如果加载项需要订阅，请让某些功能无需订阅也可使用。</span><span class="sxs-lookup"><span data-stu-id="420d6-p107">Offer [free trials](https://docs.microsoft.com/office/dev/store/decide-on-a-pricing-model#office-store-pricing-options). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="420d6-p108">让注册非常简单。预先填充某些信息（如电子邮件、显示名称），并跳过电子邮件验证。</span><span class="sxs-lookup"><span data-stu-id="420d6-p108">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="420d6-p109">避免弹出窗口。如果您必须使用它们，请引导用户启用您的弹出窗口。</span><span class="sxs-lookup"><span data-stu-id="420d6-p109">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

- <span data-ttu-id="420d6-136">使用[单一登录 (SSO) 身份验证](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-identity-token)。</span><span class="sxs-lookup"><span data-stu-id="420d6-136">Use [single sign-on (SSO) authentication](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-identity-token).</span></span>

<span data-ttu-id="420d6-137">对于说明你在开发首次运行体验时可以应用的模式的模板，请参阅[适用于 Office 外接程序的 UX 设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)。</span><span class="sxs-lookup"><span data-stu-id="420d6-137">For templates that illustrate patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="420d6-138">使用加载项命令</span><span class="sxs-lookup"><span data-stu-id="420d6-138">Use add-in commands</span></span>

- <span data-ttu-id="420d6-p110">使用加载项命令，为加载项提供相关 UI 入口点。有关详细信息（包括设计最佳做法），请参阅[加载项命令](../design/add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="420d6-p110">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="420d6-141">应用用户体验设计原则</span><span class="sxs-lookup"><span data-stu-id="420d6-141">Apply UX design principles</span></span>

- <span data-ttu-id="420d6-p111">确保你的外接程序的外观和功能很好地补充了 Office 体验。使用 [Office UI Fabric](https://developer.microsoft.com/fabric)。</span><span class="sxs-lookup"><span data-stu-id="420d6-p111">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="420d6-p112">支持内容胜过支持部件版式。避免使用对用户体验毫无价值的不必要的 UI 元素。</span><span class="sxs-lookup"><span data-stu-id="420d6-p112">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="420d6-p113">保持用户处于可控状态。确保用户了解重要的决定，并且可以轻松地倒退外接程序执行的操作。</span><span class="sxs-lookup"><span data-stu-id="420d6-p113">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="420d6-p114">使用品牌唤起用户的信任感和亲切感。但不要过度使用品牌或向用户做广告推销。</span><span class="sxs-lookup"><span data-stu-id="420d6-p114">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="420d6-p115">避免滚动。优化为 1366 x 768 分辨率。</span><span class="sxs-lookup"><span data-stu-id="420d6-p115">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="420d6-152">不包含未授权的图像。</span><span class="sxs-lookup"><span data-stu-id="420d6-152">Do not include unlicensed images.</span></span>

- <span data-ttu-id="420d6-153">在加载项中使用[简单明确的语言](../design/voice-guidelines.md)。</span><span class="sxs-lookup"><span data-stu-id="420d6-153">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="420d6-154">考虑辅助功能 - 方便所有用户都可以与加载项轻松交互，并提供屏幕阅读器等辅助技术。</span><span class="sxs-lookup"><span data-stu-id="420d6-154">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="420d6-p116">针对所有平台和输入方法（包括鼠标/键盘和 [触摸](#optimize-for-touch)）的设计。确保您的 UI 可响应不同的外观设置。</span><span class="sxs-lookup"><span data-stu-id="420d6-p116">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

<span data-ttu-id="420d6-157">对于应用你在开发外接程序时可以使用和自定义的设计原则的模板，请参阅[适用于 Office 外接程序的 UX 设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)。</span><span class="sxs-lookup"><span data-stu-id="420d6-157">For templates that apply design principles that you can use and customize as you develop your add-in, see [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="420d6-158">触摸优化</span><span class="sxs-lookup"><span data-stu-id="420d6-158">Optimize for touch</span></span>

- <span data-ttu-id="420d6-159">使用 [Context.touchEnabled](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) 属性检测运行加载项的主机应用是否已启用触控。</span><span class="sxs-lookup"><span data-stu-id="420d6-159">Use the [Context.touchEnabled](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) property to detect whether the host application your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="420d6-160">Outlook 不支持此属性。</span><span class="sxs-lookup"><span data-stu-id="420d6-160">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="420d6-p117">确保所有控件都相应符合触控交互的尺寸大小。例如，按钮有足够大的触摸目标，且输入框要足够大，方便用户输入。</span><span class="sxs-lookup"><span data-stu-id="420d6-p117">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="420d6-163">不依赖于诸如悬停或用鼠标右键单击等非触摸式输入方法。</span><span class="sxs-lookup"><span data-stu-id="420d6-163">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="420d6-p118">确保外接程序可以在纵向和横向模式中正常工作。请注意在触控设备上，外接程序的一部分可能通过软键盘隐藏。</span><span class="sxs-lookup"><span data-stu-id="420d6-p118">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="420d6-166">使用[旁加载](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)在实际设备上测试加载项。</span><span class="sxs-lookup"><span data-stu-id="420d6-166">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="420d6-167">若要对设计元素使用 [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric)，需要处理其中许多元素。</span><span class="sxs-lookup"><span data-stu-id="420d6-167">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="420d6-168">优化和监视加载项性能</span><span class="sxs-lookup"><span data-stu-id="420d6-168">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="420d6-p119">创建快速 UI 响应的感觉。外接程序的加载时间应在 500 毫秒以内。</span><span class="sxs-lookup"><span data-stu-id="420d6-p119">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="420d6-171">确保所有用户交互响应时长都在一秒内。</span><span class="sxs-lookup"><span data-stu-id="420d6-171">Ensure that all user interactions respond in under one second.</span></span>

-  <span data-ttu-id="420d6-172">为长时间运行的操作提供加载指示器。</span><span class="sxs-lookup"><span data-stu-id="420d6-172">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="420d6-p120">将 CDN 用于主机图像、资源和公用库。尽可能地从一个位置进行加载。</span><span class="sxs-lookup"><span data-stu-id="420d6-p120">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="420d6-p121">请按照标准 Web 实践来优化您的网页。在生产中，仅使用库的缩小版本。仅加载所需的资源，并优化加载资源的方式。</span><span class="sxs-lookup"><span data-stu-id="420d6-p121">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="420d6-p122">如果操作执行需要一段时间才能完成，请向用户提供反馈。请注意下表中列出的阈值。有关详细信息，请参阅 [Office 加载项的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)。</span><span class="sxs-lookup"><span data-stu-id="420d6-p122">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="420d6-181">**交互类**</span><span class="sxs-lookup"><span data-stu-id="420d6-181">**Interaction class**</span></span>|<span data-ttu-id="420d6-182">**目标**</span><span class="sxs-lookup"><span data-stu-id="420d6-182">**Target**</span></span>|<span data-ttu-id="420d6-183">**上限**</span><span class="sxs-lookup"><span data-stu-id="420d6-183">**Upper bound**</span></span>|<span data-ttu-id="420d6-184">**人类感知**</span><span class="sxs-lookup"><span data-stu-id="420d6-184">**Human perception**</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="420d6-185">即时</span><span class="sxs-lookup"><span data-stu-id="420d6-185">Instant</span></span>|<span data-ttu-id="420d6-186"><=50 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-186"><=50 ms</span></span>|<span data-ttu-id="420d6-187">100 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-187">100 ms</span></span>|<span data-ttu-id="420d6-188">没有明显的延迟。</span><span class="sxs-lookup"><span data-stu-id="420d6-188">No noticeable delay.</span></span>|
  |<span data-ttu-id="420d6-189">快速</span><span class="sxs-lookup"><span data-stu-id="420d6-189">Fast</span></span>|<span data-ttu-id="420d6-190">50-100 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-190">50-100 ms</span></span>|<span data-ttu-id="420d6-191">200 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-191">200 ms</span></span>|<span data-ttu-id="420d6-p123">最小限度的明显延迟。不需要反馈。</span><span class="sxs-lookup"><span data-stu-id="420d6-p123">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="420d6-194">典型</span><span class="sxs-lookup"><span data-stu-id="420d6-194">Typical</span></span>|<span data-ttu-id="420d6-195">100-300 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-195">100-300 ms</span></span>|<span data-ttu-id="420d6-196">500 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-196">500 ms</span></span>|<span data-ttu-id="420d6-p124">较快，但不够快，不能称之为快速。不需要反馈。</span><span class="sxs-lookup"><span data-stu-id="420d6-p124">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="420d6-199">快速响应</span><span class="sxs-lookup"><span data-stu-id="420d6-199">Responsive</span></span>|<span data-ttu-id="420d6-200">300-500 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-200">300-500 ms</span></span>|<span data-ttu-id="420d6-201">1 秒</span><span class="sxs-lookup"><span data-stu-id="420d6-201">1 second</span></span>|<span data-ttu-id="420d6-p125">不快，但仍然感觉反应灵敏。不需要反馈。</span><span class="sxs-lookup"><span data-stu-id="420d6-p125">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="420d6-204">连续</span><span class="sxs-lookup"><span data-stu-id="420d6-204">Continuous</span></span>|<span data-ttu-id="420d6-205">> 500 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-205">>500 ms</span></span>|<span data-ttu-id="420d6-206">5 秒</span><span class="sxs-lookup"><span data-stu-id="420d6-206">5 seconds</span></span>|<span data-ttu-id="420d6-p126">中等等待时间，不再感觉反应灵敏。可能需要反馈。</span><span class="sxs-lookup"><span data-stu-id="420d6-p126">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="420d6-209">受限</span><span class="sxs-lookup"><span data-stu-id="420d6-209">Captive</span></span>|<span data-ttu-id="420d6-210">> 500 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-210">>500 ms</span></span>|<span data-ttu-id="420d6-211">10 秒</span><span class="sxs-lookup"><span data-stu-id="420d6-211">10 seconds</span></span>|<span data-ttu-id="420d6-p127">较长，但不足以执行其他操作。可能需要反馈。</span><span class="sxs-lookup"><span data-stu-id="420d6-p127">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="420d6-214">扩展</span><span class="sxs-lookup"><span data-stu-id="420d6-214">Extended</span></span>|<span data-ttu-id="420d6-215">> 500 毫秒</span><span class="sxs-lookup"><span data-stu-id="420d6-215">>500 ms</span></span>|<span data-ttu-id="420d6-216">> 10 秒</span><span class="sxs-lookup"><span data-stu-id="420d6-216">>10 seconds</span></span>|<span data-ttu-id="420d6-p128">长到足以在等待时执行其他操作。可能需要反馈。</span><span class="sxs-lookup"><span data-stu-id="420d6-p128">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="420d6-219">长时间运行</span><span class="sxs-lookup"><span data-stu-id="420d6-219">Long running</span></span>|<span data-ttu-id="420d6-220">> 5 秒</span><span class="sxs-lookup"><span data-stu-id="420d6-220">5 seconds</span></span>|<span data-ttu-id="420d6-221">> 1 分钟</span><span class="sxs-lookup"><span data-stu-id="420d6-221">>1 minute</span></span>|<span data-ttu-id="420d6-222">用户当然可以执行其他操作。</span><span class="sxs-lookup"><span data-stu-id="420d6-222">Users will certainly do something else.</span></span>|

- <span data-ttu-id="420d6-223">监视您的服务运行状况，并使用遥测监视用户的成功。</span><span class="sxs-lookup"><span data-stu-id="420d6-223">Monitor your service health, and use telemetry to monitor user success.</span></span>


## <a name="market-your-add-in"></a><span data-ttu-id="420d6-224">加载项市场营销</span><span class="sxs-lookup"><span data-stu-id="420d6-224">Market your add-in</span></span>

- <span data-ttu-id="420d6-p129">将加载项发布到 [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)，并在网站中[对它进行宣传](https://docs.microsoft.com/office/dev/store/promote-your-office-store-solution)。创建[有效的 AppSource 一览](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="420d6-p129">Publish your add-in to [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store) and [promote it](https://docs.microsoft.com/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="420d6-p130">使用简洁且富有描述性的加载项标题。字符数不得超过 128 个。</span><span class="sxs-lookup"><span data-stu-id="420d6-p130">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="420d6-p131">为您的外接程序撰写简短且富有吸引力的描述。回答"此外接程序解决哪些问题？"这一问题。</span><span class="sxs-lookup"><span data-stu-id="420d6-p131">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="420d6-p132">在您的标题和说明中传达外接程序的价值主张。不要依赖于您的品牌。</span><span class="sxs-lookup"><span data-stu-id="420d6-p132">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="420d6-233">创建有助于用户查找和使用加载项的网站。</span><span class="sxs-lookup"><span data-stu-id="420d6-233">Create a website to help users find and use your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="420d6-234">另请参阅</span><span class="sxs-lookup"><span data-stu-id="420d6-234">See also</span></span>

- [<span data-ttu-id="420d6-235">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="420d6-235">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
