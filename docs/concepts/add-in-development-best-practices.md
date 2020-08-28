---
title: Office 加载项开发最佳做法
description: 在开发以创建 Office 外接程序时应用最佳实践。
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 72bcf8c7bd994c499c1e85154a383c3a91464a79
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293063"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="ef411-103">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="ef411-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="ef411-p101">有效的外接程序提供独特且极具吸引力的功能，采用具有视觉吸引力的方式扩展 Office 应用程序。若要创建出色的外接程序，需为用户提供极具吸引力的首次使用体验、设计一流的 UI 体验和优化外接程序的性能。将本文中描述的最佳实践应用于创建有助于您的用户快速有效地完成其任务的外接程序。</span><span class="sxs-lookup"><span data-stu-id="ef411-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="ef411-107">提供明确值</span><span class="sxs-lookup"><span data-stu-id="ef411-107">Provide clear value</span></span>

- <span data-ttu-id="ef411-p102">创建可帮助用户快速、高效地完成任务的外接程序。专注于对 Office 应用程序有用的方案。例如：</span><span class="sxs-lookup"><span data-stu-id="ef411-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
 - <span data-ttu-id="ef411-111">使核心创作任务更快、更简单，且中断更少。</span><span class="sxs-lookup"><span data-stu-id="ef411-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
 - <span data-ttu-id="ef411-112">在 Office 内启用新方案。</span><span class="sxs-lookup"><span data-stu-id="ef411-112">Enable new scenarios within Office.</span></span>
 - <span data-ttu-id="ef411-113">在 Office 应用程序中嵌入补充服务。</span><span class="sxs-lookup"><span data-stu-id="ef411-113">Embed complementary services within Office applications.</span></span>
 - <span data-ttu-id="ef411-114">改善 Office 体验来提高工作效率。</span><span class="sxs-lookup"><span data-stu-id="ef411-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="ef411-115">通过[创建极具吸引力的首次运行体验](#create-an-engaging-first-run-experience)，确保用户能够快速明确加载项的价值。</span><span class="sxs-lookup"><span data-stu-id="ef411-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="ef411-p103">创建[有效的 AppSource 一览](/office/dev/store/create-effective-office-store-listings)。在标题和说明中明确介绍加载项的优势。请勿依赖品牌来传达加载项的用途。</span><span class="sxs-lookup"><span data-stu-id="ef411-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>


## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="ef411-119">创建极具吸引力的首次运行体验</span><span class="sxs-lookup"><span data-stu-id="ef411-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="ef411-p104">要用具有高可用性和直观性的首次体验吸引新用户。请注意，用户从商店下载外接程序之后，仍可决定是使用还是放弃该外接程序。</span><span class="sxs-lookup"><span data-stu-id="ef411-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="ef411-p105">明确用户与您的外接程序交互所需执行的步骤。使用视频、泡沫垫、分页面板或其他资源来吸引用户。</span><span class="sxs-lookup"><span data-stu-id="ef411-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="ef411-124">在启动时强调您的外接程序的价值主张，而不只是让用户登录。</span><span class="sxs-lookup"><span data-stu-id="ef411-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="ef411-125">提供用以指导用户的教学 UI，并使您的 UI 富有个性化。</span><span class="sxs-lookup"><span data-stu-id="ef411-125">Provide teaching UI to guide users and make your UI personal.</span></span>

   ![显示没有入门步骤的外接程序旁边具有入门步骤的外接程序任务窗格的屏幕截图](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="ef411-127">如果内容外接程序绑定到用户文档中的数据，请将那些用于向用户显示要使用的数据格式的示例数据或模板包含在内。</span><span class="sxs-lookup"><span data-stu-id="ef411-127">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

   ![显示没有数据的内容外接程序旁边具有数据的内容外接程序的屏幕截图](../images/add-in-title.png)

- <span data-ttu-id="ef411-p106">提供[免费试用版](/office/dev/store/decide-on-a-pricing-model)。如果加载项需要订阅，请让某些功能无需订阅也可使用。</span><span class="sxs-lookup"><span data-stu-id="ef411-p106">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="ef411-p107">让注册非常简单。预先填充某些信息（如电子邮件、显示名称），并跳过电子邮件验证。</span><span class="sxs-lookup"><span data-stu-id="ef411-p107">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="ef411-p108">避免弹出窗口。如果必须使用它们，请引导用户启用弹出窗口。</span><span class="sxs-lookup"><span data-stu-id="ef411-p108">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="ef411-135">如需你在开发首次运行体验时可应用的模式，请参阅[适用于 Office 加载项的 UX 设计模式](../design/first-run-experience-patterns.md)。</span><span class="sxs-lookup"><span data-stu-id="ef411-135">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="ef411-136">使用加载项命令</span><span class="sxs-lookup"><span data-stu-id="ef411-136">Use add-in commands</span></span>

- <span data-ttu-id="ef411-p109">使用加载项命令，为加载项提供相关 UI 入口点。有关详细信息（包括设计最佳做法），请参阅[加载项命令](../design/add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="ef411-p109">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="ef411-139">应用用户体验设计原则</span><span class="sxs-lookup"><span data-stu-id="ef411-139">Apply UX design principles</span></span>

- <span data-ttu-id="ef411-p110">确保你的外接程序的外观和功能很好地补充了 Office 体验。使用 [Office UI Fabric](https://developer.microsoft.com/fabric)。</span><span class="sxs-lookup"><span data-stu-id="ef411-p110">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="ef411-p111">支持内容胜过支持部件版式。避免使用对用户体验毫无价值的不必要的 UI 元素。</span><span class="sxs-lookup"><span data-stu-id="ef411-p111">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="ef411-p112">保持用户处于可控状态。确保用户了解重要的决定，并且可以轻松地倒退外接程序执行的操作。</span><span class="sxs-lookup"><span data-stu-id="ef411-p112">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="ef411-p113">使用品牌唤起用户的信任感和亲切感。但不要过度使用品牌或向用户做广告推销。</span><span class="sxs-lookup"><span data-stu-id="ef411-p113">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="ef411-p114">避免滚动。优化为 1366 x 768 分辨率。</span><span class="sxs-lookup"><span data-stu-id="ef411-p114">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="ef411-150">不包含未授权的图像。</span><span class="sxs-lookup"><span data-stu-id="ef411-150">Do not include unlicensed images.</span></span>

- <span data-ttu-id="ef411-151">在加载项中使用[简单明确的语言](../design/voice-guidelines.md)。</span><span class="sxs-lookup"><span data-stu-id="ef411-151">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="ef411-152">考虑辅助功能 - 方便所有用户都可以与加载项轻松交互，并提供屏幕阅读器等辅助技术。</span><span class="sxs-lookup"><span data-stu-id="ef411-152">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="ef411-p115">针对所有平台和输入方法（包括鼠标/键盘和 [触摸](#optimize-for-touch)）的设计。确保 UI 可响应不同的外观设置。</span><span class="sxs-lookup"><span data-stu-id="ef411-p115">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="ef411-155">触摸优化</span><span class="sxs-lookup"><span data-stu-id="ef411-155">Optimize for touch</span></span>

- <span data-ttu-id="ef411-156">使用 [context.touchenabled](/javascript/api/office/office.context#touchenabled) 属性可检测您的外接程序在其上运行的 Office 应用程序是否已启用。</span><span class="sxs-lookup"><span data-stu-id="ef411-156">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="ef411-157">Outlook 不支持此属性。</span><span class="sxs-lookup"><span data-stu-id="ef411-157">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="ef411-p116">确保所有控件都相应符合触控交互的尺寸大小。例如，按钮有足够大的触摸目标，且输入框要足够大，方便用户输入。</span><span class="sxs-lookup"><span data-stu-id="ef411-p116">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="ef411-160">不依赖于诸如悬停或用鼠标右键单击等非触摸式输入方法。</span><span class="sxs-lookup"><span data-stu-id="ef411-160">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="ef411-p117">确保外接程序可以在纵向和横向模式中正常工作。请注意在触控设备上，外接程序的一部分可能通过软键盘隐藏。</span><span class="sxs-lookup"><span data-stu-id="ef411-p117">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="ef411-163">使用[旁加载](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)在实际设备上测试加载项。</span><span class="sxs-lookup"><span data-stu-id="ef411-163">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="ef411-164">若要对设计元素使用 [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric)，需要处理其中许多元素。</span><span class="sxs-lookup"><span data-stu-id="ef411-164">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="ef411-165">优化和监视加载项性能</span><span class="sxs-lookup"><span data-stu-id="ef411-165">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="ef411-p118">创建快速 UI 响应的感觉。外接程序的加载时间应在 500 毫秒以内。</span><span class="sxs-lookup"><span data-stu-id="ef411-p118">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="ef411-168">确保所有用户交互响应时长都在一秒内。</span><span class="sxs-lookup"><span data-stu-id="ef411-168">Ensure that all user interactions respond in under one second.</span></span>

-  <span data-ttu-id="ef411-169">为长时间运行的操作提供加载指示器。</span><span class="sxs-lookup"><span data-stu-id="ef411-169">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="ef411-p119">将 CDN 用于主机图像、资源和公用库。尽可能地从一个位置进行加载。</span><span class="sxs-lookup"><span data-stu-id="ef411-p119">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="ef411-p120">请按照标准 Web 实践来优化您的网页。在生产中，仅使用库的缩小版本。仅加载所需的资源，并优化加载资源的方式。</span><span class="sxs-lookup"><span data-stu-id="ef411-p120">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="ef411-p121">如果操作执行需要一段时间才能完成，请向用户提供反馈。请注意下表中列出的阈值。有关详细信息，请参阅 [Office 加载项的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)。</span><span class="sxs-lookup"><span data-stu-id="ef411-p121">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="ef411-178">**交互类**</span><span class="sxs-lookup"><span data-stu-id="ef411-178">**Interaction class**</span></span>|<span data-ttu-id="ef411-179">**目标**</span><span class="sxs-lookup"><span data-stu-id="ef411-179">**Target**</span></span>|<span data-ttu-id="ef411-180">**上限**</span><span class="sxs-lookup"><span data-stu-id="ef411-180">**Upper bound**</span></span>|<span data-ttu-id="ef411-181">**人类感知**</span><span class="sxs-lookup"><span data-stu-id="ef411-181">**Human perception**</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="ef411-182">即时</span><span class="sxs-lookup"><span data-stu-id="ef411-182">Instant</span></span>|<span data-ttu-id="ef411-183"><=50 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-183"><=50 ms</span></span>|<span data-ttu-id="ef411-184">100 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-184">100 ms</span></span>|<span data-ttu-id="ef411-185">没有明显的延迟。</span><span class="sxs-lookup"><span data-stu-id="ef411-185">No noticeable delay.</span></span>|
  |<span data-ttu-id="ef411-186">快速</span><span class="sxs-lookup"><span data-stu-id="ef411-186">Fast</span></span>|<span data-ttu-id="ef411-187">50-100 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-187">50-100 ms</span></span>|<span data-ttu-id="ef411-188">200 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-188">200 ms</span></span>|<span data-ttu-id="ef411-p122">最小限度的明显延迟。不需要反馈。</span><span class="sxs-lookup"><span data-stu-id="ef411-p122">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="ef411-191">典型</span><span class="sxs-lookup"><span data-stu-id="ef411-191">Typical</span></span>|<span data-ttu-id="ef411-192">100-300 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-192">100-300 ms</span></span>|<span data-ttu-id="ef411-193">500 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-193">500 ms</span></span>|<span data-ttu-id="ef411-p123">较快，但不够快，不能称之为快速。不需要反馈。</span><span class="sxs-lookup"><span data-stu-id="ef411-p123">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="ef411-196">快速响应</span><span class="sxs-lookup"><span data-stu-id="ef411-196">Responsive</span></span>|<span data-ttu-id="ef411-197">300-500 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-197">300-500 ms</span></span>|<span data-ttu-id="ef411-198">1 秒</span><span class="sxs-lookup"><span data-stu-id="ef411-198">1 second</span></span>|<span data-ttu-id="ef411-p124">不快，但仍然感觉反应灵敏。不需要反馈。</span><span class="sxs-lookup"><span data-stu-id="ef411-p124">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="ef411-201">连续</span><span class="sxs-lookup"><span data-stu-id="ef411-201">Continuous</span></span>|<span data-ttu-id="ef411-202">> 500 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-202">>500 ms</span></span>|<span data-ttu-id="ef411-203">5 秒</span><span class="sxs-lookup"><span data-stu-id="ef411-203">5 seconds</span></span>|<span data-ttu-id="ef411-p125">中等等待时间，不再感觉反应灵敏。可能需要反馈。</span><span class="sxs-lookup"><span data-stu-id="ef411-p125">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="ef411-206">受限</span><span class="sxs-lookup"><span data-stu-id="ef411-206">Captive</span></span>|<span data-ttu-id="ef411-207">> 500 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-207">>500 ms</span></span>|<span data-ttu-id="ef411-208">10 秒</span><span class="sxs-lookup"><span data-stu-id="ef411-208">10 seconds</span></span>|<span data-ttu-id="ef411-p126">较长，但不足以执行其他操作。可能需要反馈。</span><span class="sxs-lookup"><span data-stu-id="ef411-p126">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="ef411-211">扩展</span><span class="sxs-lookup"><span data-stu-id="ef411-211">Extended</span></span>|<span data-ttu-id="ef411-212">> 500 毫秒</span><span class="sxs-lookup"><span data-stu-id="ef411-212">>500 ms</span></span>|<span data-ttu-id="ef411-213">> 10 秒</span><span class="sxs-lookup"><span data-stu-id="ef411-213">>10 seconds</span></span>|<span data-ttu-id="ef411-p127">长到足以在等待时执行其他操作。可能需要反馈。</span><span class="sxs-lookup"><span data-stu-id="ef411-p127">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="ef411-216">长时间运行</span><span class="sxs-lookup"><span data-stu-id="ef411-216">Long running</span></span>|<span data-ttu-id="ef411-217">> 5 秒</span><span class="sxs-lookup"><span data-stu-id="ef411-217">>5 seconds</span></span>|<span data-ttu-id="ef411-218">> 1 分钟</span><span class="sxs-lookup"><span data-stu-id="ef411-218">>1 minute</span></span>|<span data-ttu-id="ef411-219">用户当然可以执行其他操作。</span><span class="sxs-lookup"><span data-stu-id="ef411-219">Users will certainly do something else.</span></span>|

- <span data-ttu-id="ef411-220">监视您的服务运行状况，并使用遥测监视用户的成功。</span><span class="sxs-lookup"><span data-stu-id="ef411-220">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="ef411-221">最大限度地减少外接加载项与 Office 文档之间的数据交换。</span><span class="sxs-lookup"><span data-stu-id="ef411-221">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="ef411-222">有关详细信息，请参阅 [避免在循环中使用 context. sync 方法](correlated-objects-pattern.md)。</span><span class="sxs-lookup"><span data-stu-id="ef411-222">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="ef411-223">加载项市场营销</span><span class="sxs-lookup"><span data-stu-id="ef411-223">Market your add-in</span></span>

- <span data-ttu-id="ef411-p129">将加载项发布到 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center)，并在网站中[对它进行宣传](/office/dev/store/promote-your-office-store-solution)。创建[有效的 AppSource 一览](/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="ef411-p129">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="ef411-p130">使用简洁且富有描述性的加载项标题。字符数不得超过 128 个。</span><span class="sxs-lookup"><span data-stu-id="ef411-p130">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="ef411-p131">为您的外接程序撰写简短且富有吸引力的描述。回答"此外接程序解决哪些问题？"这一问题。</span><span class="sxs-lookup"><span data-stu-id="ef411-p131">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="ef411-p132">在您的标题和说明中传达外接程序的价值主张。不要依赖于您的品牌。</span><span class="sxs-lookup"><span data-stu-id="ef411-p132">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="ef411-232">创建有助于用户查找和使用加载项的网站。</span><span class="sxs-lookup"><span data-stu-id="ef411-232">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="ef411-233">使用支持 Internet Explorer 的 JavaScript</span><span class="sxs-lookup"><span data-stu-id="ef411-233">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="ef411-234">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ef411-234">See also</span></span>

- [<span data-ttu-id="ef411-235">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="ef411-235">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
