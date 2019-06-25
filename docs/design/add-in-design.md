---
title: 设计 Office 加载项
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 766110b9b1ff9d22a783f592f1e38eb848b8b737
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127672"
---
# <a name="design-your-office-add-ins"></a><span data-ttu-id="8478a-102">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="8478a-102">Design your Office Add-ins</span></span>

<span data-ttu-id="8478a-p101">Office 外接程序可通过提供用户可在 Office 客户端内访问的上下文功能来扩展 Office 体验。通过外接程序，用户可以访问 Office 内的第三方功能以完成更多操作，而无需进行成本高昂的上下文切换。</span><span class="sxs-lookup"><span data-stu-id="8478a-p101">Office Add-ins extend the Office experience by providing contextual functionality that users can access within Office clients. Add-ins empower users to get more done by enabling them to access third-party functionality within Office, without costly context switches.</span></span> 

<span data-ttu-id="8478a-p102">你的外接程序 UX 设计必须与 Office 无缝集成，为用户提供高效、自然的交互。利用[外接程序命令](add-in-commands.md)提供对外接程序的访问权限，并应用创建基于 HTML 的自定义 UI 时建议的最佳实践。</span><span class="sxs-lookup"><span data-stu-id="8478a-p102">Your add-in UX design must integrate seamlessly with Office to provide an efficient, natural interaction for your users. Take advantage of [add-in commands](add-in-commands.md) to provide access to your add-in and apply the best practices that we recommend when you create custom HTML-based UI.</span></span>

## <a name="office-design-principles"></a><span data-ttu-id="8478a-107">Office 设计原则</span><span class="sxs-lookup"><span data-stu-id="8478a-107">Office design principles</span></span>

<span data-ttu-id="8478a-p103">Office 应用程序遵循一套常规交互原则。应用共享内容并具有外观和行为相似的元素。此通用性基于一套设计原则。这些原则帮助 Office 团队创建支持客户任务的界面。了解并遵循这些原则将有助于支持 Office 内部的客户目标。</span><span class="sxs-lookup"><span data-stu-id="8478a-p103">Office applications follow a general set of interaction guidelines. The apps share content and have elements that look and behave similarly. This commonality is built on a set of design principles. The principles help the Office team create interfaces that support customers’ tasks. Understanding and following them will help you support your customers’ goals inside of Office.</span></span>

<span data-ttu-id="8478a-113">若要打造积极的加载项体验，请遵循 Office 设计原则：</span><span class="sxs-lookup"><span data-stu-id="8478a-113">Follow the Office design principles to create positive add-in experiences:</span></span>

- <span data-ttu-id="8478a-p104">**对 Office 进行明确设计。** 加载项的功能、外观和感受必须和谐地完善 Office 体验。加载项应该让人感觉就像安装在本机一样。它们应无缝融入 iPad 版 Word 或 PowerPoint 网页版。设计良好的加载项将恰当地融合体验、平台和 Office 应用程序。请考虑使用 Office UI Fabric 作为设计语言。在适当的位置应用文档和 UI 主题。</span><span class="sxs-lookup"><span data-stu-id="8478a-p104">**Design explicitly for Office.** The functionality, look and feel of an add-in must harmoniously complement the Office experience. Add-ins should feel native. They should fit seamlessly into Word on an iPad or PowerPoint Online. A well-designed add-in will be an appropriate blend of your experience, the platform and the Office application. Consider using Office UI Fabric as your design language. Apply document and UI theming where appropriate.</span></span>

- <span data-ttu-id="8478a-p105">**重点关注几个关键任务；好好完成。** 帮助客户在不影响其他工作的情况下完成一项工作。为客户提供真正的价值。与 Office 文档交互时，关注常见用例并认真挑选出用户最受益的。</span><span class="sxs-lookup"><span data-stu-id="8478a-p105">**Focus on a few key tasks; do them well.** Help customers get one job done without getting in the way of other jobs. Provide real value to customers. Focus on common use cases, pick carefully those that benefit users most when interacting with Office documents.</span></span>

- <span data-ttu-id="8478a-p106">**使内容优先于 Chrome。** 使客户的页面、幻灯片或电子表格始终关注体验。外接程序是辅助界面。没有任何辅助 Chrome 应当与外接程序的内容和功能交互。请明智地品牌化你的体验。我们知道这对于向用户提供独特且可识别的功能但避免干扰十分重要。努力将重点集中于内容和任务完成，而非品牌关注。</span><span class="sxs-lookup"><span data-stu-id="8478a-p106">**Favor content over chrome.** Allow customers’ page, slide or spreadsheet to remain the focus of the experience. An add-in is an auxiliary interface. No accessory chrome should interfere with the add-in’s content and functionality. Brand your experience wisely. We know it is important to provide users with a unique, recognizable experience but avoid distraction. Strive to keep the focus on content and task completion, not brand attention.</span></span>

- <span data-ttu-id="8478a-132">**使其方便好用并保持对用户的控制。**</span><span class="sxs-lookup"><span data-stu-id="8478a-132">**Make it enjoyable and keep users in control.**</span></span> <span data-ttu-id="8478a-133">人们喜欢使用实用且外观吸引人的产品。</span><span class="sxs-lookup"><span data-stu-id="8478a-133">People enjoy using products that are both functional and visually appealing.</span></span> <span data-ttu-id="8478a-134">小心地定制你的体验。</span><span class="sxs-lookup"><span data-stu-id="8478a-134">Craft your experience carefully.</span></span> <span data-ttu-id="8478a-135">将每个交互和视觉细节考虑在内，把细节做好。</span><span class="sxs-lookup"><span data-stu-id="8478a-135">Get the details right by considering every interaction and visual detail.</span></span> <span data-ttu-id="8478a-136">允许用户控制其体验。</span><span class="sxs-lookup"><span data-stu-id="8478a-136">Allow users to control their experience.</span></span> <span data-ttu-id="8478a-137">完成任务的必要步骤必须清楚并相互关联。</span><span class="sxs-lookup"><span data-stu-id="8478a-137">The necessary steps to complete a task must be clear and relevant.</span></span> <span data-ttu-id="8478a-138">重要的决定应该是易于理解的。</span><span class="sxs-lookup"><span data-stu-id="8478a-138">Important decisions should be easy to understand.</span></span> <span data-ttu-id="8478a-139">操作应该可以轻松撤消。</span><span class="sxs-lookup"><span data-stu-id="8478a-139">Actions should be easily reversible.</span></span> <span data-ttu-id="8478a-140">外接程序不是一个目标，它是对 Office 功能的增强。</span><span class="sxs-lookup"><span data-stu-id="8478a-140">An add-in is not a destination – it’s an enhancement to Office functionality.</span></span>

- <span data-ttu-id="8478a-p108">**针对所有平台和输入方法进行设计**。外接程序设计用于 Office 支持的所有平台，您的外接程序 UI 应该进行优化，以便跨平台和外形规格运行。支持鼠标/键盘和触摸输入设备，确保您的自定义 HTML UI 响应迅速，可适应不同的外形规格。有关详细信息，请参阅[触摸](../concepts/add-in-development-best-practices.md#optimize-for-touch)。</span><span class="sxs-lookup"><span data-stu-id="8478a-p108">**Design for all platforms and input methods**. Add-ins are designed to work on all the platforms that Office supports, and your add-in UX should be optimized to work across platforms and form factors. Support mouse/keyboard and touch input devices, and ensure that your custom HTML UI is responsive to adapt to different form factors. For more information, see [Touch](../concepts/add-in-development-best-practices.md#optimize-for-touch).</span></span> 

## <a name="see-also"></a><span data-ttu-id="8478a-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8478a-145">See also</span></span>
- <span data-ttu-id="8478a-146">
  [Office UI Fabric](https://developer.microsoft.com/zh-CN/fabric)</span><span class="sxs-lookup"><span data-stu-id="8478a-146">[Office UI Fabric](https://developer.microsoft.com/en-us/fabric)</span></span> 
- [<span data-ttu-id="8478a-147">加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="8478a-147">Add-in development best practices</span></span>](../concepts/add-in-development-best-practices.md)

