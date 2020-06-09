---
title: Office 加载项辅助功能指南
description: 了解如何使你的 Office 外接程序可供所有用户访问。
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: 889563af8ab5f7bbcd4037eedb42933369a92cf2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607990"
---
# <a name="accessibility-guidelines"></a><span data-ttu-id="a3bda-103">辅助功能准则</span><span class="sxs-lookup"><span data-stu-id="a3bda-103">Accessibility guidelines</span></span>

<span data-ttu-id="a3bda-p101">在设计和开发 Office 外接程序时，你将需要确保所有潜在用户和客户都能够成功使用你的外接程序。请应用以下准则，确保所有目标用户都能访问你的解决方案。</span><span class="sxs-lookup"><span data-stu-id="a3bda-p101">As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.</span></span>

## <a name="design-for-multiple-input-methods"></a><span data-ttu-id="a3bda-106">针对多种输入方式的设计</span><span class="sxs-lookup"><span data-stu-id="a3bda-106">Design for multiple input methods</span></span>

- <span data-ttu-id="a3bda-p102">确保用户可以仅通过键盘执行操作。用户应该能够使用 Tab 键和箭头键组合移动到页面上的所有可操作元素。</span><span class="sxs-lookup"><span data-stu-id="a3bda-p102">Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.</span></span>
- <span data-ttu-id="a3bda-109">在移动设备上，当用户通过触摸操作某个控件时，设备应该提供有用的音频反馈。</span><span class="sxs-lookup"><span data-stu-id="a3bda-109">On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.</span></span>
- <span data-ttu-id="a3bda-110">为所有交互式控件提供有用的标签。</span><span class="sxs-lookup"><span data-stu-id="a3bda-110">Provide helpful labels for all interactive controls.</span></span> 

## <a name="make-your-add-in-easy-to-use"></a><span data-ttu-id="a3bda-111">使你的外接程序易于使用</span><span class="sxs-lookup"><span data-stu-id="a3bda-111">Make your add-in easy to use</span></span>

- <span data-ttu-id="a3bda-112">不依赖于单个属性（例如颜色、大小、形状、位置、方向或声音）在 UI 中传达含义。</span><span class="sxs-lookup"><span data-stu-id="a3bda-112">Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.</span></span>
- <span data-ttu-id="a3bda-113">避免对上下文的意外更改，例如在用户未操作的情况下将焦点移到其他 UI 元素。</span><span class="sxs-lookup"><span data-stu-id="a3bda-113">Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.</span></span>
- <span data-ttu-id="a3bda-114">提供验证、确认或撤消所有绑定操作的方法。</span><span class="sxs-lookup"><span data-stu-id="a3bda-114">Provide a way to verify, confirm, or reverse all binding actions.</span></span>
- <span data-ttu-id="a3bda-115">提供暂停或停止媒体（例如音频和视频）的方法。</span><span class="sxs-lookup"><span data-stu-id="a3bda-115">Provide a way to pause or stop media, such as audio and video.</span></span>
- <span data-ttu-id="a3bda-116">不对用户操作施加时间限制。</span><span class="sxs-lookup"><span data-stu-id="a3bda-116">Do not impose a time limit for user action.</span></span>

## <a name="make-your-add-in-easy-to-see"></a><span data-ttu-id="a3bda-117">使你的外接程序易于查看</span><span class="sxs-lookup"><span data-stu-id="a3bda-117">Make your add-in easy to see</span></span>

- <span data-ttu-id="a3bda-118">避免对颜色的意外更改。</span><span class="sxs-lookup"><span data-stu-id="a3bda-118">Avoid unexpected color changes.</span></span>
- <span data-ttu-id="a3bda-p103">提供有意义且及时的信息，以描述 UI 元素、标题和标头、输入和错误。确保控件名称恰当地描述了控件的意图。</span><span class="sxs-lookup"><span data-stu-id="a3bda-p103">Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.</span></span>
- <span data-ttu-id="a3bda-121">遵循针对颜色对比度的[标准准则](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html)。</span><span class="sxs-lookup"><span data-stu-id="a3bda-121">Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.</span></span>

## <a name="account-for-assistive-technologies"></a><span data-ttu-id="a3bda-122">对辅助技术负责</span><span class="sxs-lookup"><span data-stu-id="a3bda-122">Account for assistive technologies</span></span>

- <span data-ttu-id="a3bda-123">避免使用会干扰辅助技术的功能，例如视觉、音频或其他交互。</span><span class="sxs-lookup"><span data-stu-id="a3bda-123">Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.</span></span>
- <span data-ttu-id="a3bda-p104">请勿以图像格式提供文本。屏幕阅读器无法读取图像中的文本。</span><span class="sxs-lookup"><span data-stu-id="a3bda-p104">Do not provide text in an image format. Screen readers cannot read text within images.</span></span>
- <span data-ttu-id="a3bda-126">为用户提供调整或静音所有音频源的方法。</span><span class="sxs-lookup"><span data-stu-id="a3bda-126">Provide a way for users to adjust or mute all audio sources.</span></span>
- <span data-ttu-id="a3bda-127">为用户提供打开字幕或音频说明与音频源的方法。</span><span class="sxs-lookup"><span data-stu-id="a3bda-127">Provide a way for users to turn on captions or audio description with audio sources.</span></span>
- <span data-ttu-id="a3bda-128">提供用于预警用户的声音替代方法，如视觉提示或振动。</span><span class="sxs-lookup"><span data-stu-id="a3bda-128">Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.</span></span>

## <a name="see-also"></a><span data-ttu-id="a3bda-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a3bda-129">See also</span></span>

- [<span data-ttu-id="a3bda-130">Web 内容辅助功能指南 (WCAG) 2.0</span><span class="sxs-lookup"><span data-stu-id="a3bda-130">Web Content Accessibility Guidelines (WCAG) 2.0</span></span>](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [<span data-ttu-id="a3bda-131">向非 Web 信息和通信技术 (WCAG2ICT) 应用 WCAG 2.0 的指南</span><span class="sxs-lookup"><span data-stu-id="a3bda-131">Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)</span></span>](https://www.w3.org/TR/wcag2ict/)
- [<span data-ttu-id="a3bda-132">关于信息和通信技术 (ICT) 的辅助功能要求的欧洲标准</span><span class="sxs-lookup"><span data-stu-id="a3bda-132">European Standard on accessibility requirements for Information and Communication Technologies (ICT)</span></span>](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
