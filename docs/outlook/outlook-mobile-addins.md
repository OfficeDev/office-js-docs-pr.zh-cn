---
title: 适用于 Outlook Mobile 的 Outlook 外接程序
description: 所有 Office 365 商业帐户、Outlook.com 帐户均支持 Outlook 移动外接程序，并且即将提供对 Gmail 帐户的支持。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 9a7345840d5a26b27f824470efd58d846d0aab11
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44606439"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="9ccd4-103">适用于 Outlook Mobile 的外接程序</span><span class="sxs-lookup"><span data-stu-id="9ccd4-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="9ccd4-p101">现在，外接程序在 Outlook Mobile 上可用，它们使用适用于其他 Outlook 终结点的相同 API。如果已经生成适用于 Outlook 的外接程序，那么则可以很轻松地在 Outlook Mobile 上使用该外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="9ccd4-106">所有 Office 365 商业帐户、Outlook.com 帐户均支持 Outlook 移动外接程序，并且即将提供对 Gmail 帐户的支持。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-106">Outlook mobile add-ins are supported on all Office 365 Commercial accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="9ccd4-107">**iOS 版 Outlook 中的任务窗格示例**</span><span class="sxs-lookup"><span data-stu-id="9ccd4-107">**An example task pane in Outlook on iOS**</span></span>

![iOS 版 Outlook 中任务窗格的屏幕截图](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="9ccd4-109">**Android 版 Outlook 中的任务窗格示例**</span><span class="sxs-lookup"><span data-stu-id="9ccd4-109">**An example task pane in Outlook on Android**</span></span>

![Android 版 Outlook 中任务窗格的屏幕截图](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="9ccd4-111">外接程序在移动浏览器中的 Outlook 的新式版本中不起作用。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="9ccd4-112">有关详细信息，请参阅[正在升级移动浏览器上的 Outlook](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="9ccd4-113">在移动电话上会有什么不同？</span><span class="sxs-lookup"><span data-stu-id="9ccd4-113">What's different on mobile?</span></span>

- <span data-ttu-id="9ccd4-p103">移动电话尺寸小，需要进行快速交互，这为设计适用于移动电话的加载项带来了挑战。为了确保客户体验的质量，我们正在设置严格的验证标准，声明提供移动支持的加载项必须符合这一标准，以便在 AppSource 中获得批准。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p103">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="9ccd4-116">外接程序**必须**遵循 [UI 准则](outlook-addin-design.md)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="9ccd4-117">外接程序的方案**必须**[能够在移动电话上实现](#what-makes-a-good-scenario-for-mobile-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="9ccd4-118">通常情况下，仅支持邮件阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="9ccd4-119">这意味着， `MobileMessageReadCommandSurface` 您应在清单的移动部分中声明唯一的[ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) 。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="9ccd4-120">但是，"约会组织者" 模式受联机会议提供程序集成的外接程序支持，而这些外接程序则声明[MobileOnlineMeetingCommandSurface 扩展点](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span></span> <span data-ttu-id="9ccd4-121">有关此方案的详细信息，请参阅[创建适用于联机会议提供商文章的 Outlook mobile 外](online-meeting.md)接程序。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="9ccd4-p105">[makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API 在移动电话上不受支持，因为移动应用使用 REST API 与服务器进行通信。如果应用后端需要连接到 Exchange 服务器，则可以使用回调令牌进行 REST API 调用。有关详细信息，请参阅[从 Outlook 外接程序使用 Outlook REST API](use-rest-api.md)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p105">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="9ccd4-125">如果将外接程序和清单中的 [MobileFormFactor](../reference/manifest/mobileformfactor.md) 一起提交至应用商店，则需要同意我们添加针对 iOS 上的外接程序的开发人员附录，并且必须提交你的 Apple 开发人员 ID 以进行验证。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="9ccd4-126">最后，清单将需要声明 `MobileFormFactor`，并包含正确的[控件](../reference/manifest/control.md)和[图标大小](../reference/manifest/icon.md)类型。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="9ccd4-127">适用于移动外接程序的优秀方案应具备哪些特点？</span><span class="sxs-lookup"><span data-stu-id="9ccd4-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="9ccd4-p106">请记住，电话上 Outlook 会话的平均长度要比在 PC 上短得多。这意味着外接程序必须快速运行，且方案必须允许用户进入、退出，并继续处理他们的电子邮件工作流。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p106">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="9ccd4-130">以下是在 Outlook Mobile 中可用的方案示例。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="9ccd4-p107">外接程序为 Outlook 带来了有价值的信息，帮助用户会审他们的电子邮件并进行适当地响应。示例：可让用户查看客户信息并共享相应信息的 CRM 外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p107">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="9ccd4-p108">外接程序通过将信息保存到跟踪、协作或类似系统，为用户的电子邮件内容增加价值。示例：允许用户将电子邮件转化为任务项以供项目跟踪，或转化为支持团队的帮助票证的外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p108">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="9ccd4-135">**从 iOS 上的电子邮件创建 Trello 卡片的用户交互示例**</span><span class="sxs-lookup"><span data-stu-id="9ccd4-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![显示用户与 iOS 上的 Outlook Mobile 外接程序交互的动态 GIF](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="9ccd4-137">**从 Android 上的电子邮件创建 Trello 卡片的用户交互示例**</span><span class="sxs-lookup"><span data-stu-id="9ccd4-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![显示用户与 Android 上的 Outlook Mobile 外接程序交互的动态 GIF](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="9ccd4-139">在移动电话上测试外接程序</span><span class="sxs-lookup"><span data-stu-id="9ccd4-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="9ccd4-p109">若要在 Outlook Mobile 上测试加载项，可以将加载项旁加载到 O365 或 Outlook.com 帐户。在 Outlook 网页版中，转到设置齿轮，并选择“**管理集成**”或“**管理加载项**”。在靠近顶部的位置，单击显示的“**单击此处添加自定义加载项**”并上传清单。请确保清单格式正确以包含 `MobileFormFactor`，否则将无法上传。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p109">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="9ccd4-p110">在加载项正常运行后，请务必在不同尺寸的屏幕（包括电话和平板电脑）上测试加载项。应确保加载项符合与对比度、字号和颜色有关的辅助功能准则，并且还适用于屏幕阅读器（如 iOS 上的 VoiceOver 或 Android 上的 TalkBack）。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-p110">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="9ccd4-145">由于您可能没有使用的工具，移动的故障排除可能非常困难。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-145">Troubleshooting on mobile can be hard since you may not have the tools you're used to.</span></span> <span data-ttu-id="9ccd4-146">但是，在 iOS 上进行故障排除的一种方法是使用 Fiddler （请参阅本教程，了解如何[在 ios 设备上使用它](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)）。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-146">However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span></span>

## <a name="next-steps"></a><span data-ttu-id="9ccd4-147">后续步骤</span><span class="sxs-lookup"><span data-stu-id="9ccd4-147">Next steps</span></span>

<span data-ttu-id="9ccd4-148">了解如何：</span><span class="sxs-lookup"><span data-stu-id="9ccd4-148">Learn how to:</span></span>

- <span data-ttu-id="9ccd4-149">[向外接程序的清单添加移动支持](add-mobile-support.md)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-149">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="9ccd4-150">[为外接程序设计出色的移动体验](outlook-addin-design.md)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-150">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="9ccd4-151">[从外接程序获取访问令牌并调用 Outlook REST API](use-rest-api.md)。</span><span class="sxs-lookup"><span data-stu-id="9ccd4-151">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
