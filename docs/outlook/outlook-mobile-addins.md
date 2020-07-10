---
title: 适用于 Outlook Mobile 的 Outlook 外接程序
description: 在所有 Microsoft 365 商业版帐户、Outlook.com 帐户以及支持即将向 gmail 帐户提供支持的 Outlook mobile 外接程序。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 34fbb01d596c4da38fe81438088cd71d8c7e152a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093894"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="19752-103">适用于 Outlook Mobile 的外接程序</span><span class="sxs-lookup"><span data-stu-id="19752-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="19752-104">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints.</span><span class="sxs-lookup"><span data-stu-id="19752-104">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints.</span></span> <span data-ttu-id="19752-105">If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="19752-105">If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="19752-106">在所有 Microsoft 365 商业版帐户、Outlook.com 帐户以及支持即将向 Gmail 帐户提供支持的 Outlook mobile 外接程序。</span><span class="sxs-lookup"><span data-stu-id="19752-106">Outlook mobile add-ins are supported on all Microsoft 365 business accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="19752-107">**iOS 版 Outlook 中的任务窗格示例**</span><span class="sxs-lookup"><span data-stu-id="19752-107">**An example task pane in Outlook on iOS**</span></span>

![iOS 版 Outlook 中任务窗格的屏幕截图](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="19752-109">**Android 版 Outlook 中的任务窗格示例**</span><span class="sxs-lookup"><span data-stu-id="19752-109">**An example task pane in Outlook on Android**</span></span>

![Android 版 Outlook 中任务窗格的屏幕截图](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="19752-111">外接程序在移动浏览器中的 Outlook 的新式版本中不起作用。</span><span class="sxs-lookup"><span data-stu-id="19752-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="19752-112">有关详细信息，请参阅[正在升级移动浏览器上的 Outlook](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)。</span><span class="sxs-lookup"><span data-stu-id="19752-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="19752-113">在移动电话上会有什么不同？</span><span class="sxs-lookup"><span data-stu-id="19752-113">What's different on mobile?</span></span>

- <span data-ttu-id="19752-114">The small size and quick interactions make designing for mobile a challenge.</span><span class="sxs-lookup"><span data-stu-id="19752-114">The small size and quick interactions make designing for mobile a challenge.</span></span> <span data-ttu-id="19752-115">To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span><span class="sxs-lookup"><span data-stu-id="19752-115">To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="19752-116">外接程序**必须**遵循 [UI 准则](outlook-addin-design.md)。</span><span class="sxs-lookup"><span data-stu-id="19752-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="19752-117">外接程序的方案**必须**[能够在移动电话上实现](#what-makes-a-good-scenario-for-mobile-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="19752-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="19752-118">通常情况下，仅支持邮件阅读模式。</span><span class="sxs-lookup"><span data-stu-id="19752-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="19752-119">这意味着， `MobileMessageReadCommandSurface` 您应在清单的移动部分中声明唯一的[ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) 。</span><span class="sxs-lookup"><span data-stu-id="19752-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="19752-120">但是，"约会组织者" 模式受联机会议提供程序集成的外接程序支持，而这些外接程序则声明[MobileOnlineMeetingCommandSurface 扩展点](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)。</span><span class="sxs-lookup"><span data-stu-id="19752-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span></span> <span data-ttu-id="19752-121">有关此方案的详细信息，请参阅[创建适用于联机会议提供商文章的 Outlook mobile 外](online-meeting.md)接程序。</span><span class="sxs-lookup"><span data-stu-id="19752-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="19752-122">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server.</span><span class="sxs-lookup"><span data-stu-id="19752-122">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server.</span></span> <span data-ttu-id="19752-123">If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls.</span><span class="sxs-lookup"><span data-stu-id="19752-123">If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls.</span></span> <span data-ttu-id="19752-124">For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="19752-124">For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="19752-125">如果将外接程序和清单中的 [MobileFormFactor](../reference/manifest/mobileformfactor.md) 一起提交至应用商店，则需要同意我们添加针对 iOS 上的外接程序的开发人员附录，并且必须提交你的 Apple 开发人员 ID 以进行验证。</span><span class="sxs-lookup"><span data-stu-id="19752-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="19752-126">最后，清单将需要声明 `MobileFormFactor`，并包含正确的[控件](../reference/manifest/control.md)和[图标大小](../reference/manifest/icon.md)类型。</span><span class="sxs-lookup"><span data-stu-id="19752-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="19752-127">适用于移动外接程序的优秀方案应具备哪些特点？</span><span class="sxs-lookup"><span data-stu-id="19752-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="19752-128">Remember that the average Outlook session length on a phone is much shorter than on a PC.</span><span class="sxs-lookup"><span data-stu-id="19752-128">Remember that the average Outlook session length on a phone is much shorter than on a PC.</span></span> <span data-ttu-id="19752-129">That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span><span class="sxs-lookup"><span data-stu-id="19752-129">That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="19752-130">以下是在 Outlook Mobile 中可用的方案示例。</span><span class="sxs-lookup"><span data-stu-id="19752-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="19752-131">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately.</span><span class="sxs-lookup"><span data-stu-id="19752-131">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately.</span></span> <span data-ttu-id="19752-132">Example: a CRM add-in that lets the user see customer information and share appropriate information.</span><span class="sxs-lookup"><span data-stu-id="19752-132">Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="19752-133">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system.</span><span class="sxs-lookup"><span data-stu-id="19752-133">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system.</span></span> <span data-ttu-id="19752-134">Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span><span class="sxs-lookup"><span data-stu-id="19752-134">Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="19752-135">**从 iOS 上的电子邮件创建 Trello 卡片的用户交互示例**</span><span class="sxs-lookup"><span data-stu-id="19752-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![显示用户与 iOS 上的 Outlook Mobile 外接程序交互的动态 GIF](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="19752-137">**从 Android 上的电子邮件创建 Trello 卡片的用户交互示例**</span><span class="sxs-lookup"><span data-stu-id="19752-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![显示用户与 Android 上的 Outlook Mobile 外接程序交互的动态 GIF](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="19752-139">在移动电话上测试外接程序</span><span class="sxs-lookup"><span data-stu-id="19752-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="19752-140">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account.</span><span class="sxs-lookup"><span data-stu-id="19752-140">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account.</span></span> <span data-ttu-id="19752-141">In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest.</span><span class="sxs-lookup"><span data-stu-id="19752-141">In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest.</span></span> <span data-ttu-id="19752-142">Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span><span class="sxs-lookup"><span data-stu-id="19752-142">Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="19752-143">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets.</span><span class="sxs-lookup"><span data-stu-id="19752-143">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets.</span></span> <span data-ttu-id="19752-144">You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span><span class="sxs-lookup"><span data-stu-id="19752-144">You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="19752-145">由于您可能没有使用的工具，移动的故障排除可能非常困难。</span><span class="sxs-lookup"><span data-stu-id="19752-145">Troubleshooting on mobile can be hard since you may not have the tools you're used to.</span></span> <span data-ttu-id="19752-146">但是，在 iOS 上进行故障排除的一种方法是使用 Fiddler (请参阅本教程，了解如何[在 ios 设备) 使用它](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)。</span><span class="sxs-lookup"><span data-stu-id="19752-146">However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span></span>

## <a name="next-steps"></a><span data-ttu-id="19752-147">后续步骤</span><span class="sxs-lookup"><span data-stu-id="19752-147">Next steps</span></span>

<span data-ttu-id="19752-148">了解如何：</span><span class="sxs-lookup"><span data-stu-id="19752-148">Learn how to:</span></span>

- <span data-ttu-id="19752-149">[向外接程序的清单添加移动支持](add-mobile-support.md)。</span><span class="sxs-lookup"><span data-stu-id="19752-149">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="19752-150">[为外接程序设计出色的移动体验](outlook-addin-design.md)。</span><span class="sxs-lookup"><span data-stu-id="19752-150">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="19752-151">[从外接程序获取访问令牌并调用 Outlook REST API](use-rest-api.md)。</span><span class="sxs-lookup"><span data-stu-id="19752-151">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
