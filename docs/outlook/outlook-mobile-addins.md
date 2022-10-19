---
title: 适用于 Outlook Mobile 的 Outlook 外接程序
description: 所有 Microsoft 365 商业帐户和 Outlook.com 帐户都支持 Outlook 移动外接程序。
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca09ba550d8d2ed6e9003e85a8d042f413a6ab52
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607560"
---
# <a name="add-ins-for-outlook-mobile"></a>适用于 Outlook Mobile 的加载项

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

所有 Microsoft 365 商业帐户和 Outlook.com 帐户都支持 Outlook 移动外接程序。 但是，Gmail 帐户目前不提供支持。

**iOS 版 Outlook 中的任务窗格示例**

![iOS 版 Outlook 中任务窗格的屏幕截图。](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Android 版 Outlook 中的任务窗格示例**

![Android 版 Outlook 中任务窗格的屏幕截图。](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>在移动电话上会有什么不同？

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
  - 外接程序 **必须** 遵循 [UI 准则](outlook-addin-design.md)。
  - 外接程序的方案 **必须**[能够在移动电话上实现](#what-makes-a-good-scenario-for-mobile-add-ins)。

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

- 一般情况下，目前仅支持消息读取模式。 这意味着 `MobileMessageReadCommandSurface` ，应在清单的移动部分中声明的唯一 [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) 。 但是，有几个例外情况：
  1. 联机会议提供商集成加载项支持约会组织者模式，而后者则声明 [MobileOnlineMeetingCommandSurface 扩展点](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface)。 有关此方案的详细信息，请参阅 [“为联机会议提供商创建 Outlook 移动外接程序”一](online-meeting.md) 文。
  1. 由记事和客户关系管理提供商创建的集成外接程序支持约会与会者模式 (CRM) 应用程序。 此类加载项应改为声明 [MobileLogEventAppointmentAttendee 扩展点](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)。 有关此方案的详细信息，请参阅 [Outlook 移动外接程序文章中外部应用程序的日志约会说明](mobile-log-appointments.md) 。

- The [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- 如果将外接程序和清单中的 [MobileFormFactor](/javascript/api/manifest/mobileformfactor) 一起提交至应用商店，则需要同意我们添加针对 iOS 上的外接程序的开发人员附录，并且必须提交你的 Apple 开发人员 ID 以进行验证。

- 最后，清单将需要声明 `MobileFormFactor`，并包含正确的[控件](/javascript/api/manifest/control)和[图标大小](/javascript/api/manifest/icon)类型。

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>适用于移动外接程序的优秀方案应具备哪些特点？

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

以下是在 Outlook Mobile 中可用的方案示例。

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**从 iOS 上的电子邮件创建 Trello 卡片的用户交互示例**

![显示用户与 iOS 上的 Outlook Mobile 加载项交互的动画 GIF。](../images/outlook-mobile-addin-interaction.gif)

<br/>

**从 Android 上的电子邮件创建 Trello 卡片的用户交互示例**

![显示用户与 Android 上的 Outlook Mobile 加载项交互的动画 GIF。](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>在移动电话上测试外接程序

若要在 Outlook Mobile 上测试加载项，请首先将 [加载项旁加载](sideload-outlook-add-ins-for-testing.md) 到 Web、Windows 或 Mac 上的 Microsoft 365 或 Outlook.com 帐户。 确保清单格式正确，以包含 `MobileFormFactor` 或不会在移动版 Outlook 客户端中加载。

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

在移动设备上进行故障排除可能很困难，因为你可能没有习惯的工具。 但是，在 iOS 上进行故障排除的一个选项是使用 Fiddler (查看 [本教程，了解如何将其与 iOS 设备) 配合使用](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices) 。

> [!NOTE]
> iPhone 和 Android 智能手机上的新式Outlook 网页版不再需要或可用于测试 Outlook 加载项。此外，在 Outlook on Android、iOS 和具有本地 Exchange 帐户的新式移动 Web 中不支持外接程序。 使用具有经典Outlook 网页版的本地 Exchange 帐户时，某些 iOS 设备仍支持加载项。 有关支持的设备的信息，请参阅[运行 Office 加载项的要求](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet)。

## <a name="next-steps"></a>后续步骤

了解如何：

- [向外接程序的清单添加移动支持](add-mobile-support.md)。
- [为外接程序设计出色的移动体验](outlook-addin-design.md)。
- [从外接程序获取访问令牌并调用 Outlook REST API](use-rest-api.md)。
