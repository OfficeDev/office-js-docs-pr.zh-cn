---
title: 适用于 Outlook Mobile 的 Outlook 外接程序
description: 所有 Microsoft 365 商业帐户和 Outlook.com 帐户都支持 Outlook 移动外接程序。
ms.date: 04/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: dfa314ad91646e2ed4de47cae1bcbb8cfb1f121a
ms.sourcegitcommit: 57258dd38507f791bbb39cbb01d6bbd5a9d226b9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2022
ms.locfileid: "67318800"
---
# <a name="add-ins-for-outlook-mobile"></a>适用于 Outlook Mobile 的外接程序

现在，外接程序在 Outlook Mobile 上可用，它们使用适用于其他 Outlook 终结点的相同 API。如果已经生成适用于 Outlook 的外接程序，那么则可以很轻松地在 Outlook Mobile 上使用该外接程序。

所有 Microsoft 365 商业帐户和 Outlook.com 帐户都支持 Outlook 移动外接程序。 但是，Gmail 帐户目前不提供支持。

**iOS 版 Outlook 中的任务窗格示例**

![iOS 版 Outlook 中任务窗格的屏幕截图。](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Android 版 Outlook 中的任务窗格示例**

![Android 版 Outlook 中任务窗格的屏幕截图。](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>在移动电话上会有什么不同？

- 移动电话尺寸小，需要进行快速交互，这为设计适用于移动电话的加载项带来了挑战。为了确保客户体验的质量，我们正在设置严格的验证标准，声明提供移动支持的加载项必须符合这一标准，以便在 AppSource 中获得批准。
  - 外接程序 **必须** 遵循 [UI 准则](outlook-addin-design.md)。
  - 外接程序的方案 **必须**[能够在移动电话上实现](#what-makes-a-good-scenario-for-mobile-add-ins)。

- 一般情况下，目前仅支持消息读取模式。 这意味着 `MobileMessageReadCommandSurface` ，应在清单的移动部分中声明的唯一 [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) 。 但是，有几个例外情况：
  1. 联机会议提供商集成加载项支持约会组织者模式，而后者则声明 [MobileOnlineMeetingCommandSurface 扩展点](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface)。 有关此方案的详细信息，请参阅 [“为联机会议提供商创建 Outlook 移动外接程序”一](online-meeting.md) 文。
  1. 由记事和客户关系管理提供商创建的集成外接程序支持约会与会者模式 (CRM) 应用程序。 此类加载项应改为声明 [MobileLogEventAppointmentAttendee 扩展点](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)。 有关此方案的详细信息，请参阅 [Outlook 移动外接程序文章中外部应用程序的日志约会说明](mobile-log-appointments.md) 。

- [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API 在移动电话上不受支持，因为移动应用使用 REST API 与服务器进行通信。如果应用后端需要连接到 Exchange 服务器，则可以使用回调令牌进行 REST API 调用。有关详细信息，请参阅[从 Outlook 外接程序使用 Outlook REST API](use-rest-api.md)。

- 如果将外接程序和清单中的 [MobileFormFactor](/javascript/api/manifest/mobileformfactor) 一起提交至应用商店，则需要同意我们添加针对 iOS 上的外接程序的开发人员附录，并且必须提交你的 Apple 开发人员 ID 以进行验证。

- 最后，清单将需要声明 `MobileFormFactor`，并包含正确的[控件](/javascript/api/manifest/control)和[图标大小](/javascript/api/manifest/icon)类型。

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>适用于移动外接程序的优秀方案应具备哪些特点？

请记住，电话上 Outlook 会话的平均长度要比在 PC 上短得多。这意味着外接程序必须快速运行，且方案必须允许用户进入、退出，并继续处理他们的电子邮件工作流。

以下是在 Outlook Mobile 中可用的方案示例。

- 外接程序为 Outlook 带来了有价值的信息，帮助用户会审他们的电子邮件并进行适当地响应。示例：可让用户查看客户信息并共享相应信息的 CRM 外接程序。

- 外接程序通过将信息保存到跟踪、协作或类似系统，为用户的电子邮件内容增加价值。示例：允许用户将电子邮件转化为任务项以供项目跟踪，或转化为支持团队的帮助票证的外接程序。

**从 iOS 上的电子邮件创建 Trello 卡片的用户交互示例**

![显示用户与 iOS 上的 Outlook Mobile 加载项交互的动画 GIF。](../images/outlook-mobile-addin-interaction.gif)

<br/>

**从 Android 上的电子邮件创建 Trello 卡片的用户交互示例**

![显示用户与 Android 上的 Outlook Mobile 加载项交互的动画 GIF。](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>在移动电话上测试外接程序

若要在 Outlook Mobile 上测试加载项，请首先将 [加载项旁加载](sideload-outlook-add-ins-for-testing.md) 到 Web、Windows 或 Mac 上的 Microsoft 365 或 Outlook.com 帐户。 确保清单格式正确，以包含 `MobileFormFactor` 或不会在移动版 Outlook 客户端中加载。

在加载项正常运行后，请务必在不同尺寸的屏幕（包括电话和平板电脑）上测试加载项。应确保加载项符合与对比度、字号和颜色有关的辅助功能准则，并且还适用于屏幕阅读器（如 iOS 上的 VoiceOver 或 Android 上的 TalkBack）。

在移动设备上进行故障排除可能很困难，因为你可能没有习惯的工具。 但是，在 iOS 上进行故障排除的一个选项是使用 Fiddler (查看 [本教程，了解如何将其与 iOS 设备) 配合使用](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices) 。

> [!NOTE]
> iPhone 和 Android 智能手机上的新式Outlook 网页版不再需要或可用于测试 Outlook 加载项。此外，在 Outlook on Android、iOS 和具有本地 Exchange 帐户的新式移动 Web 中不支持外接程序。 使用具有经典Outlook 网页版的本地 Exchange 帐户时，某些 iOS 设备仍支持加载项。 有关支持的设备的信息，请参阅[运行 Office 加载项的要求](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet)。

## <a name="next-steps"></a>后续步骤

了解如何：

- [向外接程序的清单添加移动支持](add-mobile-support.md)。
- [为外接程序设计出色的移动体验](outlook-addin-design.md)。
- [从外接程序获取访问令牌并调用 Outlook REST API](use-rest-api.md)。
