---
title: 适用于 Outlook Mobile 的 Outlook 外接程序
description: 所有 Office 365 商业帐户、Outlook.com 帐户均支持 Outlook 移动外接程序，并且即将提供对 Gmail 帐户的支持。
ms.date: 02/28/2020
localization_priority: Normal
ms.openlocfilehash: 7fc4ac511fe7e101775334cad6d4b000f7dc24ae
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561793"
---
# <a name="add-ins-for-outlook-mobile"></a>适用于 Outlook Mobile 的外接程序

现在，外接程序在 Outlook Mobile 上可用，它们使用适用于其他 Outlook 终结点的相同 API。如果已经生成适用于 Outlook 的外接程序，那么则可以很轻松地在 Outlook Mobile 上使用该外接程序。

所有 Office 365 商业帐户、Outlook.com 帐户均支持 Outlook 移动外接程序，并且即将提供对 Gmail 帐户的支持。

**iOS 版 Outlook 中的任务窗格示例**

![iOS 版 Outlook 中任务窗格的屏幕截图](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Android 版 Outlook 中的任务窗格示例**

![Android 版 Outlook 中任务窗格的屏幕截图](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> 外接程序在移动浏览器中的 Outlook 的新式版本中不起作用。 有关详细信息，请参阅[正在升级移动浏览器上的 Outlook](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)。

## <a name="whats-different-on-mobile"></a>在移动电话上会有什么不同？

- 移动电话尺寸小，需要进行快速交互，这为设计适用于移动电话的加载项带来了挑战。为了确保客户体验的质量，我们正在设置严格的验证标准，声明提供移动支持的加载项必须符合这一标准，以便在 AppSource 中获得批准。
    - 外接程序**必须**遵循 [UI 准则](outlook-addin-design.md)。
    - 外接程序的方案**必须**[能够在移动电话上实现](#what-makes-a-good-scenario-for-mobile-add-ins)。

- 目前仅支持邮件读取。这意味着 `MobileMessageReadCommandSurface` 是在清单的移动电话部分唯一需要声明的 [ExtensionPoint](../reference/manifest/extensionpoint.md)。

- [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API 在移动电话上不受支持，因为移动应用使用 REST API 与服务器进行通信。如果应用后端需要连接到 Exchange 服务器，则可以使用回调令牌进行 REST API 调用。有关详细信息，请参阅[从 Outlook 外接程序使用 Outlook REST API](use-rest-api.md)。

- 如果将外接程序和清单中的 [MobileFormFactor](../reference/manifest/mobileformfactor.md) 一起提交至应用商店，则需要同意我们添加针对 iOS 上的外接程序的开发人员附录，并且必须提交你的 Apple 开发人员 ID 以进行验证。

- 最后，清单将需要声明 `MobileFormFactor`，并包含正确的[控件](../reference/manifest/control.md)和[图标大小](../reference/manifest/icon.md)类型。

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>适用于移动外接程序的优秀方案应具备哪些特点？

请记住，电话上 Outlook 会话的平均长度要比在 PC 上短得多。这意味着外接程序必须快速运行，且方案必须允许用户进入、退出，并继续处理他们的电子邮件工作流。

以下是在 Outlook Mobile 中可用的方案示例。

- 外接程序为 Outlook 带来了有价值的信息，帮助用户会审他们的电子邮件并进行适当地响应。示例：可让用户查看客户信息并共享相应信息的 CRM 外接程序。

- 外接程序通过将信息保存到跟踪、协作或类似系统，为用户的电子邮件内容增加价值。示例：允许用户将电子邮件转化为任务项以供项目跟踪，或转化为支持团队的帮助票证的外接程序。

**从 iOS 上的电子邮件创建 Trello 卡片的用户交互示例**

![显示用户与 iOS 上的 Outlook Mobile 外接程序交互的动态 GIF](../images/outlook-mobile-addin-interaction.gif)

<br/>

**从 Android 上的电子邮件创建 Trello 卡片的用户交互示例**

![显示用户与 Android 上的 Outlook Mobile 外接程序交互的动态 GIF](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>在移动电话上测试外接程序

若要在 Outlook Mobile 上测试加载项，可以将加载项旁加载到 O365 或 Outlook.com 帐户。在 Outlook 网页版中，转到设置齿轮，并选择“**管理集成**”或“**管理加载项**”。在靠近顶部的位置，单击显示的“**单击此处添加自定义加载项**”并上传清单。请确保清单格式正确以包含 `MobileFormFactor`，否则将无法上传。

在加载项正常运行后，请务必在不同尺寸的屏幕（包括电话和平板电脑）上测试加载项。应确保加载项符合与对比度、字号和颜色有关的辅助功能准则，并且还适用于屏幕阅读器（如 iOS 上的 VoiceOver 或 Android 上的 TalkBack）。

在移动电话上进行故障排除可能会比较困难，因为可能你没有习惯使用的工具。进行故障排除的一种选择是[使用 Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md)。或者，如果之前使用过 Fiddler，请查看[本教程中有关在 iOS 设备上使用 Fiddler 的内容](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)。

## <a name="next-steps"></a>后续步骤

了解如何：

- [向外接程序的清单添加移动支持](add-mobile-support.md)。
- [为外接程序设计出色的移动体验](outlook-addin-design.md)。
- [从外接程序获取访问令牌并调用 Outlook REST API](use-rest-api.md)。
