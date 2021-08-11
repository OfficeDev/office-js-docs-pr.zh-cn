---
title: 基于事件的加载项的 AppSource Outlook选项
description: 了解可用于实现基于事件的激活Outlook加载项的 AppSource 一览选项。
ms.topic: article
ms.date: 08/05/2021
localization_priority: Normal
ms.openlocfilehash: cbc4f43340b5dba4c10c5cf9362c3c6104289ea6ba32a46fb7df758494e27b64
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098359"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a>基于事件的加载项的 AppSource Outlook选项

目前，外接程序必须由组织的管理员部署，以便最终用户能够访问基于事件的功能。 如果最终用户直接从 AppSource 获取了加载项，我们将限制基于事件的激活。 例如，如果 Contoso 外接程序包括扩展点，节点下至少定义了一个扩展点，则只有当其组织的管理员为最终用户安装了外接程序时，才自动调用外接程序。否则，外接程序的自动调用将被阻止。 `LaunchEvent` `LaunchEvent Type` `LaunchEvents` 请参阅示例外接程序清单中的以下摘录。

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

最终用户或管理员可以通过 AppSource 或应用内应用商店获取Office加载项。 如果您的外接程序的主要方案或工作流需要基于事件的激活，您可能需要限制可用于管理员部署的外接程序。 若要启用此限制，我们可以提供测试代码 URL。 由于有航班代码，只有具有这些特殊 URL 的最终用户才能访问列表。 下面是一个示例 URL。

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

为加载项启用外部测试代码后，用户和管理员无法按加载项在 AppSource 或应用内 Office 应用商店中的名称显式搜索加载项。 作为外接程序创建者，你可以与外接程序部署的组织管理员私人共享这些测试代码。

> [!NOTE]
> 虽然最终用户可以使用测试代码安装外接程序，但外接程序不包括基于事件的激活。

## <a name="specify-a-flight-code"></a>指定航班代码

若要指定外接程序的运行代码，在发布外接程序时，在认证说明中共享该信息。 _**重要提示**：_ 航班代码区分大小写。

![Screenshot showing example request for flight code in Notes for certification screen during publishing process.](../images/outlook-publish-notes-for-certification-1.png)

## <a name="deploy-add-in-with-flight-code"></a>使用测试代码部署外接程序

设置测试代码后，你将从应用认证团队收到 URL。 然后，你可以与管理员私人共享 URL。

若要部署外接程序，管理员可以使用以下步骤。

- Sign in to admin.microsoft.com or AppSource.com with your Microsoft 365 admin account. 如果加载项启用了单一登录 (SSO) ，需要全局管理员凭据。
- 将测试版本代码 URL 打开到 Web 浏览器中。
- 在外接程序列表页面上，选择"**现在获取"。** 应重定向到集成应用门户。

## <a name="unrestricted-appsource-listing"></a>无限制 AppSource 一览

如果你的加载项未对关键方案（即 (无需自动调用) 即可正常使用基于事件的激活，请考虑在没有任何特殊外部测试代码的情况下在 AppSource 中列出加载项。 如果最终用户从 AppSource 获取加载项，则用户不会进行自动激活。 但是，他们可以使用外接程序的其他组件，如任务窗格或无 UI 命令。

> [!IMPORTANT]
> 这是一个临时限制。 将来，我们计划为直接获取外接程序的最终用户启用基于事件的外接程序激活。

## <a name="update-existing-add-ins-to-include-event-based-activation"></a>更新现有外接程序以包含基于事件的激活

你可以更新现有加载项以包含基于事件的激活，然后重新提交它进行验证，并决定是需要受限还是不受限制的 AppSource 一览。

更新后的加载项获得批准后，之前部署了加载项的组织管理员将在管理中心的"集成应用"部分收到更新消息。  该消息会向管理员建议基于事件的激活更改。 管理员接受更改后，更新将部署到最终用户。

!["集成应用"屏幕上的应用更新通知屏幕截图。](../images/outlook-deploy-update-notification.png)

对于自己安装加载项的最终用户，即使加载项已更新，基于事件的激活功能也不起作用。

## <a name="admin-consent-for-installing-event-based-add-ins"></a>管理员同意安装基于事件的加载项

只要从"集成应用"屏幕部署基于事件的加载项，管理员就会在部署向导中了解有关加载项基于事件的激活功能的详细信息。 详细信息显示在" **应用程序权限和功能"部分** 。 管理员应看到加载项可以自动激活的所有事件。

![部署新应用时"接受权限请求"屏幕的屏幕截图。](../images/outlook-deploy-accept-permissions-requests.png)

同样，当现有加载项更新为基于事件的功能时，管理员在加载项上会看到"更新挂起"状态。 只有在管理员同意"应用权限和功能"部分中介绍的更改（包括加载项可自动激活的事件集）时，才部署更新的加载项。

每次向加载项添加新内容时，管理员都会在管理门户中看到更新流，并 `LaunchEvent Type` 需要同意其他事件。

![部署更新后的应用时"更新"流的屏幕截图。](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a>另请参阅

- [配置Outlook加载项进行基于事件的激活](autolaunch.md)
