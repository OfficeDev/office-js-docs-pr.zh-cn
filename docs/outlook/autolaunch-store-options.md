---
title: 基于事件的 Outlook 外接程序的 AppSource 列表选项
description: 了解适用于实现基于事件的激活的 Outlook 外接程序的 AppSource 列表选项。
ms.topic: article
ms.date: 09/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: cf99959b31bae665df250941abf88405906acb5c
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674716"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a>基于事件的 Outlook 外接程序的 AppSource 列表选项

加载项必须由组织的管理员部署，最终用户才能访问基于事件的功能。 如果最终用户直接从 [AppSource](https://appsource.microsoft.com) 获取加载项，则限制基于事件的激活。 例如，如果 Contoso 加载项包含`LaunchEvent`在节点下`LaunchEvents`至少定义`LaunchEvent Type`了一个扩展点，则仅当加载项由其组织的管理员为最终用户安装外接程序时才发生自动调用。否则，将阻止自动调用加载项。 请参阅示例加载项清单中的以下摘录。

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

最终用户或管理员可以通过 AppSource 或应用内 Office 应用商店获取加载项。 如果外接程序的主要方案或工作流需要基于事件的激活，则可能需要限制可用于管理员部署的外接程序。 若要启用此限制，我们可以提供外部测试码 URL。 由于外部测试版代码，只有具有这些特殊 URL 的最终用户才能访问列表。 下面是一个示例 URL。

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

启用外部测试版代码时，用户和管理员无法在 AppSource 或应用内 Office 应用商店中按其名称显式搜索加载项。 作为加载项创建者，可以私下与组织管理员共享这些外部测试版代码以进行外接程序部署。

> [!NOTE]
> 虽然最终用户可以使用外部测试版代码安装外接程序，但外接程序不会包含基于事件的激活。

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="specify-a-flight-code"></a>指定外部测试版代码

若要指定外接程序的外部测试版代码，请在发布加载项时在 **“备注”中共享代码以进行认证** 。 **重要** 说明：外部测试版代码区分大小写。

![发布过程中用于认证屏幕的备注中外部测试代码的示例请求。](../images/outlook-publish-notes-for-certification.png)

## <a name="deploy-add-in-with-flight-code"></a>使用外部测试版代码部署加载项

设置外部测试版代码后，你将从应用认证团队收到 URL。 然后，可以私下与管理员共享 URL。

若要部署外接程序，管理员可以使用以下步骤。

- 使用 Microsoft 365 管理员帐户登录到 admin.microsoft.com 或 AppSource.com。 如果加载项已启用单一登录 (SSO) ，则需要全局管理员凭据。
- 将外部测试版代码 URL 打开到 Web 浏览器中。
- 在加载项列表页上，选择 **“立即获取”。** 应重定向到集成应用门户。

## <a name="unrestricted-appsource-listing"></a>不受限制的 AppSource 列表

如果外接程序不对关键方案使用基于事件的激活， (，则加载项在不自动调用) 的情况下运行良好，请考虑在 AppSource 中列出加载项，而无需任何特殊的外部测试版代码。 如果最终用户从 AppSource 获取加载项，则不会对用户进行自动激活。 但是，他们可以使用外接程序的其他组件，例如任务窗格或函数命令。

> [!IMPORTANT]
> 这是一个临时限制。 将来，我们计划为直接获取加载项的最终用户启用基于事件的加载项激活。

## <a name="update-existing-add-ins-to-include-event-based-activation"></a>更新现有加载项以包括基于事件的激活

可以更新现有外接程序以包含基于事件的激活，然后重新提交它进行验证，并确定是否需要受限或不受限制的 AppSource 列表。

在已更新的加载项获得批准后，以前部署过外接程序的组织管理员将在管理中心的 **“集成应用** ”部分收到更新消息。 该消息向管理员提供有关基于事件的激活更改的建议。 管理员接受更改后，更新将部署到最终用户。

![“集成应用”屏幕上的应用更新通知。](../images/outlook-deploy-update-notification.png)

对于自行安装外接程序的最终用户，即使加载项已更新，基于事件的激活功能也不会工作。

## <a name="admin-consent-for-installing-event-based-add-ins"></a>管理员许可安装基于事件的加载项

每当从 **集成应用** 屏幕部署基于事件的外接程序时，管理员都会在部署向导中获取有关加载项基于事件的激活功能的详细信息。 详细信息显示在 **“应用权限和功能** ”部分。 管理员应看到加载项可以自动激活的所有事件。

![部署新应用时的“接受权限请求”屏幕。](../images/outlook-deploy-accept-permissions-requests.png)

同样，当现有加载项更新为基于事件的功能时，管理员会在加载项上看到“更新挂起”状态。 仅当管理员同意应用 **权限和功能** 部分中所述的更改（包括加载项可以自动激活的事件集）时，才会部署更新的加载项。

每次向外接程序添加任何新 `LaunchEvent Type` 内容时，管理员都会在管理门户中看到更新流，并且需要为其他事件提供许可。

![部署更新后的应用时的“汇报”流。](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a>另请参阅

- [配置 Outlook 外接程序以进行基于事件的激活](autolaunch.md)
