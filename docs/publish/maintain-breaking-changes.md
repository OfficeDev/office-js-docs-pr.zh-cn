---
title: 维护 Office 加载项
description: 了解我们对兼容性的承诺，以及如何使加载项保持最新。
ms.date: 05/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7f70eab252af516ab8dda591668d48392ce9f04
ms.sourcegitcommit: e63d8e32b25a9987f4a39b92a342a82b37a3404c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/17/2022
ms.locfileid: "65432188"
---
# <a name="maintain-your-office-add-in"></a>维护 Office 加载项

发布加载项后，应通过上游库中的任何重要更改使其保持最新。 修补安全问题对于建立客户信任至关重要。 由于这些更改对已发布的清单没有影响，因此客户无需执行任何操作即可获取最新版本的外接程序。

## <a name="breaking-changes-in-officejs"></a>Office.js的重大更改

Microsoft 365开发人员平台致力于确保加载项的兼容性。 我们努力避免对 API 图面和行为进行重大更改。 但是，在某些情况下，为了安全或可靠性，我们需要进行中断性更新。 在这些极少数情况下，将执行以下步骤，确保加载项的用户不受影响。

- 有关受影响功能和建议更改的公告，请在[Microsoft 365开发人员博客](https://devblogs.microsoft.com/microsoft365dev/)上进行。
- 如果加载项已在 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 中发布，将通过你提供的信息与你联系。
- 如果可能，将通过[消息中心](/microsoft-365/admin/manage/message-center)联系受影响的Microsoft 365租户的管理员 (包括[开发人员租户](https://developer.microsoft.com/microsoft-365/dev-program)) 。 管理员负责联系 AppSource 外部发布的外接程序解决方案提供商。

### <a name="deprecation-policy"></a>弃用策略

可以弃用具有更好替代方法的 API 或工具。 在停用之前至少 24 个月，Microsoft 会尽最大努力声明某些内容已弃用。 同样，对于通常可用的 (GA) 单个 API，Microsoft 会在从 GA 版本将其删除之前至少 24 个月时声明其为弃用产品。

弃用并不一定意味着开发人员将删除该功能或 API 且不可用。 它确实表明，在 24 个月的时间段后，Microsoft 将不再支持 API 或功能。

当 API 被标记为已弃用时，我们强烈建议你尽快迁移到最新版本。 在某些情况下，我们将宣布，在弃用原始 API 后，新应用程序必须在短时间内开始使用新 API。 在这些情况下，仅当前使用已弃用 API 的活动应用程序能够继续使用它们。

> [!IMPORTANT]
> 如果等待时间过长会对加载项或 Microsoft 造成安全风险，则 24 个月的弃用期将会加速。

### <a name="app-assure"></a>应用保证

Microsoft 的 [App Assure](https://www.microsoft.com/fasttrack/microsoft-365/app-assure) 服务履行了 Microsoft 的应用程序兼容性承诺：应用将致力于Windows和Microsoft 365 应用版。 应用保证工程师可以帮助解决你可能遇到的任何问题，无需额外付费。

如果确实遇到应用兼容性问题，应用服务工程师将与你合作，帮助你解决问题。 我们的专家将：

- 帮助你排查和确定根本原因。
- 提供指导，帮助你修正应用程序兼容性问题。
- 请代表你与独立软件供应商 (ISV) ，以修正其应用的某些部分，使其在最新式版本的产品上正常工作。
- 与 Microsoft 产品工程团队合作，修复产品 bug。

若要了解有关应用保证的详细信息，请观看[使用应用保证Microsoft Edge应用：提示和技巧](https://techcommunity.microsoft.com/t5/video-hub/bring-your-apps-to-microsoft-edge-with-app-assure-tips-and/ba-p/2167619)。 若要提交应用与 App Assure 兼容的请求，请完成[Microsoft FastTrack注册表单](https://aka.ms/AppAssureRequest)或向 [achelp@microsoft.com](mailto:achelp@microsoft.com) 发送电子邮件。

## <a name="changes-to-yeoman-templates-and-web-dependencies"></a>对 Yeoman 模板和 Web 依赖项的更改

[用于Office加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)依赖于 Microsoft 和其他部门提供的多个库。 这些库独立于任何Microsoft 365活动进行更新。 在开发、发布和维护外接程序时，使用生成器创建的任何项目都应保持最新。 以下工具可帮助确保项目使用任何依赖库的安全版本。

- [npm审核](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot 和其他GitHub安全功能](https://github.com/features/security)

本指南还适用于从[Office加载项代码示例和其他源提取的示例](https://github.com/OfficeDev/Office-Add-in-samples)副本。

### <a name="officejs-npm-package"></a>office.js NPM 包

[office-js NPM 包](https://www.npmjs.com/package/@microsoft/office-js)是托管在[Office.js内容分发网络 (CDN) ](../develop/understanding-the-javascript-api-for-office.md#accessing-the-office-javascript-api-library)的副本。 它适用于无法直接访问CDN的情况。 NPM 包不用于提供对office.js的版本控制引用。 强烈建议始终使用CDN，以确保使用最新版本的 Office JavaScript API。

## <a name="current-best-practices"></a>当前最佳做法

虽然我们努力保持向后兼容性，但我们建议不断改进的模式和做法。 我们的文档致力于介绍当前的最佳做法。 若要随时了解可能改进现有功能的新功能，请加入我们的每月[Office加载项Community呼叫](../overview/office-add-ins-community-call.md)。

## <a name="community-engagement"></a>Community参与

随着Microsoft 365开发人员平台的更新建议，我们将听取反馈。 请向[Office加载项其他资源](../resources/resources-links-help.md)中列出的频道报告问题、潜在后果或其他问题。
