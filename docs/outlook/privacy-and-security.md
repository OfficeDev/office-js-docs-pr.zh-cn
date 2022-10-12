---
title: Outlook 加载项的隐私、权限和安全性
description: 了解如何管理 Outlook 加载项中的隐私、权限和安全性。
ms.date: 10/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 560c9bbdfcde849b66d86e9c000d78f094b3e561
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541247"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Outlook 外接程序的隐私、权限和安全性

最终用户、开发人员和管理员可以使用 Outlook 外接程序的安全模型的分层权限级别来控制隐私和性能。

本文介绍了 Outlook 加载项可以请求的可能权限，并从以下几个角度审视安全模型。

- **AppSource**：加载项完整性

- **最终用户**：隐私和性能问题

- **开发人员**：权限选择和资源使用限制

- **管理员**：设置性能阈值的权限

## <a name="permissions-model"></a>权限模型

Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.

有四个级别的权限。

[!include[Table of Outlook permissions](../includes/outlook-permission-levels-table.md)]

四个级别的权限具有累积性：**读/写邮箱** 权限包括 **读/写项** 权限、**读取项** 权限和 **受限** 权限；**读/写项** 权限包括 **读取项** 权限和 **受限** 权限；**读取项** 权限包括 **受限** 权限。

下图显示了四个级别的权限并说明了每一层提供给最终用户、开发人员和管理员的功能。 有关这些权限的详细信息，请参阅 [最终用户：隐私和性能问题](#end-users-privacy-and-performance-concerns)、[开发人员：权限选择和资源使用限制](#developers-permission-choices-and-resource-usage-limits) 和[了解 Outlook 加载项权限](understanding-outlook-add-in-permissions.md)。

**将四层权限模型与最终用户、开发人员和管理员关联**

![邮件应用架构 v1.1 的四层权限模型图。](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a>AppSource：加载项完整性

[AppSource](https://appsource.microsoft.com) 托管可由最终用户和管理员安装的加载项。 AppSource 强制执行以下措施来维护这些 Outlook 加载项的完整性。

- 要求加载项的主机服务器始终使用安全套接字层 (SSL) 进行通信。

- 要求开发人员在提交加载项时提供身份证明、合约协议和适合的隐私策略。

- 以只读模式存档加载项。

- 支持针对可用加载项的用户审阅系统以推广自我管理的社区。

## <a name="optional-connected-experiences"></a>可选连接体验

最终用户和 IT 管理员可在 Office 桌面和移动客户端中关闭[可选的已连接体验](/deployoffice/privacy/optional-connected-experiences)。 对于 Outlook 外接程序，禁用 **可选连接体验** 设置的影响取决于客户端，但通常意味着不允许用户安装的加载项和 Office 应用商店的访问权限。 组织的 IT 管理员通过[集中部署](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)而部署的加载项仍然可用。

- Windows\*、Mac：未显示 **“获取加载项** ”按钮，因此用户无法再管理其加载项或访问 Office 应用商店。
- Android、iOS：**“获取外接程序”** 对话框仅显示管理员部署的加载项。
- 浏览器：加载项的可用性和对应用商店的访问不受影响，因此用户可以继续[管理其加载项](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce)（包括由管理员部署的加载项）。

  > [!NOTE]
  > \* 对于 Windows，版本 2008 (内部版本 13127.20296) 中提供了对此体验/行为的支持。 如需了解你的版本的更多详情，请参阅 [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) 的更新历史记录页，以及如何[查找 Office 客户端版本和更新频道](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)。

有关常规加载项行为，请参阅 [Office 加载项的隐私和安全性](../concepts/privacy-and-security.md#optional-connected-experiences)。

## <a name="end-users-privacy-and-performance-concerns"></a>最终用户：隐私和性能问题

安全模型通过下列方式解决最终用户的安全、隐私和性能问题。

- 受 Outlook 信息权限管理 (IRM 保护的最终用户消息) 不会与非 Windows 客户端上的 Outlook 加载项交互。

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed. No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.

- 授予“**受限**”权限可允许 Outlook 加载项仅具有对当前项目的有限访问权限。 授予 **读取项** 权限允许 Outlook 外接程序仅访问当前项上的个人可识别信息，例如发件人、收件人姓名和电子邮件地址。

- An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.

- 最终用户可以安装支持上下文相关方案的低信任度 Outlook 外接程序，这不仅对用户具有吸引力，同时还可以最大限度地降低用户的安全风险。

- 已安装 Outlook 外接程序的清单文件在用户电子邮件帐户中受到保护。

- 通过托管 Office 外接程序的服务器传送的数据始终根据安全套接字层 (SSL) 协议进行加密。

- 仅适用于 Outlook 富客户端：Outlook 富客户端监视已安装 Outlook 外接程序的性能，实施管治控制，以及禁用在以下方面超过限制的 Outlook 外接程序。

  - 激活响应时间

  - 激活或重新激活失败次数

  - 内存使用率

  - CPU 使用率  

  Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.

- 无论何时，最终用户都可以验证所安装 Outlook 外接程序请求的权限，在 Exchange 管理中心禁用或随后启用任何 Outlook 外接程序。

## <a name="developers-permission-choices-and-resource-usage-limits"></a>开发人员：权限选择和资源使用限制

安全模型向开发人员提供精细级别的权限以供选择，以及严格的性能准则以供遵循。

### <a name="tiered-permissions-increases-transparency"></a>多层权限将增加透明度

开发人员应按照多层权限模型提供透明度，并解决用户有关哪些加载项可以处理其数据和邮箱的问题，间接促进加载项采用。

- 开发人员根据 Outlook 外接程序应激活的方式、Outlook 外接程序读取或写入项目特定属性的需求，或者创建和发送项目的需求来针对 Outlook 外接程序请求适当级别的权限。

- 如上所述，开发人员在清单中请求权限。

  以下示例请求 XML 清单中的 **读取项** 权限。

  ```XML
  <Permissions>ReadItem</Permissions>
  ```

  以下示例请求 Teams 清单中的 **读取项** 权限 (预览) 。

```json
"authorization": {
  "permissions": {
    "resourceSpecific": [
      ...
      {
        "name": "MailboxItem.Read.User",
        "type": "Delegated"
      },
    ]
  }
},
```

- 如果 Outlook 外接程序在特定类型的 Outlook 项目 (约会或消息) 激活，或者在特定提取的实体上激活， (电话号码、地址、URL) 存在于项目的主题或正文中，则开发人员可以请求 **受限** 权限。 例如，如果在当前邮件的主题或正文中找到一个或多个实体（共三个）- 电话号码、邮寄地址或 URL，以下规则将激活 Outlook 外接程序。

> [!NOTE]
> 使用 Office 外接程序的 Teams 清单的外接程序不支持激活规则 [， (预览版) ](../develop/json-manifest-overview.md)。

  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- 如果 Outlook 外接程序需要读取除默认提取的实体以外的当前项目的属性，或者编写外接程序在当前项上设置的自定义属性，但不需要读取或写入其他项目，或者在用户的邮箱中创建或发送邮件，开发人员应请求 **读** 取项权限。 例如，如果 Outlook 外接程序需要寻找项目主体或正文中的会议建议、任务建议、电子邮件地址或联系人姓名等实体，或者需要使用一个正则表达式来激活，则开发人员应请求“**读取项目**”权限。

- 如果 Outlook 加载项需要向撰写的项目的属性（如收件人姓名、电子邮件地址、正文和主题）写入，或需要添加或删除项目附件，那么开发人员应请求“**读/写项目**”权限。

- 仅在 Outlook 外接程序需要使用 [mailbox.makeEWSRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法执行下列一个或多个操作时，开发人员才请求 **“读/写邮箱”** 权限。

  - 读取或写入邮箱中项目的属性。
  - 创建、读取、写入或发送邮箱中的项目。
  - 创建、读取或写入邮箱文件夹。

### <a name="resource-usage-tuning"></a>资源使用调整

Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.

### <a name="other-measures-to-promote-user-security"></a>提高用户安全性的其他措施

开发人员还应该注意并规划以下内容。

- 开发人员无法在加载项中使用 ActiveX 控件，因为它们不受支持。

- 开发人员应在将 Outlook 加载项提交到 AppSource 时执行以下操作。

  - 生成扩展验证 (EV) SSL 证书作为身份证明。

  - 在支持 SSL 的 Web 服务器上承载其提交的加载项。

  - 生成合规隐私策略。

  - 准备好在提交加载项后签订合约协议。

## <a name="administrators-privileges"></a>管理员：权限

安全模型向管理员提供以下权限和责任。

- 可以阻止最终用户安装任何 Outlook 加载项，包括来自 AppSource 的加载项。

- 可以在 Exchange 管理中心上禁用或启用任何 Outlook 加载项。

- 仅适用于 Windows 版 Outlook：可以通过 GPO 注册表设置覆盖性能阈值设置。

## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全性](../concepts/privacy-and-security.md)
- [Microsoft 365 应用的隐私控制](/deployoffice/privacy/overview-privacy-controls)
- [Outlook 外接程序 API](apis.md)
- [Outlook 外接程序的激活和 JavaScript API 限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
