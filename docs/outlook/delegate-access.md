---
title: 在 Outlook 外接程序中启用共享文件夹和共享邮箱方案
description: 讨论如何配置对共享文件夹 (a.k.a 的外接程序支持。 委托访问) 和共享邮箱。
ms.date: 09/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: bae8a0f8cd63eed5feea7460e57ecfc212a06d61
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674664"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>在 Outlook 外接程序中启用共享文件夹和共享邮箱方案

本文介绍如何在 Outlook 外接程序的 [预览](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes)) 方案（包括 Office JavaScript API 支持的权限）中启用共享文件夹 (也称为委托访问) 和共享邮箱 (。

## <a name="supported-clients-and-platforms"></a>支持的客户端和平台

下表显示了此功能支持的客户端-服务器组合，包括在适用的情况下所需的最低累积更新。 不支持排除的组合。

| Client | Exchange Online | 本地 Exchange 2019<br> (累积更新 1 或更高版本)  | 本地 Exchange 2016<br> (累积更新 6 或更高版本)  | 本地 Exchange 2013 |
|---|:---:|:---:|:---:|:---:|
|Windows：<br>版本 1910 (内部版本 12130.20272) 或更高版本|是|是\*|是\*|是\*|
|Mac：<br>内部版本 16.47 或更高版本|是|是|是|是|
|Web 浏览器：<br>新式 Outlook UI|是|不适用|不适用|不适用|
|Web 浏览器：<br>经典 Outlook UI|不适用|否|否|否|

> [!NOTE]
> \* 从版本 2206 开始，可在本地 Exchange 环境中支持此功能， (当前频道版本 15330.20000) 和版本 2207 (每月企业频道版本 15427.20000) 。

> [!IMPORTANT]
> 要求 [集 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) (中引入了对此功能的支持，有关详细信息，请参阅 [客户端和平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)) 。 但是，请注意，该功能的支持矩阵是要求集的超集。

## <a name="supported-setups"></a>支持的设置

以下各节介绍共享邮箱的受支持配置 (现在以预览版) 和共享文件夹提供。 在其他配置中，功能 API 可能无法按预期工作。 选择要了解如何配置的平台。

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>共享文件夹

邮箱所有者必须首先 [提供对委托的访问权限](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)。 然后，委托必须按照“将其他人的邮箱添加到你的个人资料”一文中所述的说明进行操作，该部分将 [管理其他人的邮件和日历项目](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5)。

#### <a name="shared-mailboxes-preview"></a>共享邮箱 (预览) 

Exchange 服务器管理员可以创建和管理共享邮箱，以便用户集访问。 [支持Exchange Online](/exchange/collaboration-exo/shared-mailboxes)和[本地 Exchange 环境](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes)。

默认情况下，将启用名为“自动映射”的Exchange Server功能，这意味着在 Outlook 关闭并重新打开后，[共享邮箱应自动显示](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)在用户的 Outlook 应用中。 但是，如果管理员关闭自动映射，则用户必须按照“将共享邮箱添加到 Outlook”一文中所述的手动步骤进行 [操作，并在 Outlook 中使用共享邮箱](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd)。

> [!WARNING]
> **请勿** 使用密码登录到共享邮箱。 在这种情况下，功能 API 将不起作用。

### <a name="web-browser---modern-outlook"></a>[Web 浏览器 - 新式 Outlook](#tab/modern)

#### <a name="shared-folders"></a>共享文件夹

邮箱所有者必须首先通过更新邮箱文件夹权 [限来提供对委托的访问权限](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) 。 然后，委托必须按照“在Outlook Web App中将其他人的邮箱添加到文件夹列表”一文中所述[的说明访问其他人的邮箱](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081)。

#### <a name="shared-mailboxes"></a>共享邮箱

新式Outlook 网页版目前不支持 Outlook 加载项中的共享邮箱方案。

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>共享邮箱 (预览) 

邮件和日历与委托或共享邮箱用户共享。 在邮件和约会读取和撰写模式下，外接程序可供委托或用户使用。

#### <a name="shared-folders"></a>共享文件夹

如果 **收件箱** 文件夹与委托共享，则加载项在消息读取模式下可供委托使用。

如果 **Drafts** 文件夹也与委托共享，则加载项在撰写模式下可用。

#### <a name="local-shared-calendar-new-model"></a>本地共享日历 (新模型) 

如果日历所有者与委托显式共享其日历 (整个邮箱可能无法共享) ，则在约会读取和撰写模式下，外接程序可供委托使用。

#### <a name="remote-shared-calendar-previous-model"></a>远程共享日历 (以前的模型) 

例如，如果日历所有者授予了对其日历的广泛访问权限 (使其可编辑到特定的 DL 或整个组织) ，则用户可能具有间接或隐式权限，并且在约会读取和撰写模式下，这些用户可以使用加载项。

---

若要详细了解加载项在一般情况下执行和不激活的位置，请参阅 Outlook 外接程序概述页的“外接程序 [”部分可用的邮箱项](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) 。

## <a name="supported-permissions"></a>支持的权限

下表描述了 Office JavaScript API 为委托和共享邮箱用户支持的权限。

|权限|值|说明|
|---|---:|---|
|读取|1 (000001) |可以读取项目。|
|写入|2 (000010) |可以创建项。|
|DeleteOwn|4 (000100) |只能删除他们创建的项。|
|DeleteAll|8 (001000) |可以删除任何项。|
|EditOwn|16 (010000) |只能编辑他们创建的项。|
|EditAll|32 (100000) |可以编辑任何项。|

> [!NOTE]
> 目前，API 支持获取现有权限，但不支持设置权限。

[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) 对象是使用位掩码来指示权限实现的。 位掩码中的每个位置都表示特定权限，如果设置为 `1` ，则用户具有相应的权限。 例如，如果右侧的第二位是 `1`，则用户具有 **“写入** ”权限。 本文稍后将介绍如何在 [“以委托身份执行操作或共享邮箱用户](#perform-an-operation-as-delegate-or-shared-mailbox-user) ”部分中查看特定权限的示例。

## <a name="sync-across-shared-folder-clients"></a>跨共享文件夹客户端同步

委托对所有者邮箱的更新通常会立即跨邮箱同步。

但是，如果使用 REST 或 Exchange Web Services (EWS) 操作在项上设置扩展属性，则此类更改可能需要几个小时才能同步。建议改用 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象和相关 API 以避免此类延迟。 若要了解详细信息，请参阅“在 Outlook 加载项中获取和设置元数据”一文的 [自定义属性部分](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) 。

> [!IMPORTANT]
> 在委托方案中，不能将 EWS 与 office.js API 当前提供的令牌配合使用。

## <a name="configure-the-manifest"></a>配置清单

若要在外接程序中启用共享文件夹和共享邮箱方案，必须将 [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) 元素设置为 `true` 父元素 `DesktopFormFactor`下的清单中。 目前，不支持其他外形因素。

若要支持来自委托的 REST 调用，请将清单中的 [权限](/javascript/api/manifest/permissions) 节点设置为 `ReadWriteMailbox`。

以下示例显示 `SupportsSharedFolders` 在清单的某个部分中设置 `true` 的元素。

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>以委托或共享邮箱用户身份执行操作

可以通过调用 [item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法，在 Compose 或 Read 模式下获取项目的共享属性。 这会返回一个 [SharedProperties](/javascript/api/outlook/office.sharedproperties) 对象，该对象当前提供用户的权限、所有者的电子邮件地址、REST API 的基 URL 和目标邮箱。

以下示例演示如何获取邮件或约会的共享属性、检查委托或共享邮箱用户是否具有 **写** 入权限，以及进行 REST 调用。

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> 作为委托，可以使用 REST [获取附加到 Outlook 项目或组帖子的 Outlook 消息的内容](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>处理对共享项和非共享项调用 REST

如果要对某个项调用 REST 操作，无论是否共享该项目，都可以使用 `getSharedPropertiesAsync` API 来确定项目是否共享。 之后，可以使用相应的对象为操作构造 REST URL。

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>限制

根据加载项的方案，在处理共享文件夹或共享邮箱情况时，需要考虑一些限制。

### <a name="message-compose-mode"></a>消息撰写模式

在消息撰写模式下，除非满足以下条件，否则在 Outlook 网页版 或 Windows 上不支持 [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1))。

a. **委托访问/共享文件夹**

1. 邮箱所有者将启动邮件。 这可以是新消息、答复或转发。
1. 保存消息，然后将其从自己的 **Drafts** 文件夹移动到与委托共享的文件夹。
1. 委托从共享文件夹打开草稿，然后继续撰写。

b. **共享邮箱 (仅适用于 Windows 上的 Outlook)**

1. 共享邮箱用户启动邮件。 这可以是新消息、答复或转发。
1. 保存邮件，然后将其从自己的 **Drafts** 文件夹移动到共享邮箱中的文件夹。
1. 另一个共享邮箱用户从共享邮箱打开草稿，然后继续撰写。

消息现在位于共享上下文中，支持这些共享方案的加载项可以获取项目的共享属性。 发送邮件后，通常会在发件人的“ **已发送邮件** ”文件夹中找到它。

### <a name="rest-and-ews"></a>REST 和 EWS

外接程序可以使用 REST，并且必须设置外接程序的权限，以便 `ReadWriteMailbox` 根据需要启用对所有者邮箱或共享邮箱的 REST 访问。 不支持 EWS。

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>从地址列表中隐藏的用户或共享邮箱

如果管理员将用户或共享邮箱地址从地址列表（如全局地址列表 (GAL) ）中隐藏，则邮箱报 `Office.context.mailbox.item` 表中打开的受影响邮件项目为 null。 例如，如果用户在从 GAL 隐藏的共享邮箱中打开邮件项， `Office.context.mailbox.item` 则表示该邮件项为 null。

## <a name="see-also"></a>另请参阅

- [允许其他人管理你的邮件和日历](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Microsoft 365 中的日历共享](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [将共享邮箱添加到 Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [如何对清单元素进行排序](../develop/manifest-element-ordering.md)
- [掩码 (计算) ](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript 按位运算符](https://www.w3schools.com/js/js_bitwise.asp)
