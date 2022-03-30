---
title: 在加载项中启用共享文件夹Outlook邮箱方案
description: 讨论如何为共享文件夹配置外接程序支持 (。。。 委派访问) 和共享邮箱。
ms.date: 10/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8ff71ad12fc3c0488c8c73040b125a1ae4674d88
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496927"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>在加载项中启用共享文件夹Outlook邮箱方案

本文介绍如何在 Outlook 外接程序的[预览) 方案中](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes)启用共享文件夹 (也称为委派访问) 和共享邮箱 (，包括 Office JavaScript API 支持哪些权限。

## <a name="supported-clients-and-platforms"></a>支持的客户端和平台

下表显示了此功能支持的客户端-服务器组合，包括所需的最低累积更新（如果适用）。 不支持排除的组合。

| Client | Exchange Online | Exchange 2019 本地部署<br> (累积更新 1 或更高版本)  | Exchange 2016 本地部署<br> (累积更新 6 或更高版本)  | Exchange 2013 本地部署 |
|---|:---:|:---:|:---:|:---:|
|Windows：<br>版本 1910 (版本 12130.20272) 或更高版本|是|否|否|否|
|Mac：<br>内部版本 16.47 或更高版本|是|是|是|是|
|Web 浏览器：<br>新式 Outlook UI|是|不适用|不适用|不适用|
|Web 浏览器：<br>经典Outlook UI|不适用|否|否|否|

> [!IMPORTANT]
> 要求集 [1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) 中引入了对此功能 (有关详细信息，请参阅客户端和 [平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)) 。 但是，请注意，功能的支持矩阵是要求集的超集。

## <a name="supported-setups"></a>支持的安装程序

以下各节介绍共享邮箱和共享文件夹 (预览) 的配置。 在其他配置中，功能 API 可能无法如预期工作。 选择要了解如何配置的平台。

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>共享文件夹

邮箱所有者必须先 [向代理提供访问权限](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)。 然后，代理人必须遵循管理其他人的邮件和日历项目一文的"将其他人的邮箱添加到你的配置文件"部分中 [概述的说明](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5)。

#### <a name="shared-mailboxes-preview"></a>共享邮箱 (预览) 

Exchange管理员可以为要访问的用户集创建和管理共享邮箱。 目前，[Exchange Online](/exchange/collaboration-exo/shared-mailboxes)是此功能唯一受支持的服务器版本。

默认情况下Exchange Server自动映射"功能是启用的，这意味着共享邮箱随后应在关闭并重新打开共享邮箱后自动显示在[](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)用户的 Outlook Outlook 应用中。 但是，如果管理员关闭自动映射，用户必须遵循在 Outlook 中打开和使用共享邮箱一文的"将共享邮箱添加到 Outlook"一节中概述[的手动步骤](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd)。

> [!WARNING]
> **请勿使用** 密码登录共享邮箱。 在这种情况下，功能 API 将不起作用。

### <a name="web-browser---modern-outlook"></a>[Web 浏览器 - 新式 Outlook](#tab/modern)

#### <a name="shared-folders"></a>共享文件夹

邮箱所有者必须先 [通过更新邮箱文件夹](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) 权限来向代理提供访问权限。 然后，代理人必须遵循文章访问其他人的邮箱"将其他人的邮箱添加到 Outlook Web App 中的文件夹列表"部分中概述[的说明](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081)。

#### <a name="shared-mailboxes-preview"></a>共享邮箱 (预览) 

Exchange管理员可以为要访问的用户集创建和管理共享邮箱。 目前，[Exchange Online](/exchange/collaboration-exo/shared-mailboxes)是此功能唯一受支持的服务器版本。

获得访问权限后，共享邮箱用户必须遵循在"在邮箱中打开和使用共享邮箱"一文的"添加共享邮箱，以便它显示在主邮箱[下"一节中概述Outlook 网页版](https://support.microsoft.com/office/98b5a90d-4e38-415d-a030-f09a4cd28207)。

> [!WARNING]
> 请勿 **使用** "打开另一个邮箱"等其他选项。 然后，功能 API 可能无法正常运行。

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>共享邮箱 (预览) 

邮件和日历与代理或共享邮箱用户共享。 在邮件和约会阅读和撰写模式下，代理或用户可以使用外接程序。

#### <a name="shared-folders"></a>共享文件夹

如果 **"收件箱** "文件夹与代理共享，则外接程序在邮件阅读模式下对代理可用。

如果 **"草稿** "文件夹也与代理共享，则外接程序在撰写模式下可用。

#### <a name="local-shared-calendar-new-model"></a>本地共享日历 (模型) 

如果日历所有者与代理显式共享日历 (整个邮箱可能不会共享) ，则代理可以在约会阅读和撰写模式下使用外接程序。

#### <a name="remote-shared-calendar-previous-model"></a>远程共享日历 (模型) 

例如，如果日历所有者授予了对日历 (的广泛访问权限，使日历所有者能够编辑特定的 DL 或整个组织) ，则用户随后可能拥有间接或隐式权限，并且这些用户在约会阅读和撰写模式下可以使用外接程序。

---

若要了解有关外接程序在一般情况下是在哪里激活和不激活的更多信息，请参阅 Outlook 外接程序概述[](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)页的"可用于外接程序的邮箱项目"部分。

## <a name="supported-permissions"></a>支持的权限

下表介绍了 JavaScript API 支持Office和共享邮箱用户的权限。

|权限|值|Description|
|---|---:|---|
|阅读|1 (000001) |可读取项目。|
|写入|2 (000010) |可以创建项目。|
|DeleteOwn|4 (000100) |只能删除他们创建的项。|
|DeleteAll|8 (001000) |可以删除任何项目。|
|EditOwn|16 (010000) |只能编辑他们创建的项。|
|EditAll|32 (1000000) |可以编辑任何项目。|

> [!NOTE]
> 目前，API 支持获取现有权限，但无法设置权限。

使用位掩码来指示权限实现 [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) 对象。 位掩码中的每个位置表示特定权限 `1` ，如果设置为 ，则用户具有各自的权限。 例如，如果右边的第二位是 `1`，则用户具有 **写入** 权限。 您可以在本文稍后的以委派或共享邮箱用户角色执行操作部分查看[](#perform-an-operation-as-delegate-or-shared-mailbox-user)如何检查特定权限的示例。

## <a name="sync-across-shared-folder-clients"></a>跨共享文件夹客户端同步

代理对所有者邮箱的更新通常会立即跨邮箱同步。

但是，如果使用 REST 或 Exchange Web (EWS) 操作来设置项目的扩展属性，则此类更改可能需要几个小时才能同步。我们建议你改为使用 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象和相关 API 以避免此类延迟。 若要了解更多信息，请参阅"[](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)在加载项中获取和设置Outlook元数据"一文的自定义属性部分。

> [!IMPORTANT]
> 在委派方案中，不能将 EWS 与当前由 office.js API 提供的令牌一同使用。

## <a name="configure-the-manifest"></a>配置清单

若要在加载项中启用共享文件夹和共享邮箱方案，必须在父元素 下的清单中将 [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) `true` 元素设置为 `DesktopFormFactor`。 目前，不支持其他外形因素。

若要支持从代理进行 REST 调用，将清单中的 ["权限](/javascript/api/manifest/permissions) "节点设置为 `ReadWriteMailbox`。

以下示例显示清单 `SupportsSharedFolders` 的一节中 `true` 设置为 的 元素。

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>以委派邮箱用户或共享邮箱用户模式执行操作

可以通过调用 [item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法在撰写或阅读模式下获取项目的共享属性。 这将返回 [一个 SharedProperties](/javascript/api/outlook/office.sharedproperties) 对象，该对象当前提供用户的权限、所有者的电子邮件地址、REST API 的基本 URL 和目标邮箱。

以下示例演示如何获取邮件或约会的共享属性、检查代理或共享邮箱用户是否具有写入权限以及进行 REST 调用。

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
> 作为代理，您可以使用 REST 获取附加到项目或组帖子Outlook邮件Outlook[内容](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>处理对共享项和非共享项的调用 REST

如果要对项目调用 REST `getSharedPropertiesAsync` 操作（无论该项是否共享）都可以使用 API 来确定该项目是否共享。 然后，您可以使用适当的对象构造该操作的 REST URL。

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

根据外接程序的方案，在处理共享文件夹或共享邮箱情况时需要考虑一些限制。

### <a name="message-compose-mode"></a>邮件撰写模式

在邮件撰写模式下，[getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) 在 Outlook 网页版 或 Windows都不受支持，除非满足以下条件。

a. **委派访问权限/共享文件夹**

1. 邮箱所有者启动一封邮件。 这可以是新邮件、回复或转发。
1. 他们保存邮件，然后将邮件从自己的 **"草稿** "文件夹移动到与代理共享的文件夹。
1. 代理从共享文件夹打开草稿，然后继续撰写。

b. **共享邮箱**

1. 共享邮箱用户启动邮件。 这可以是新邮件、回复或转发。
1. 他们保存邮件，然后将邮件从自己的 **"草稿** "文件夹移动到共享邮箱中的文件夹。
1. 另一个共享邮箱用户从共享邮箱打开草稿，然后继续撰写。

消息现在位于共享上下文中，并且支持这些共享方案的外接程序可以获取项目的共享属性。 邮件发送后，通常会在发件人的"已发送邮件" **文件夹中找到** 该邮件。

### <a name="rest-and-ews"></a>REST 和 EWS

您的外接程序可以使用 REST `ReadWriteMailbox` ，并且外接程序的权限必须设置为启用对所有者邮箱或共享邮箱的 REST 访问（如果适用）。 不支持 EWS。

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>从地址列表中隐藏的用户或共享邮箱

如果管理员从地址列表中隐藏用户或共享邮箱地址，如全局地址列表 (GAL) `Office.context.mailbox.item` ，则邮箱报告中打开的受影响的邮件项目为 null。 例如，如果用户在共享邮箱中打开一个在 GAL `Office.context.mailbox.item` 中隐藏的邮件项目，则代表该邮件项目为空。

## <a name="see-also"></a>另请参阅

- [允许其他人管理邮件和日历](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [日历中的日历Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [将共享邮箱添加到Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [如何对清单元素排序](../develop/manifest-element-ordering.md)
- [计算 (的) ](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript 位运算符](https://www.w3schools.com/js/js_bitwise.asp)