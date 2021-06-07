---
title: 在加载项中启用Outlook访问方案
description: 简要介绍委派访问权限并讨论如何配置外接程序支持。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 256c37087b10eaf9c8025e19a4990852f9550458
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783489"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>在加载项中启用Outlook访问方案

邮箱所有者可以使用委派访问功能 [允许其他人管理他们的邮件和日历](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。 本文指定 JavaScript API 支持Office委派权限，并介绍如何在加载项中启用Outlook方案。

> [!IMPORTANT]
> 代理访问当前在 Android 和 iOS Outlook不可用。 此外，此功能当前不适用于 Web 上的[](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)Outlook组共享邮箱。 将来可能会提供此功能。
>
> 要求集 1.8 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="supported-permissions-for-delegate-access"></a>委派访问权限支持的权限

下表介绍了 JavaScript API 支持Office委派权限。

|权限|值|说明|
|---|---:|---|
|读取|1 (0000001) |可读取项目。|
|写入|2 (000010) |可以创建项目。|
|DeleteOwn|4 (000100) |只能删除他们创建的项。|
|DeleteAll|8 (001000) |可以删除任何项目。|
|EditOwn|16 (010000) |只能编辑他们创建的项。|
|EditAll|32 (1000000) |可以编辑任何项目。|

> [!NOTE]
> 目前，API 支持获取现有委派权限，但无法设置委派权限。

使用位掩码实现 [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) 对象以指示代理的权限。 位掩码中的每个位置表示特定权限，如果设置为 ，则代理 `1` 具有各自的权限。 例如，如果右边的第二位是 `1` ，则委托具有 **写入** 权限。 您可以在本文稍后的以委派方式执行操作一节中查看如何检查特定[](#perform-an-operation-as-delegate)权限的示例。

## <a name="sync-across-mailbox-clients"></a>跨邮箱客户端同步

代理对所有者邮箱的更新通常会立即跨邮箱同步。

但是，如果使用 REST 或 Exchange Web (EWS) 操作来设置项目的扩展属性，则此类更改可能需要几个小时才能同步。我们建议你改为使用[CustomProperties](/javascript/api/outlook/office.customproperties)对象和相关 API 以避免此类延迟。 若要了解更多信息，请参阅"[](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)在加载项中获取和设置Outlook元数据"一文的自定义属性部分。

> [!IMPORTANT]
> 在委派方案中，不能将 EWS 与当前由 office.js API 提供的令牌一起使用。

## <a name="configure-the-manifest"></a>配置清单

若要在加载项中启用委派访问方案，必须在父元素 下的清单中将 [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` 元素设置为 `DesktopFormFactor` 。 目前，不支持其他外形因素。

若要支持从代理进行 REST 调用，将清单 [中的"权限"](../reference/manifest/permissions.md) 节点设置为 `ReadWriteMailbox` 。

以下示例显示清单 `SupportsSharedFolders` 的一节中设置为 `true` 的 元素。

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

## <a name="perform-an-operation-as-delegate"></a>以委派方式执行操作

可以通过调用 [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法在撰写或阅读模式下获取项目的共享属性。 这将返回 [一个 SharedProperties](/javascript/api/outlook/office.sharedproperties) 对象，该对象当前提供代理的权限、所有者的电子邮件地址、REST API 的基本 URL 和目标邮箱。

以下示例演示如何获取邮件或约会的共享属性、检查代理是否具有 **写入** 权限以及进行 REST 调用。

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

如果要对项目调用 REST 操作（无论该项是否共享）都可以使用 API 来确定 `getSharedPropertiesAsync` 该项目是否共享。 然后，您可以使用适当的对象构造该操作的 REST URL。

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

根据外接程序的方案，在处理委托情况时需要考虑一些限制。

### <a name="rest-and-ews"></a>REST 和 EWS

您的外接程序可以使用 REST，但不能使用 EWS，并且外接程序的权限必须设置为，才能启用对所有者邮箱的 `ReadWriteMailbox` REST 访问。

### <a name="message-compose-mode"></a>邮件撰写模式

在邮件撰写模式下[，getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_)在 web 或 Outlook Windows 不支持，除非满足以下条件。

1. 所有者至少与代理共享一个邮箱文件夹。
1. 代理在共享文件夹中草稿邮件。

    示例：

    - 代理答复或转发共享文件夹中的电子邮件。
    - 然后，代理保存草稿邮件，然后从其自己的 **"草稿"** 文件夹移动到共享文件夹。 代理从共享文件夹打开草稿，然后继续撰写。

邮件发送后，通常会在代理的"已发送项目" **文件夹中找到** 。

## <a name="see-also"></a>另请参阅

- [允许其他人管理邮件和日历](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [日历中的日历Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [如何对清单元素排序](../develop/manifest-element-ordering.md)
- [计算 (的) ](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript 位运算符](https://www.w3schools.com/js/js_bitwise.asp)