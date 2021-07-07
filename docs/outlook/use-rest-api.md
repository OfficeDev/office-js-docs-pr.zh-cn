---
title: 从 Outlook 加载项使用 Outlook REST API
description: 了解如何从 Outlook 加载项使用 Outlook REST API 获得访问令牌。
ms.date: 07/06/2021
localization_priority: Normal
ms.openlocfilehash: 9f6642afcfae8efd54c4ade6165aa2a6823e3bd2
ms.sourcegitcommit: 488b26b29c7534e3bbc862b688ed2319cc028f71
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53315146"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>从 Outlook 加载项使用 Outlook REST API

[Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 命名空间提供访问许多邮件和约会的公用字段的权限。但是，在某些方案中，外接程序可能需要访问命名空间未公开的数据。例如，外接程序可能依赖于外部应用设置的自定义属性，或需要搜索用户邮箱中来自同一发件人的邮件。在这些方案中，[Outlook REST API](/outlook/rest) 是推荐的检索信息的方法。

> [!IMPORTANT]
> **已Outlook REST API**
>
> 有关Outlook， (2022 年 11 月将完全停用 REST 终结点，请参阅[2020 年 11](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)月) 。 应迁移现有加载项，以使用 Microsoft [Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph)。 此外，[比较 Microsoft Graph 和 Outlook REST API 终结点](/outlook/rest/compare-graph)。

## <a name="get-an-access-token"></a>获取访问令牌

Outlook REST API 需要 `Authorization` 标头中的持有者令牌。应用通常使用 OAuth2 流检索令牌。不过，加载项也可以使用邮箱要求集 1.5 中引入的新 [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法检索令牌，而无需实现 OAuth2。

通过将 `isRest` 选项设置为 `true`，可以请求获取与 REST API 兼容的令牌。

### <a name="add-in-permissions-and-token-scope"></a>外接程序权限和令牌范围

请务必考虑外接程序通过 REST API 所需要的访问级别。在大多数情况下，由 `getCallbackTokenAsync` 返回的令牌将仅提供对当前项的只读访问权限。即使外接程序在其清单中指定了 `ReadWriteItem` 权限级别也是如此。

如果外接程序需要当前项目或用户邮箱中的其他项目的写权限，则外接程序必须在其清单中指定 `ReadWriteMailbox` 权限级别。在这种情况下，所返回的令牌将包含用户的邮件、事件和联系人的读取/写入访问权限。

### <a name="example"></a>示例

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>获取项 ID

加载项需要针对 REST 正确设置格式的项 ID，才能通过 REST 检索当前项。 这可从 [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性获取，但应进行一些检查，以确保它是针对 REST 正确设置格式的 ID。

- 在 Outlook Mobile 中，由 `Office.context.mailbox.item.itemId` 返回的值是适用于 REST 格式的 ID 并可按原样使用。
- 在其他 Outlook 客户端中，由 `Office.context.mailbox.item.itemId` 返回的值是适用于 EWS 格式的 ID，且必须使用 [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法进行转换。
- 请注意，还必须将附件 ID 转换为带 REST 格式的 ID，才能使用它。 必须转换 ID 的原因是，EWS ID 可能包含非 URL 安全值，这会导致 REST 问题出现。

通过检查 [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) 属性，加载项可以确定它所加载的是哪个 Outlook 客户端。

### <a name="example"></a>示例

```js
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## <a name="get-the-rest-api-url"></a>获取 REST API URL

外接程序调用 REST API 所需的最后一部分信息是其发送 API 请求应使用的主机名。此信息在 [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) 属性中。

### <a name="example"></a>示例

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>调用 API

有访问令牌、项 ID 和 REST API URL 后，加载项可以将这些信息传递到调用 REST API 的后端服务，也可以使用 AJAX 直接调用 API。 下面的示例展示了如何调用 Outlook 邮件 REST API 来获取当前消息。

> [!IMPORTANT]
> 对于内部部署Exchange，使用 AJAX 或类似库的客户端请求将失败，因为该服务器安装程序不支持 CORS。

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    var subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>另请参阅

- 有关从 Outlook 加载项调用 REST API 的示例，请参阅 GitHub 上的 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。
- Outlook REST API 也可通过 Microsoft Graph 终结点获得，但存在一些关键区别，包括加载项如何获取访问令牌。 有关详细信息，请参阅[通过 Microsoft Graph 使用的 Outlook REST API](/outlook/rest/index#outlook-rest-api-via-microsoft-graph)。