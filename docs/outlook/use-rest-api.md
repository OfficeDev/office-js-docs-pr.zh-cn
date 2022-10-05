---
title: 从 Outlook 加载项使用 Outlook REST API
description: 了解如何从 Outlook 加载项使用 Outlook REST API 获得访问令牌。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f62b2514f05341531a826c29e18c593a590fca0
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467214"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>从 Outlook 加载项使用 Outlook REST API

The [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.

> [!IMPORTANT]
> **Outlook REST API 已弃用**
>
> 有关更多详细信息，请参阅 2020 年 11 月公告) ，Outlook REST 终结点将于 [2022 年 11 月 30](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/) 日完全停用 (。 应迁移现有加载项以使用 [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph)。 有关指南，请参阅 [“比较 Microsoft Graph”和“Outlook REST API”终结点](/outlook/rest/compare-graph)。
>
> 为了帮助你进行迁移，使用 REST 服务的活动加载项有资格获得豁免，以继续使用该服务，直到 [2025 年 10 月 14 日 Outlook 2019 的扩展支持结束](/lifecycle/end-of-support/end-of-support-2025)。 这包括 2022 年 11 月 30 日之后开发的新加载项。 豁免基于加载项的清单 ID，适用于私密发布和 AppSource 托管的外接程序。
>
> 目前，正在测试使用 REST 服务的 Outlook 加载项的自动流量标识以进行豁免验证。 如果要参与此测试阶段，请在 2022 年 11 月之前完成 [REST API 加载项验证表](https://aka.ms/RESTCheck) 单。 有关详细信息，请参阅 [Office 加载项 2022 年 8 月社区呼叫博客文章](https://pnp.github.io/blog/office-add-ins-community-call/2022-08-10/)。

## <a name="get-an-access-token"></a>获取访问令牌

The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method introduced in the Mailbox requirement set 1.5.

通过将 `isRest` 选项设置为 `true`，可以请求获取与 REST API 兼容的令牌。

### <a name="add-in-permissions-and-token-scope"></a>外接程序权限和令牌范围

请务必考虑外接程序通过 REST API 所需要的访问级别。 在大多数情况下，由 `getCallbackTokenAsync` 返回的令牌将仅提供对当前项的只读访问权限。 即使加载项在其清单中指定 [了读/写项权限](understanding-outlook-add-in-permissions.md#readwrite-item-permission) 级别，也是如此。

如果外接程序需要对用户邮箱中的当前项目或其他项目进行写入访问，外接程序必须指定 [读/写邮箱权限](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission)。
其清单中的级别。 在这种情况下，所返回的令牌将包含用户的邮件、事件和联系人的读取/写入访问权限。

### <a name="example"></a>示例

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    const accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>获取项 ID

加载项需要针对 REST 正确设置格式的项 ID，才能通过 REST 检索当前项。 这可从 [Office.context.mailbox.item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性获取，但应进行一些检查，以确保它是针对 REST 正确设置格式的 ID。

- 在 Outlook Mobile 中，由 `Office.context.mailbox.item.itemId` 返回的值是适用于 REST 格式的 ID 并可按原样使用。
- 在其他 Outlook 客户端中，由 `Office.context.mailbox.item.itemId` 返回的值是适用于 EWS 格式的 ID，且必须使用 [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法进行转换。
- 请注意，还必须将附件 ID 转换为带 REST 格式的 ID，才能使用它。 必须转换 ID 的原因是，EWS ID 可能包含非 URL 安全值，这会导致 REST 问题出现。

通过检查 [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostname-member) 属性，加载项可以确定它所加载的是哪个 Outlook 客户端。

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

The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) property.

### <a name="example"></a>示例

```js
// Example: https://outlook.office.com
const restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>调用 API

有访问令牌、项 ID 和 REST API URL 后，加载项可以将这些信息传递到调用 REST API 的后端服务，也可以使用 AJAX 直接调用 API。 下面的示例展示了如何调用 Outlook 邮件 REST API 来获取当前消息。

> [!IMPORTANT]
> 对于本地 Exchange 部署，使用 AJAX 或类似库的客户端请求会失败，因为该服务器设置不支持 CORS。

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  const itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  const getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    const subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>另请参阅

- 有关从 Outlook 加载项调用 REST API 的示例，请参阅 GitHub 上的 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。
- Outlook REST API 也可通过 Microsoft Graph 终结点获得，但存在一些关键区别，包括加载项如何获取访问令牌。 有关详细信息，请参阅[通过 Microsoft Graph 使用的 Outlook REST API](/outlook/rest/index#outlook-rest-api-via-microsoft-graph)。
