---
title: 从 Outlook 加载项使用 Outlook REST API
description: 了解如何从 Outlook 加载项使用 Outlook REST API 获得访问令牌。
ms.date: 09/18/2020
localization_priority: Normal
ms.openlocfilehash: 067934f18b02d5106b58a7ec2a0de11a6ea35581
ms.sourcegitcommit: 09e1d8ff14b3c09a3eb11c91432c224a539181a4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2020
ms.locfileid: "48268549"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a><span data-ttu-id="10309-103">从 Outlook 加载项使用 Outlook REST API</span><span class="sxs-lookup"><span data-stu-id="10309-103">Use the Outlook REST APIs from an Outlook add-in</span></span>

<span data-ttu-id="10309-p101">[Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 命名空间提供访问许多邮件和约会的公用字段的权限。但是，在某些方案中，外接程序可能需要访问命名空间未公开的数据。例如，外接程序可能依赖于外部应用设置的自定义属性，或需要搜索用户邮箱中来自同一发件人的邮件。在这些方案中，[Outlook REST API](/outlook/rest/index) 是推荐的检索信息的方法。</span><span class="sxs-lookup"><span data-stu-id="10309-p101">The [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest/index) is the recommended method to retrieve the information.</span></span>

## <a name="get-an-access-token"></a><span data-ttu-id="10309-108">获取访问令牌</span><span class="sxs-lookup"><span data-stu-id="10309-108">Get an access token</span></span>

<span data-ttu-id="10309-p102">Outlook REST API 需要 `Authorization` 标头中的持有者令牌。应用通常使用 OAuth2 流检索令牌。不过，加载项也可以使用邮箱要求集 1.5 中引入的新 [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法检索令牌，而无需实现 OAuth2。</span><span class="sxs-lookup"><span data-stu-id="10309-p102">The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method introduced in the Mailbox requirement set 1.5.</span></span>

<span data-ttu-id="10309-112">通过将 `isRest` 选项设置为 `true`，可以请求获取与 REST API 兼容的令牌。</span><span class="sxs-lookup"><span data-stu-id="10309-112">By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.</span></span>

### <a name="add-in-permissions-and-token-scope"></a><span data-ttu-id="10309-113">外接程序权限和令牌范围</span><span class="sxs-lookup"><span data-stu-id="10309-113">Add-in permissions and token scope</span></span>

<span data-ttu-id="10309-p103">请务必考虑外接程序通过 REST API 所需要的访问级别。在大多数情况下，由 `getCallbackTokenAsync` 返回的令牌将仅提供对当前项的只读访问权限。即使外接程序在其清单中指定了 `ReadWriteItem` 权限级别也是如此。</span><span class="sxs-lookup"><span data-stu-id="10309-p103">It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.</span></span>

<span data-ttu-id="10309-p104">如果外接程序需要当前项目或用户邮箱中的其他项目的写权限，则外接程序必须在其清单中指定 `ReadWriteMailbox` 权限级别。在这种情况下，所返回的令牌将包含用户的邮件、事件和联系人的读取/写入访问权限。</span><span class="sxs-lookup"><span data-stu-id="10309-p104">If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.</span></span>

### <a name="example"></a><span data-ttu-id="10309-119">示例</span><span class="sxs-lookup"><span data-stu-id="10309-119">Example</span></span>

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

## <a name="get-the-item-id"></a><span data-ttu-id="10309-120">获取项 ID</span><span class="sxs-lookup"><span data-stu-id="10309-120">Get the item ID</span></span>

<span data-ttu-id="10309-121">加载项需要针对 REST 正确设置格式的项 ID，才能通过 REST 检索当前项。</span><span class="sxs-lookup"><span data-stu-id="10309-121">To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST.</span></span> <span data-ttu-id="10309-122">这可从 [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性获取，但应进行一些检查，以确保它是针对 REST 正确设置格式的 ID。</span><span class="sxs-lookup"><span data-stu-id="10309-122">This is obtained from the [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.</span></span>

- <span data-ttu-id="10309-123">在 Outlook Mobile 中，由 `Office.context.mailbox.item.itemId` 返回的值是适用于 REST 格式的 ID 并可按原样使用。</span><span class="sxs-lookup"><span data-stu-id="10309-123">In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.</span></span>
- <span data-ttu-id="10309-124">在其他 Outlook 客户端中，由 `Office.context.mailbox.item.itemId` 返回的值是适用于 EWS 格式的 ID，且必须使用 [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法进行转换。</span><span class="sxs-lookup"><span data-stu-id="10309-124">In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>
- <span data-ttu-id="10309-125">请注意，还必须将附件 ID 转换为带 REST 格式的 ID，才能使用它。</span><span class="sxs-lookup"><span data-stu-id="10309-125">Note you must also convert Attachment ID to a REST-formatted ID in order to use it.</span></span> <span data-ttu-id="10309-126">必须转换 ID 的原因是，EWS ID 可能包含非 URL 安全值，这会导致 REST 问题出现。</span><span class="sxs-lookup"><span data-stu-id="10309-126">The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.</span></span>

<span data-ttu-id="10309-127">通过检查 [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) 属性，加载项可以确定它所加载的是哪个 Outlook 客户端。</span><span class="sxs-lookup"><span data-stu-id="10309-127">Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) property.</span></span>

### <a name="example"></a><span data-ttu-id="10309-128">示例</span><span class="sxs-lookup"><span data-stu-id="10309-128">Example</span></span>

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

## <a name="get-the-rest-api-url"></a><span data-ttu-id="10309-129">获取 REST API URL</span><span class="sxs-lookup"><span data-stu-id="10309-129">Get the REST API URL</span></span>

<span data-ttu-id="10309-p107">外接程序调用 REST API 所需的最后一部分信息是其发送 API 请求应使用的主机名。此信息在 [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) 属性中。</span><span class="sxs-lookup"><span data-stu-id="10309-p107">The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property.</span></span>

### <a name="example"></a><span data-ttu-id="10309-132">示例</span><span class="sxs-lookup"><span data-stu-id="10309-132">Example</span></span>

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a><span data-ttu-id="10309-133">调用 API</span><span class="sxs-lookup"><span data-stu-id="10309-133">Call the API</span></span>

<span data-ttu-id="10309-134">有访问令牌、项 ID 和 REST API URL 后，加载项可以将这些信息传递到调用 REST API 的后端服务，也可以使用 AJAX 直接调用 API。</span><span class="sxs-lookup"><span data-stu-id="10309-134">After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX.</span></span> <span data-ttu-id="10309-135">下面的示例展示了如何调用 Outlook 邮件 REST API 来获取当前消息。</span><span class="sxs-lookup"><span data-stu-id="10309-135">The following example calls the Outlook Mail REST API to get the current message.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="10309-136">对于内部部署 Exchange 部署，使用 AJAX 或类似库的客户端请求将会失败，因为该服务器安装程序不支持 CORS。</span><span class="sxs-lookup"><span data-stu-id="10309-136">For on-premises Exchange deployments, client-side requests using AJAX or similar libraries fail because CORS isn't supported in that server setup.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="10309-137">另请参阅</span><span class="sxs-lookup"><span data-stu-id="10309-137">See also</span></span>

- <span data-ttu-id="10309-138">有关从 Outlook 加载项调用 REST API 的示例，请参阅 GitHub 上的 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。</span><span class="sxs-lookup"><span data-stu-id="10309-138">For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
- <span data-ttu-id="10309-139">Outlook REST API 也可通过 Microsoft Graph 终结点获得，但存在一些关键区别，包括加载项如何获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="10309-139">Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token.</span></span> <span data-ttu-id="10309-140">有关详细信息，请参阅[通过 Microsoft Graph 使用的 Outlook REST API](/outlook/rest/index#outlook-rest-api-via-microsoft-graph)。</span><span class="sxs-lookup"><span data-stu-id="10309-140">For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span></span>