---
title: 从 Outlook 加载项使用 Exchange Web 服务 (EWS)
description: 提供的示例显示 Outlook 加载项如何通过 Exchange Web 服务请求信息。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6590967ef79e03cdbeee612199aba7a681b6dcdb
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713061"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>从 Outlook 加载项调用 Web 服务

您的外接程序可使用运行 Exchange Server 2013 的计算机中的 Exchange Web 服务 (EWS)，该 Web 服务可在为外接程序的 UI 提供源位置的服务器上获得，也可在 Internet 上获得。本文提供展示 Outlook 外接程序如何从 EWS 请求信息的示例。

您用来调用 Web 服务的方法随 Web 服务所在的位置的不同而不同。表 1 列出了可以基于位置调用 Web 服务的不同方法。

**表 1.从 Outlook 外接程序调用 Web 服务的方式**

|**Web 服务位置**|**调用 Web 服务的方法**|
|:-----|:-----|
|托管客户端邮箱的 Exchange 服务器|使用 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法可调用外接程序支持的 EWS 操作。承载邮箱的 Exchange 服务器还会公开 EWS。|
|为加载项 UI 提供源位置的 Web 服务器|使用标准 JavaScript 技术调用 Web 服务。UI 框架中的 JavaScript 代码将在提供 UI 的 Web 服务器的上下文中运行。因此，此代码可以调用该服务器上的 Web 服务，而不会导致出现跨网站脚本错误。|
|所有其他位置|为提供 UI 源位置的 Web 服务器上的 Web 服务创建代理。如果您不提供代理，跨网站脚本错误将阻止外接程序运行。提供代理的一种方式是使用 JSON/P。有关详细信息，请参阅 [Office 外接程序的隐私和安全性](../concepts/privacy-and-security.md)。|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>使用 makeEwsRequestAsync 方法访问 EWS 操作

可以使用 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法向托管用户邮箱的 Exchange 服务器发出 EWS 请求。

EWS 服务支持 Exchange 服务器中的不同操作；例如复制、查找、更新或发送项目的项目级操作，以及创建、获取或更新文件夹的文件夹级操作。若要执行 EWS 操作，请创建一个执行该操作的 XML SOAP 请求。当操作完成时，你将获得包含该操作相关数据的 XML SOAP 响应。EWS SOAP 请求和响应遵循 Messages.xsd 文件中定义的架构。正如其他 EWS 架构文件一样，Message.xsd 文件位于托管 EWS 的 IIS 虚拟目录中。

若要使用该 `makeEwsRequestAsync` 方法启动 EWS 操作，请提供以下内容：

- 针对该 EWS 操作的 SOAP 请求的 XML，作为  _data_ 形参的实参

- 回调方法（作为  _callback_ 实参）

- 该回调方法的任何可选输入数据（作为  _userContext_ 实参）

EWS SOAP 请求完成后，Outlook 将使用一个实参（是一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用该回调方法。 回调方法可以访问对象的`AsyncResult`两个属性：`value`包含 EWS 操作的 XML SOAP 响应的属性，以及包含作为`userContext`参数传递的任何数据的属性（可选）`asyncContext`。 通常，回调方法稍后会解析 SOAP 响应中的 XML 以获取所有相关信息，并相应地处理这些信息。

## <a name="tips-for-parsing-ews-responses"></a>解析 EWS 响应的提示

从 EWS 操作分析 SOAP 响应时，请注意以下与浏览器相关的问题。

- 使用 DOM 方法 `getElementsByTagName`时指定标记名称的前缀，以包含对 Internet Explorer 的支持。

  `getElementsByTagName` 行为不同，具体取决于浏览器类型。 例如，EWS 响应可以包含以下 XML (格式化和缩写，用于显示目的) 。

   ```XML
   <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
   PropertyName="MyProperty" 
   PropertyType="String"/>
   <t:Value>{
   ...
   }</t:Value></t:ExtendedProperty>
   ```

   如下所示，代码将适用于 Chrome 等浏览器，以获取标记括起来的 `ExtendedProperty` XML。

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("ExtendedProperty")
   });
   ```

   在 Internet Explorer 上，必须包括 `t:` 标记名称的前缀，如下所示。

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("t:ExtendedProperty")
   });
   ```

- 使用 DOM 属性 `textContent` 获取 EWS 响应中标记的内容，如下所示。

   ```js
   content = $.parseJSON(value.textContent);
   ```

   对于 EWS 响应中的某些标记，其他属性（例如 `innerHTML` 可能不适用于 Internet Explorer）。

## <a name="example"></a>示例

以下示例调用 `makeEwsRequestAsync` 以使用 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作获取项的主题。 此示例包括以下三个函数。

- `getSubjectRequest`&ndash;将项 ID 作为输入，并返回要调用指定项的 SOAP 请求`GetItem`的 XML。

- `sendRequest`&ndash;调用`getSubjectRequest`以获取所选项的 SOAP 请求，然后传递 SOAP 请求和回调方法`callback`，以`makeEwsRequestAsync`获取指定项的主题。

- `callback` &ndash; 处理包含有关指定项目的任何主题和其他信息的 SOAP 响应。

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   const result = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return result;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   const mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   const result = asyncResult.value;
   const context = asyncResult.context;

   // Process the returned response here.
}
```

## <a name="ews-operations-that-add-ins-support"></a>外接程序支持的 EWS 操作

Outlook 外接程序可以通过该方法访问 EWS `makeEwsRequestAsync` 中可用的一部分操作。 如果不熟悉 EWS 操作以及如何使用 `makeEwsRequestAsync` 该方法访问操作，请从 SOAP 请求示例开始自定义 _数据_ 参数。

下面介绍了如何使用该 `makeEwsRequestAsync` 方法。

1. 在 XML 中，用适当值替换所有项目 ID 和相关 EWS 操作属性。

1. 将 SOAP 请求作为  _数据_ 参数的 `makeEwsRequestAsync`参数包括在内。

1. 指定回调方法和调用 `makeEwsRequestAsync`。

1. 在回调方法中，验证 SOAP 响应中操作的结果。

1. 根据需要使用 EWS 操作的结果。

下表列出了外接程序支持的 EWS 操作。若要查看 SOAP 请求和响应的示例，请选择各操作对应的链接。有关 EWS 操作的详细信息，请参阅 [在交换 EWS 操作](/exchange/client-developer/web-service-reference/ews-operations-in-exchange)。

**表 2.支持的 EWS 操作**

|**EWS 操作**|**说明**|
|:-----|:-----|
|[CopyItem 操作](/exchange/client-developer/web-service-reference/copyitem-operation)|在 Exchange 存储的指定文件夹中复制指定项目并在其中放入新项目。|
|[CreateFolder 操作](/exchange/client-developer/web-service-reference/createfolder-operation)|在 Exchange 存储中的指定位置创建文件夹。|
|[CreateItem 操作](/exchange/client-developer/web-service-reference/createitem-operation)|在 Exchange 存储中创建指定项目。|
|[ExpandDL 操作](/exchange/client-developer/web-service-reference/expanddl-operation)|显示通讯组列表的完整成员身份。|
|[FindConversation 操作](/exchange/client-developer/web-service-reference/findconversation-operation)|在 Exchange 存储的指定文件夹中枚举会话列表。|
|[FindFolder 操作](/exchange/client-developer/web-service-reference/findfolder-operation)|查找指定文件夹的子文件夹并返回描述这组子文件夹的一组属性。|
|[FindItem 操作](/exchange/client-developer/web-service-reference/finditem-operation)|标识位于 Exchange 存储的指定文件夹中的项目。|
|[GetConversationItems 操作](/exchange/client-developer/web-service-reference/getconversationitems-operation)|在会话中获取排列为节点的一个或多个项集。|
|[GetFolder 操作](/exchange/client-developer/web-service-reference/getfolder-operation)|从 Exchange 存储中获取文件夹的指定属性和内容。|
|[GetItem 操作](/exchange/client-developer/web-service-reference/getitem-operation)|从 Exchange 存储中获取项目的指定属性和内容。|
|[GetUserAvailability 操作](/exchange/client-developer/web-service-reference/getuseravailability-operation)|提供特定时间段内有关一组用户、会议室和资源的可用性的详细信息。|
|[MarkAsJunk 操作](/exchange/client-developer/web-service-reference/markasjunk-operation)|将电子邮件移动到"垃圾邮件"文件夹，并相应地在阻止的发件人名单中添加或删除邮件的发件人。|
|[MoveItem 操作](/exchange/client-developer/web-service-reference/moveitem-operation)|将项目移动到 Exchange 存储中的单个目标文件夹。|
|[ResolveNames 操作](/exchange/client-developer/web-service-reference/resolvenames-operation)|解析不确定的电子邮件地址和显示名称。|
|[SendItem 操作](/exchange/client-developer/web-service-reference/senditem-operation)|发送位于 Exchange 存储中的电子邮件。|
|[UpdateFolder 操作](/exchange/client-developer/web-service-reference/updatefolder-operation)|修改 Exchange 存储中现有文件夹的属性。|
|[UpdateItem 操作](/exchange/client-developer/web-service-reference/updateitem-operation)|修改 Exchange 存储中现有项的属性。|

 > [!NOTE]
 > FAI（文件夹关联信息）项不能通过外接程序进行更新（或创建）。 这些隐藏的消息存储在文件夹中，用于存储各种设置和辅助数据。  尝试使用 UpdateItem 操作会导致以下 ErrorAccessDenied 错误抛出：“不得使用 Office 扩展来更新此类项”。 此外，也可以使用 [EWS 托管 API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) 通过 Windows 客户端或服务器应用更新这些项。 建议谨慎操作，因为内部服务类型数据结构可能会发生变化并破坏解决方案。

## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a>makeEwsRequestAsync 的身份验证和权限注意事项

使用此 `makeEwsRequestAsync` 方法时，将使用当前用户的电子邮件帐户凭据对请求进行身份验证。 该 `makeEwsRequestAsync` 方法为你管理凭据，这样就不必向请求提供身份验证凭据。

> [!NOTE]
> 服务器管理员必须使用 [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) 或 [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) cmdlet 在客户端访问服务器 EWS 目录上将 _OAuthAuthentication_ 参数设置为 **true** ，以便使 `makeEwsRequestAsync` 该方法能够发出 EWS 请求。

加载项必须在其外接程序清单中指定 `ReadWriteMailbox` 权限才能使用该 `makeEwsRequestAsync` 方法。 有关使用权`ReadWriteMailbox`限的信息，请参阅[了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)中的 [ReadWriteMailbox](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) 权限部分。

## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全性](../concepts/privacy-and-security.md)
- [解决 Office 外接程序中的同源策略限制](../develop/addressing-same-origin-policy-limitations.md)
- [Exchange 的 EWS 引用](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- [Outlook 和 Exchange 中的 EWS 的邮件应用程序](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

请参阅以下内容，了解如何使用 ASP.NET Web API为加载项创建后端服务。

- [使用 ASP.NET Web API 为 Office 外接程序创建 Web 服务](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [使用 ASP.NET Web API 构建 HTTP 服务的基础知识](https://dotnet.microsoft.com/apps/aspnet/apis)
