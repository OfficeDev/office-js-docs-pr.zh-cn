---
title: 获取和设置 Outlook 加载项中的元数据
description: 可以使用以下漫游设置或自定义属性，管理 Outlook 加载项中的自定义数据。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: a06936892d9f2cdb7d83bc0c5097dfd2bdea0156
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839780"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>获取和设置 Outlook 加载项的元数据

您可以通过使用以下任一项管理 Outlook 外接程序中的自定义数据：

- 漫游设置，可管理用户邮箱的自定义数据。
- 自定义属性，可管理用户邮箱中某个项目的自定义数据。

两种方法均允许您访问仅可供 Outlook 外接程序访问的自定义数据，但两种方法分别存储数据。也就是说，自定义属性不能访问通过漫游设置存储的数据，反之亦然。数据存储在该邮箱的服务器上，并且在外接程序支持的所有外形因素上的后续 Outlook 会话中可访问。

## <a name="custom-data-per-mailbox-roaming-settings"></a>每个邮箱的自定义数据：漫游设置

您可以使用 [RoamingSettings](/javascript/api/outlook/office.RoamingSettings) 对象指定特定于用户的 Exchange 邮箱的数据，例如用户的个人数据和首选项。当您的邮件外接程序在设计在其上运行的任何设备（台式机、平板电脑或智能手机）上漫游时，可以访问漫游设置。

对该数据的更改存储在当前 Outlook 会话的这些设置的内存副本中。您应该在更新后显式保存所有漫游设置，以便用户下次在同一设备或任何其他受支持设备上打开您的外接程序时可以使用这些设置。


### <a name="roaming-settings-format"></a>漫游设置格式

**RoamingSettings** 对象中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 字符串。 

下面的结构示例假定有分别名为 `add-in_setting_name_0`、`add-in_setting_name_1` 和 `add-in_setting_name_2` 的三个已定义漫游设置。


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a>加载漫游设置

邮件加载项通常在 [Office.initialize](/javascript/api/office#office-initialize-reason-) 事件处理程序中加载漫游设置。 以下 JavaScript 代码示例演示了如何加载现有漫游设置并获取两个设置的值，即 **customerName** 和 **customerBalance**：


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>创建或分配漫游设置

紧接着前面的示例，下面的 JavaScript 函数 `setAddInSetting` 演示了如何使用 [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) 方法通过当天的日期设置名为 `cookie` 的设置，并使用 [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) 方法将所有漫游设置重新保存到服务器，以使数据持续存在。

如果设置不存在，该方法将创建该设置，并将 `set` 该设置分配给指定的值。 此方法 `saveAsync` 异步保存漫游设置。 此代码示例使用一个参数 `saveMyAddInSettingsCallback` `saveAsync`  `saveMyAddInSettingsCallback` _asyncResult_ 将回调方法传递给异步调用完成时。 此参数是一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象，其中包含异步调用的结果和所有详细信息。 可以使用可选的 _userContext_ 参数从异步调用向回调函数传递任何状态信息。

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a>删除漫游设置

通过扩展前面的示例，以下 JavaScript 函数  `removeAddInSetting` 显示了如何使用 [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a>邮箱中每个项目的自定义数据：自定义属性

可以使用 [CustomProperties](/javascript/api/outlook/office.CustomProperties) 对象指定用户邮箱中某个项目的特定数据。例如，邮件加载项可以对特定邮件进行分类，并使用自定义属性 `messageCategory` 标记类别。或者，如果邮件加载项使用邮件中的会议建议创建约会，则可以使用自定义属性跟踪这些约会。这可以确保当用户再次打开邮件时，邮件加载项不会再次创建约会。

与漫游设置类似，对自定义属性的更改将存储在当前 Outlook 会话的属性的内存副本中。为确保这些自定义属性在下次会话中可用，请使用 [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-)。

这些特定于加载项、特定于项目的自定义属性只能使用对象 `CustomProperties` 访问。 这些属性不同于 Outlook 对象模型中基于 MAPI 的自定义 [UserProperties，](/office/vba/api/Outlook.UserProperties) 以及 Exchange Web Services (EWS) 。 不能使用 Outlook `CustomProperties` 对象模型、EWS 或 REST 直接访问。 若要了解如何使用 EWS 或 REST 访问，请参阅"使用 `CustomProperties` [EWS](#get-custom-properties-using-ews-or-rest)或 REST 获取自定义属性"部分。

### <a name="using-custom-properties"></a>使用自定义属性

使用自定义属性之前，必须通过调用 [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法加载这些自定义属性。 创建属性包后，可以使用 [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) 和 [get](/javascript/api/outlook/office.CustomProperties) 方法添加和检索自定义属性。 必须使用 [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) 方法才能保存对属性包所做的任何更改。


 > [!NOTE]
 > 由于 Mac 版 Outlook 不缓存自定义属性，如果用户的网络断开，则 Mac 版 Outlook 中的邮件加载项将无法访问其自定义属性。


### <a name="custom-properties-example"></a>自定义属性示例


下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的方法。可以将此示例用作使用自定义属性的外接程序的起点。

此示例包括以下方法：


- [Office.initialize](/javascript/api/office#office-initialize-reason-) -- 初始化外接程序并从 Exchange 服务器中加载自定义属性包。

- **customPropsCallback** -- 获取并保存从服务器返回的自定义属性包以供将来使用。

- **updateProperty** -- 设置或更新特定属性，然后将更改保存到服务器。

- **removeProperty** -- 从属性包中删除特定属性，然后将该删除操作保存到服务器。


```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="get-custom-properties-using-ews-or-rest"></a>使用 EWS 或 REST 获取自定义属性

要使用 EWS 或 REST 获取 **CustomProperties**，首先应确定基于 MAPI 的扩展属性的名称。 然后可使用与获取基于 MAPI 的扩展属性相同的方式获取属性。

#### <a name="how-custom-properties-are-stored-on-an-item"></a>如何存储项的自定义属性

通过加载项设置的自定义属性并不等同于常规的基于 MAPI 的属性。 外接程序 API 将序列化所有加载项作为 JSON 有效负载，然后将它们保存在一个基于 MAPI 的扩展属性中，其名称为 (是加载项的 ID) ，属性集 `CustomProperties` `cecp-<app-guid>` `<app-guid>` GUID 是 `{00020329-0000-0000-C000-000000000046}` 。 （有关此对象的详细信息，请参阅 [MS-OXCEXT 2.2.5 邮件应用程序自定义属性](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481)。）随后可使用 EWS 或 REST 获取此基于 MAPI 的属性。

#### <a name="get-custom-properties-using-ews"></a>使用 EWS 获取自定义属性

您的邮件外接程序可以使用 `CustomProperties` EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作获取基于 MAPI 的扩展属性。 使用回调令牌访问服务器端，或在客户端使用 `GetItem` [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法访问。 在请求中，使用上一节提供的详细信息在其属性集内指定基于 MAPI 的属性。自定义属性如何 `GetItem` `CustomProperties` [存储在项目上](#how-custom-properties-are-stored-on-an-item)。

以下示例显示如何获取某个项目及其自定义属性。

> [!IMPORTANT]
> 在以下示例中，将 `<app-guid>` 替换为外接程序 ID。

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

如果在请求字符串中将其指定为其他 [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) 元素，也可以获得更多自定义属性。

#### <a name="get-custom-properties-using-rest"></a>使用 REST 获取自定义属性

可以在加载项中构建针对消息和事件的 REST 查询，以获取具有自定义属性的消息和事件。 在查询中，应包括使用 [如何存储项的自定义属性](#how-custom-properties-are-stored-on-an-item)部分提供的详细信息的 **CustomProperties** 基于 MAPI 的属性及其属性集。

以下示例显示了如何获取具有加载项设置的自定义属性的所有事件，并确保响应中包括对应的属性值，以便你能够应用其他筛选逻辑。

> [!IMPORTANT]
> 在以下示例中，将 `<app-guid>` 替换为加载项 ID。

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

有关使用 REST 获取基于 MAPI 的单值扩展属性的其他示例，请参阅[获取 singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true)。

以下示例显示如何获取某个项目及其自定义属性。 在 `done` 方法的回调函数中，`item.SingleValueExtendedProperties` 包含所请求的自定义属性的列表。

> [!IMPORTANT]
> 在以下示例中，将 `<app-guid>` 替换为外接程序 ID。

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## <a name="see-also"></a>另请参阅

- [MAPI 属性概述](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Outlook 属性概述](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [从 Outlook 加载项调用 Outlook REST API](use-rest-api.md)
- [从 Outlook 加载项调用 Web 服务](web-services.md)
- [Exchange 中 EWS 的属性和扩展属性](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [Exchange 中 EWS 的属性集和响应形状](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)