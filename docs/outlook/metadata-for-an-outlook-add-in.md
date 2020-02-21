---
title: 获取和设置 Outlook 加载项中的元数据
description: 可以使用以下漫游设置或自定义属性，管理 Outlook 加载项中的自定义数据。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 86cc260b1a2fcb2a52145781fbcbef14ba5b2c96
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165916"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a><span data-ttu-id="7d51f-103">获取和设置 Outlook 加载项的元数据</span><span class="sxs-lookup"><span data-stu-id="7d51f-103">Get and set add-in metadata for an Outlook add-in</span></span>

<span data-ttu-id="7d51f-104">您可以通过使用以下任一项管理 Outlook 外接程序中的自定义数据：</span><span class="sxs-lookup"><span data-stu-id="7d51f-104">You can manage custom data in your Outlook add-in by using either of the following:</span></span>

- <span data-ttu-id="7d51f-105">漫游设置，可管理用户邮箱的自定义数据。</span><span class="sxs-lookup"><span data-stu-id="7d51f-105">Roaming settings, which manage custom data for a user's mailbox.</span></span>
- <span data-ttu-id="7d51f-106">自定义属性，可管理用户邮箱中某个项目的自定义数据。</span><span class="sxs-lookup"><span data-stu-id="7d51f-106">Custom properties, which manage custom data for an item in a user's mailbox.</span></span>

<span data-ttu-id="7d51f-p101">两种方法均允许您访问仅可供 Outlook 外接程序访问的自定义数据，但两种方法分别存储数据。也就是说，自定义属性不能访问通过漫游设置存储的数据，反之亦然。数据存储在该邮箱的服务器上，并且在外接程序支持的所有外形因素上的后续 Outlook 会话中可访问。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p101">Both of these give access to custom data that is only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings is not accessible by custom properties, and vice versa. The data is stored on the server for that mailbox, and is accessible in subsequent Outlook sessions on all the form factors that the add-in supports.</span></span>

## <a name="custom-data-per-mailbox-roaming-settings"></a><span data-ttu-id="7d51f-110">每个邮箱的自定义数据：漫游设置</span><span class="sxs-lookup"><span data-stu-id="7d51f-110">Custom data per mailbox: roaming settings</span></span>

<span data-ttu-id="7d51f-p102">您可以使用 [RoamingSettings](/javascript/api/outlook/office.RoamingSettings) 对象指定特定于用户的 Exchange 邮箱的数据，例如用户的个人数据和首选项。当您的邮件外接程序在设计在其上运行的任何设备（台式机、平板电脑或智能手机）上漫游时，可以访问漫游设置。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p102">You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](/javascript/api/outlook/office.RoamingSettings) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).</span></span>

<span data-ttu-id="7d51f-p103">对该数据的更改存储在当前 Outlook 会话的这些设置的内存副本中。您应该在更新后显式保存所有漫游设置，以便用户下次在同一设备或任何其他受支持设备上打开您的外接程序时可以使用这些设置。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p103">Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they will be available the next time the user opens your add-in, on the same or any other supported device.</span></span>


### <a name="roaming-settings-format"></a><span data-ttu-id="7d51f-116">漫游设置格式</span><span class="sxs-lookup"><span data-stu-id="7d51f-116">Roaming settings format</span></span>

<span data-ttu-id="7d51f-117">**RoamingSettings** 对象中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 字符串。</span><span class="sxs-lookup"><span data-stu-id="7d51f-117">The data in a **RoamingSettings** object is stored as a serialized JavaScript Object Notation (JSON) string.</span></span> 

<span data-ttu-id="7d51f-118">下面的结构示例假定有分别名为 `add-in_setting_name_0`、`add-in_setting_name_1` 和 `add-in_setting_name_2` 的三个已定义漫游设置。</span><span class="sxs-lookup"><span data-stu-id="7d51f-118">The following is an example of the structure, assuming there are three defined roaming settings named `add-in_setting_name_0`,  `add-in_setting_name_1`, and  `add-in_setting_name_2`.</span></span>


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a><span data-ttu-id="7d51f-119">加载漫游设置</span><span class="sxs-lookup"><span data-stu-id="7d51f-119">Loading roaming settings</span></span>

<span data-ttu-id="7d51f-120">邮件加载项通常在 [Office.initialize](/javascript/api/office#office-initialize-reason-) 事件处理程序中加载漫游设置。</span><span class="sxs-lookup"><span data-stu-id="7d51f-120">A mail add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office#office-initialize-reason-) event handler.</span></span> <span data-ttu-id="7d51f-121">以下 JavaScript 代码示例演示了如何加载现有漫游设置并获取两个设置的值，即 **customerName** 和 **customerBalance**：</span><span class="sxs-lookup"><span data-stu-id="7d51f-121">The following JavaScript code example shows how to load existing roaming settings and get the values of 2 settings, **customerName** and **customerBalance**:</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="7d51f-122">创建或分配漫游设置</span><span class="sxs-lookup"><span data-stu-id="7d51f-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="7d51f-123">紧接着前面的示例，下面的 JavaScript 函数 `setAddInSetting` 演示了如何使用 [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) 方法通过当天的日期设置名为 `cookie` 的设置，并使用 [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) 方法将所有漫游设置重新保存到服务器，以使数据持续存在。</span><span class="sxs-lookup"><span data-stu-id="7d51f-123">Continuing with the preceding example, the following JavaScript function,  `setAddInSetting`, shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) method to set a setting named `cookie` with today's date, and persist the data by using the [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) method to save all the roaming settings back to the server.</span></span>

<span data-ttu-id="7d51f-124">如果设置已不存在，**set** 方法会创建设置，并将其分配给指定的值。</span><span class="sxs-lookup"><span data-stu-id="7d51f-124">The **set** method creates the setting if the setting does not already exist, and assigns the setting to the specified value.</span></span> <span data-ttu-id="7d51f-125">**saveAsync** 方法会异步保存漫游设置。</span><span class="sxs-lookup"><span data-stu-id="7d51f-125">The **saveAsync** method saves roaming settings asynchronously.</span></span> <span data-ttu-id="7d51f-126">此代码示例将回叫方法 `saveMyAddInSettingsCallback` 传递给 **saveAsync**。</span><span class="sxs-lookup"><span data-stu-id="7d51f-126">This code sample passes a callback method, `saveMyAddInSettingsCallback`, to **saveAsync**.</span></span> <span data-ttu-id="7d51f-127">当异步调用完成时，会使用参数 _asyncResult_ 调用 `saveMyAddInSettingsCallback`。</span><span class="sxs-lookup"><span data-stu-id="7d51f-127">When the asynchronous call finishes,  `saveMyAddInSettingsCallback` is called by using one parameter, _asyncResult_.</span></span> <span data-ttu-id="7d51f-128">此参数是一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象，其中包含异步调用的结果和所有详细信息。</span><span class="sxs-lookup"><span data-stu-id="7d51f-128">This parameter is an [AsyncResult](/javascript/api/office/office.asyncresult) object that contains the result of and any details about the asynchronous call.</span></span> <span data-ttu-id="7d51f-129">可以使用可选的 _userContext_ 参数从异步调用向回调函数传递任何状态信息。</span><span class="sxs-lookup"><span data-stu-id="7d51f-129">You can use the optional _userContext_ parameter to pass any state information from the asynchronous call to the callback function.</span></span>

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


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="7d51f-130">删除漫游设置</span><span class="sxs-lookup"><span data-stu-id="7d51f-130">Removing a roaming setting</span></span>

<span data-ttu-id="7d51f-131">通过扩展前面的示例，以下 JavaScript 函数  `removeAddInSetting` 显示了如何使用 [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。</span><span class="sxs-lookup"><span data-stu-id="7d51f-131">Also extending the preceding examples, the following JavaScript function,  `removeAddInSetting`, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


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


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a><span data-ttu-id="7d51f-132">邮箱中每个项目的自定义数据：自定义属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-132">Custom data per item in a mailbox: custom properties</span></span>

<span data-ttu-id="7d51f-p106">可以使用 [CustomProperties](/javascript/api/outlook/office.CustomProperties) 对象指定用户邮箱中某个项目的特定数据。例如，邮件加载项可以对特定邮件进行分类，并使用自定义属性 `messageCategory` 标记类别。或者，如果邮件加载项使用邮件中的会议建议创建约会，则可以使用自定义属性跟踪这些约会。这可以确保当用户再次打开邮件时，邮件加载项不会再次创建约会。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p106">You can specify data specific to an item in the user's mailbox using the [CustomProperties](/javascript/api/outlook/office.CustomProperties) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in doesn't offer to create the appointment a second time.</span></span>

<span data-ttu-id="7d51f-p107">与漫游设置类似，对自定义属性的更改将存储在当前 Outlook 会话的属性的内存副本中。为确保这些自定义属性在下次会话中可用，请使用 [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-)。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p107">Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, use [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span></span>

<span data-ttu-id="7d51f-p108">这些加载项和项目特定的自定义属性只能使用 **CustomProperties** 对象访问。这些属性不同于 Outlook 对象模型中基于 MAPI 的自定义属性 [UserProperties](/office/vba/api/Outlook.UserProperties)，也不同于 Exchange Web 服务 (EWS) 中的扩展属性。无法使用 Outlook 对象模型、EWS 或 REST 直接访问 **CustomProperties**。若要了解如何使用 EWS 或 REST 访问 **CustomProperties**，请参阅[使用 EWS 或 REST 获取自定义属性](#get-custom-properties-using-ews-or-rest)部分。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p108">These add-in-specific, item-specific custom properties can only be accessed by using the **CustomProperties** object. These properties are different from the custom, MAPI-based [UserProperties](/office/vba/api/Outlook.UserProperties) in the Outlook object model, and extended properties in Exchange Web Services (EWS). You cannot directly access **CustomProperties** by using the Outlook object model, EWS, or REST. To learn how to access **CustomProperties** using EWS or REST, see the section [Get custom properties using EWS or REST](#get-custom-properties-using-ews-or-rest).</span></span>

### <a name="using-custom-properties"></a><span data-ttu-id="7d51f-143">使用自定义属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-143">Using custom properties</span></span>

<span data-ttu-id="7d51f-144">使用自定义属性之前，必须通过调用 [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法加载这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-144">Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="7d51f-145">创建属性包后，可以使用 [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) 和 [get](/javascript/api/outlook/office.CustomProperties) 方法添加和检索自定义属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-145">After you have created the property bag, you can use the [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) and [get](/javascript/api/outlook/office.CustomProperties) methods to add and retrieve custom properties.</span></span> <span data-ttu-id="7d51f-146">必须使用 [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) 方法才能保存对属性包所做的任何更改。</span><span class="sxs-lookup"><span data-stu-id="7d51f-146">You must use the [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) method to save any changes that you make to the property bag.</span></span>


 > [!NOTE]
 > <span data-ttu-id="7d51f-147">由于 Mac 版 Outlook 不缓存自定义属性，如果用户的网络断开，则 Mac 版 Outlook 中的邮件加载项将无法访问其自定义属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-147">Because Outlook on Mac doesn't cache custom properties, if the user's network goes down, mail add-ins in Outlook on Mac would not be able to access their custom properties.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="7d51f-148">自定义属性示例</span><span class="sxs-lookup"><span data-stu-id="7d51f-148">Custom properties example</span></span>


<span data-ttu-id="7d51f-p110">下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的方法。可以将此示例用作使用自定义属性的外接程序的起点。</span><span class="sxs-lookup"><span data-stu-id="7d51f-p110">The following example shows a simplified set of methods for an Outlook add-in that uses custom properties. You can use this example as a starting point for your add-in that uses custom properties.</span></span>

<span data-ttu-id="7d51f-151">此示例包括以下方法：</span><span class="sxs-lookup"><span data-stu-id="7d51f-151">This example includes the following methods:</span></span>


- <span data-ttu-id="7d51f-152">[Office.initialize](/javascript/api/office#office-initialize-reason-) -- 初始化外接程序并从 Exchange 服务器中加载自定义属性包。</span><span class="sxs-lookup"><span data-stu-id="7d51f-152">[Office.initialize](/javascript/api/office#office-initialize-reason-) -- Initializes the add-in and loads the custom property bag from the Exchange server.</span></span>

- <span data-ttu-id="7d51f-153">**customPropsCallback** -- 获取并保存从服务器返回的自定义属性包以供将来使用。</span><span class="sxs-lookup"><span data-stu-id="7d51f-153">**customPropsCallback** -- Gets the custom property bag that is returned from the server and saves it for later use.</span></span>

- <span data-ttu-id="7d51f-154">**updateProperty** -- 设置或更新特定属性，然后将更改保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="7d51f-154">**updateProperty** -- Sets or updates a specific property, and then saves the change to the server.</span></span>

- <span data-ttu-id="7d51f-155">**removeProperty** -- 从属性包中删除特定属性，然后将该删除操作保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="7d51f-155">**removeProperty** -- Removes a specific property from the property bag, and then saves the removal to the server.</span></span>


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

### <a name="get-custom-properties-using-ews-or-rest"></a><span data-ttu-id="7d51f-156">使用 EWS 或 REST 获取自定义属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-156">Get custom properties using EWS or REST</span></span>

<span data-ttu-id="7d51f-157">要使用 EWS 或 REST 获取 **CustomProperties**，首先应确定基于 MAPI 的扩展属性的名称。</span><span class="sxs-lookup"><span data-stu-id="7d51f-157">To get **CustomProperties** using EWS or REST, you should first determine the name of its MAPI-based extended property.</span></span> <span data-ttu-id="7d51f-158">然后可使用与获取基于 MAPI 的扩展属性相同的方式获取属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-158">You can then get that property in the same way you would get any MAPI-based extended property.</span></span>

#### <a name="how-custom-properties-are-stored-on-an-item"></a><span data-ttu-id="7d51f-159">如何存储项的自定义属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-159">How custom properties are stored on an item</span></span>

<span data-ttu-id="7d51f-160">通过加载项设置的自定义属性并不等同于常规的基于 MAPI 的属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-160">Custom properties set by an add-in are not equivalent to normal MAPI-based properties.</span></span> <span data-ttu-id="7d51f-161">加载项 API 会将所有加载项的 **CustomProperties** 序列化为 JSON 有效负载，然后再将其保存到名为 `cecp-<app-guid>`（`<app-guid>` 为加载项的 ID）且属性集 GUID 为 `{00020329-0000-0000-C000-000000000046}` 的单个基于 MAPI 的扩展属性中。</span><span class="sxs-lookup"><span data-stu-id="7d51f-161">Add-in APIs serialize all your add-in's **CustomProperties** as a JSON payload and then save them in a single MAPI-based extended property whose name is `cecp-<app-guid>` (`<app-guid>` is your add-in's ID) and property set GUID is `{00020329-0000-0000-C000-000000000046}`.</span></span> <span data-ttu-id="7d51f-162">（有关此对象的详细信息，请参阅 [MS-OXCEXT 2.2.5 邮件应用程序自定义属性](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx)。）随后可使用 EWS 或 REST 获取此基于 MAPI 的属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-162">(For more information about this object, see [MS-OXCEXT 2.2.5 Mail App Custom Properties](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx).) You can then use EWS or REST to get this MAPI-based property.</span></span>

#### <a name="get-custom-properties-using-ews"></a><span data-ttu-id="7d51f-163">使用 EWS 获取自定义属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-163">Get custom properties using EWS</span></span>

<span data-ttu-id="7d51f-164">邮件加载项可以使用 EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作获取 **CustomProperties** 基于 MAPI 的扩展属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-164">Your mail add-in can get the **CustomProperties** MAPI-based extended property by using the EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation.</span></span> <span data-ttu-id="7d51f-165">使用回调令牌或者使用 [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法，访问服务器端的 **GetItem**。</span><span class="sxs-lookup"><span data-stu-id="7d51f-165">Access **GetItem** on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span> <span data-ttu-id="7d51f-166">在 **GetItem** 请求中，指定使用前面部分[如何存储项的自定义属性](#how-custom-properties-are-stored-on-an-item)提供的详细信息的属性集中的 **CustomProperties** 基于 MAPI 的属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-166">In the **GetItem** request, specify the **CustomProperties** MAPI-based property in its property set using the details provided in the preceding section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="7d51f-167">以下示例显示如何获取某个项目及其自定义属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-167">The following example shows how to get an item and its custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d51f-168">在以下示例中，将 `<app-guid>` 替换为外接程序 ID。</span><span class="sxs-lookup"><span data-stu-id="7d51f-168">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

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

<span data-ttu-id="7d51f-169">如果在请求字符串中将其指定为其他 [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) 元素，也可以获得更多自定义属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-169">You can also get more custom properties if you specify them in the request string as other [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elements.</span></span>

#### <a name="get-custom-properties-using-rest"></a><span data-ttu-id="7d51f-170">使用 REST 获取自定义属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-170">Get custom properties using REST</span></span>

<span data-ttu-id="7d51f-171">可以在加载项中构建针对消息和事件的 REST 查询，以获取具有自定义属性的消息和事件。</span><span class="sxs-lookup"><span data-stu-id="7d51f-171">In your add-in, you can construct your REST query against messages and events to get the ones that already have custom properties.</span></span> <span data-ttu-id="7d51f-172">在查询中，应包括使用[如何存储项的自定义属性](#how-custom-properties-are-stored-on-an-item)部分提供的详细信息的 **CustomProperties** 基于 MAPI 的属性及其属性集。</span><span class="sxs-lookup"><span data-stu-id="7d51f-172">In your query, you should include the **CustomProperties** MAPI-based property and its property set using the details provided in the section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="7d51f-173">以下示例显示了如何获取具有加载项设置的自定义属性的所有事件，并确保响应中包括对应的属性值，以便你能够应用其他筛选逻辑。</span><span class="sxs-lookup"><span data-stu-id="7d51f-173">The following example shows how to get all events that have any custom properties set by your add-in and ensure that the response includes the value of the property so you can apply further filtering logic.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d51f-174">在以下示例中，将 `<app-guid>` 替换为加载项 ID。</span><span class="sxs-lookup"><span data-stu-id="7d51f-174">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

<span data-ttu-id="7d51f-175">有关使用 REST 获取基于 MAPI 的单值扩展属性的其他示例，请参阅[获取 singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0)。</span><span class="sxs-lookup"><span data-stu-id="7d51f-175">For other examples that use REST to get single-value MAPI-based extended properties, see [Get singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0).</span></span>

<span data-ttu-id="7d51f-176">以下示例显示如何获取某个项目及其自定义属性。</span><span class="sxs-lookup"><span data-stu-id="7d51f-176">The following example shows how to get an item and its custom properties.</span></span> <span data-ttu-id="7d51f-177">在 `done` 方法的回调函数中，`item.SingleValueExtendedProperties` 包含所请求的自定义属性的列表。</span><span class="sxs-lookup"><span data-stu-id="7d51f-177">In the callback function for the `done` method, `item.SingleValueExtendedProperties` contains a list of the requested custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d51f-178">在以下示例中，将 `<app-guid>` 替换为外接程序 ID。</span><span class="sxs-lookup"><span data-stu-id="7d51f-178">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="7d51f-179">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7d51f-179">See also</span></span>

- [<span data-ttu-id="7d51f-180">MAPI 属性概述</span><span class="sxs-lookup"><span data-stu-id="7d51f-180">MAPI Property Overview</span></span>](/office/client-developer/outlook/mapi/mapi-property-overview)
- [<span data-ttu-id="7d51f-181">Outlook 属性概述</span><span class="sxs-lookup"><span data-stu-id="7d51f-181">Outlook Properties Overview</span></span>](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [<span data-ttu-id="7d51f-182">从 Outlook 加载项调用 Outlook REST API</span><span class="sxs-lookup"><span data-stu-id="7d51f-182">Call Outlook REST APIs from an Outlook add-in</span></span>](use-rest-api.md)
- [<span data-ttu-id="7d51f-183">从 Outlook 加载项调用 Web 服务</span><span class="sxs-lookup"><span data-stu-id="7d51f-183">Call web services from an Outlook add-in</span></span>](web-services.md)
- [<span data-ttu-id="7d51f-184">Exchange 中 EWS 的属性和扩展属性</span><span class="sxs-lookup"><span data-stu-id="7d51f-184">Properties and extended properties in EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [<span data-ttu-id="7d51f-185">Exchange 中 EWS 的属性集和响应形状</span><span class="sxs-lookup"><span data-stu-id="7d51f-185">Property sets and response shapes in EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)
