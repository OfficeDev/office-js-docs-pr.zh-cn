---
title: 管理 Outlook 外接程序的状态和设置
description: 了解如何保留 Outlook 外接程序的外接程序状态和设置。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 7d981107da68c329d209834059bfac494d6ccae4
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596646"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a><span data-ttu-id="ae68f-103">管理 Outlook 外接程序的状态和设置</span><span class="sxs-lookup"><span data-stu-id="ae68f-103">Manage state and settings for an Outlook add-in</span></span>

> [!NOTE]
> <span data-ttu-id="ae68f-104">阅读本文之前，请查看本文档的 "**核心概念**" 一节中的 "[保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)"。</span><span class="sxs-lookup"><span data-stu-id="ae68f-104">Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.</span></span>

<span data-ttu-id="ae68f-105">对于 Outlook 外接程序，Office JavaScript API 提供[RoamingSettings](/javascript/api/outlook/office.roamingsettings)和[CustomProperties](/javascript/api/outlook/office.customproperties)对象，以在各会话之间保存外接程序状态，如下表所述。</span><span class="sxs-lookup"><span data-stu-id="ae68f-105">For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="ae68f-106">在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](../reference/manifest/id.md) 相关联。</span><span class="sxs-lookup"><span data-stu-id="ae68f-106">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="ae68f-107">**对象**</span><span class="sxs-lookup"><span data-stu-id="ae68f-107">**Object**</span></span>|<span data-ttu-id="ae68f-108">**存储位置**</span><span class="sxs-lookup"><span data-stu-id="ae68f-108">**Storage location**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="ae68f-109">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae68f-109">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="ae68f-110">安装了加载项的用户 Exchange 服务器邮箱。</span><span class="sxs-lookup"><span data-stu-id="ae68f-110">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="ae68f-111">由于这些设置存储在用户的服务器邮箱中，因此当加载项运行在任何访问该用户邮箱的受支持客户端主机应用程序或浏览器的上下文中时，这些设置可随用户“漫游”且可供加载项使用。</span><span class="sxs-lookup"><span data-stu-id="ae68f-111">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="ae68f-112">Outlook 外接程序漫游设置仅供创建它们的外接程序使用，并且只能从安装了外接程序的邮箱使用。</span><span class="sxs-lookup"><span data-stu-id="ae68f-112">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|
|[<span data-ttu-id="ae68f-113">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="ae68f-113">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="ae68f-p103">加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。</span><span class="sxs-lookup"><span data-stu-id="ae68f-p103">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="ae68f-116">如何将 Outlook 加载项用户邮箱中的设置保存为漫游设置</span><span class="sxs-lookup"><span data-stu-id="ae68f-116">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>

<span data-ttu-id="ae68f-117">Outlook 加载项可以使用 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) 对象保存特定于用户邮箱的加载项状态和设置数据。</span><span class="sxs-lookup"><span data-stu-id="ae68f-117">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="ae68f-118">仅代表用户运行该加载项的 Outlook 加载项才可访问此数据。</span><span class="sxs-lookup"><span data-stu-id="ae68f-118">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="ae68f-119">这些数据将存储在用户的 Exchange Server 邮箱上，并且在用户登录到其帐户并运行 Outlook 加载项时可访问这些数据。</span><span class="sxs-lookup"><span data-stu-id="ae68f-119">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>

### <a name="loading-roaming-settings"></a><span data-ttu-id="ae68f-120">加载漫游设置</span><span class="sxs-lookup"><span data-stu-id="ae68f-120">Loading roaming settings</span></span>

<span data-ttu-id="ae68f-p105">Outlook 外接程序通常在 [Office.initialize](/javascript/api/office) 事件处理程序中加载漫游设置。以下 JavaScript 代码示例演示了如何加载现有漫游设置。</span><span class="sxs-lookup"><span data-stu-id="ae68f-p105">An Outlook add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>

```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="ae68f-123">创建或分配漫游设置</span><span class="sxs-lookup"><span data-stu-id="ae68f-123">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="ae68f-p106">紧接着前面的示例，下面的  `setAppSetting` 函数演示如何使用 [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) 方法通过当天的日期设置或更新名为 `cookie` 的设置。然后使用 [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) 方法将所有漫游设置保存回 Exchange Server。</span><span class="sxs-lookup"><span data-stu-id="ae68f-p106">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>

```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

<span data-ttu-id="ae68f-126">**saveAsync** 方法将异步保存漫游设置，并采用可选回调函数。</span><span class="sxs-lookup"><span data-stu-id="ae68f-126">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="ae68f-127">此代码示例会将名为 `saveMyAppSettingsCallback` 的回调函数传递给 **saveAsync** 方法。</span><span class="sxs-lookup"><span data-stu-id="ae68f-127">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="ae68f-128">当异步调用返回时，`saveMyAppSettingsCallback` 函数的 _asyncResult_ 参数提供对 [AsyncResult](/javascript/api/outlook) 对象的访问权限，你可以使用该对象通过 **AsyncResult.status** 属性确定操作是否成功。</span><span class="sxs-lookup"><span data-stu-id="ae68f-128">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/outlook) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>

### <a name="removing-a-roaming-setting"></a><span data-ttu-id="ae68f-129">删除漫游设置</span><span class="sxs-lookup"><span data-stu-id="ae68f-129">Removing a roaming setting</span></span>

<span data-ttu-id="ae68f-130">进一步展开前面的示例，以下  `removeAppSetting` 函数演示了如何使用 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。</span><span class="sxs-lookup"><span data-stu-id="ae68f-130">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="ae68f-131">如何按项目将 Outlook 外接程序的设置保存为自定义属性</span><span class="sxs-lookup"><span data-stu-id="ae68f-131">How to save settings per item for Outlook add-ins as custom properties</span></span>

<span data-ttu-id="ae68f-p108">自定义属性允许 Outlook 外接程序存储其使用的有关项目的信息。例如，如果 Outlook 外接程序根据邮件中的会议建议创建约会，则可以使用自定义属性存储创建了会议的事实。这确保了如果再次打开邮件，Outlook 外接程序不再可供创建约会。</span><span class="sxs-lookup"><span data-stu-id="ae68f-p108">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="ae68f-p109">在您将自定义属性用于特定邮件、约会或会议请求项目之前，必须通过调用  [Item](/javascript/api/outlook/office.mailbox) 对象的 **loadCustomPropertiesAsync** 方法将属性加载到内存中。如果为当前项目设置了任何自定义属性，此时会从 Exchanger Server 加载这些属性。在您加载了属性以后，可以使用 [CustomProperties](/javascript/api/outlook/office.customproperties#set-name--value-) 对象的 [set](/javascript/api/outlook/office.roamingsettings) 和 **get** 方法添加、更新和检索内存中的属性。要保存对于项目的自定义属性所做的任何更改，必须使用 [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) 方法在 Exchanger Server上保留对项目所做的更改。</span><span class="sxs-lookup"><span data-stu-id="ae68f-p109">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>

### <a name="custom-properties-example"></a><span data-ttu-id="ae68f-139">自定义属性示例</span><span class="sxs-lookup"><span data-stu-id="ae68f-139">Custom properties example</span></span>

<span data-ttu-id="ae68f-p110">下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的函数。可以将此示例用作使用自定义属性的 Outlook 外接程序的起点。</span><span class="sxs-lookup"><span data-stu-id="ae68f-p110">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="ae68f-142">使用这些函数的 Outlook 加载项通过对 `_customProps` 变量调用 **get** 方法来检索任何自定义属性，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="ae68f-142">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>

```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="ae68f-143">此示例包括以下函数：</span><span class="sxs-lookup"><span data-stu-id="ae68f-143">This example includes the following functions:</span></span>

|<span data-ttu-id="ae68f-144">**函数名称**</span><span class="sxs-lookup"><span data-stu-id="ae68f-144">**Function name**</span></span>|<span data-ttu-id="ae68f-145">**说明**</span><span class="sxs-lookup"><span data-stu-id="ae68f-145">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="ae68f-146">从 Exchange 服务器初始化外接程序并加载当前项目的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="ae68f-146">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="ae68f-147">获取从 Exchange 服务器返回的自定义属性并将其保存以供后续之用。</span><span class="sxs-lookup"><span data-stu-id="ae68f-147">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="ae68f-148">设置或更新特定属性，然后将更改保存到 Exchange 服务器。</span><span class="sxs-lookup"><span data-stu-id="ae68f-148">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="ae68f-149">删除特定的属性，然后保留删除操作到 Exchange 服务器。</span><span class="sxs-lookup"><span data-stu-id="ae68f-149">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="ae68f-150">对 `updateProperty` 和 `removeProperty` 函数中 **saveAsync** 方法调用的回调。</span><span class="sxs-lookup"><span data-stu-id="ae68f-150">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|

```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method.
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

## <a name="see-also"></a><span data-ttu-id="ae68f-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ae68f-151">See also</span></span>

- [<span data-ttu-id="ae68f-152">保留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="ae68f-152">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="ae68f-153">初始化 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="ae68f-153">Initialize your Office Add-in</span></span>](../develop/initialize-add-in.md)