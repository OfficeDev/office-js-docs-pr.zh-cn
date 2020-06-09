---
title: 管理 Outlook 外接程序的状态和设置
description: 了解如何保留 Outlook 外接程序的外接程序状态和设置。
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 796c7b38f8c85a5680c9b7de43297c754a0ebc1b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609060"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>管理 Outlook 外接程序的状态和设置

> [!NOTE]
> 阅读本文之前，请查看本文档的 "**核心概念**" 一节中的 "[保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)"。

对于 Outlook 外接程序，Office JavaScript API 提供[RoamingSettings](/javascript/api/outlook/office.roamingsettings)和[CustomProperties](/javascript/api/outlook/office.customproperties)对象，以在各会话之间保存外接程序状态，如下表所述。 在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](../reference/manifest/id.md) 相关联。

|**对象**|**存储位置**|
|:-----|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|安装了加载项的用户 Exchange 服务器邮箱。 由于这些设置存储在用户的服务器邮箱中，因此当加载项运行在任何访问该用户邮箱的受支持客户端主机应用程序或浏览器的上下文中时，这些设置可随用户“漫游”且可供加载项使用。<br/><br/> Outlook 外接程序漫游设置仅供创建它们的外接程序使用，并且只能从安装了外接程序的邮箱使用。|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>如何将 Outlook 加载项用户邮箱中的设置保存为漫游设置

Outlook 加载项可以使用 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) 对象保存特定于用户邮箱的加载项状态和设置数据。 仅代表用户运行该加载项的 Outlook 加载项才可访问此数据。 这些数据将存储在用户的 Exchange Server 邮箱上，并且在用户登录到其帐户并运行 Outlook 加载项时可访问这些数据。

### <a name="loading-roaming-settings"></a>加载漫游设置

Outlook 外接程序通常在 [Office.initialize](/javascript/api/office) 事件处理程序中加载漫游设置。以下 JavaScript 代码示例演示了如何加载现有漫游设置。

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

### <a name="creating-or-assigning-a-roaming-setting"></a>创建或分配漫游设置

紧接着前面的示例，下面的  `setAppSetting` 函数演示如何使用 [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) 方法通过当天的日期设置或更新名为 `cookie` 的设置。然后使用 [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) 方法将所有漫游设置保存回 Exchange Server。

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

**saveAsync** 方法将异步保存漫游设置，并采用可选回调函数。 此代码示例会将名为 `saveMyAppSettingsCallback` 的回调函数传递给 **saveAsync** 方法。 当异步调用返回时，`saveMyAppSettingsCallback` 函数的 _asyncResult_ 参数提供对 [AsyncResult](/javascript/api/office/office.asyncresult) 对象的访问权限，你可以使用该对象通过 **AsyncResult.status** 属性确定操作是否成功。

### <a name="removing-a-roaming-setting"></a>删除漫游设置

进一步展开前面的示例，以下  `removeAppSetting` 函数演示了如何使用 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>如何按项目将 Outlook 外接程序的设置保存为自定义属性

自定义属性允许 Outlook 外接程序存储其使用的有关项目的信息。例如，如果 Outlook 外接程序根据邮件中的会议建议创建约会，则可以使用自定义属性存储创建了会议的事实。这确保了如果再次打开邮件，Outlook 外接程序不再可供创建约会。

在您将自定义属性用于特定邮件、约会或会议请求项目之前，必须通过调用  [Item](/javascript/api/outlook/office.mailbox) 对象的 **loadCustomPropertiesAsync** 方法将属性加载到内存中。如果为当前项目设置了任何自定义属性，此时会从 Exchanger Server 加载这些属性。在您加载了属性以后，可以使用 [CustomProperties](/javascript/api/outlook/office.customproperties#set-name--value-) 对象的 [set](/javascript/api/outlook/office.roamingsettings) 和 **get** 方法添加、更新和检索内存中的属性。要保存对于项目的自定义属性所做的任何更改，必须使用 [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) 方法在 Exchanger Server上保留对项目所做的更改。

### <a name="custom-properties-example"></a>自定义属性示例

下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的函数。可以将此示例用作使用自定义属性的 Outlook 外接程序的起点。 

使用这些函数的 Outlook 加载项通过对 `_customProps` 变量调用 **get** 方法来检索任何自定义属性，如下面的示例所示。

```js
var property = _customProps.get("propertyName");
```

此示例包括以下函数：

|**函数名称**|**说明**|
|:-----|:-----|
| `Office.initialize`|从 Exchange 服务器初始化外接程序并加载当前项目的自定义属性。|
| `customPropsCallback`|获取从 Exchange 服务器返回的自定义属性并将其保存以供后续之用。|
| `updateProperty`|设置或更新特定属性，然后将更改保存到 Exchange 服务器。|
| `removeProperty`|删除特定的属性，然后保留删除操作到 Exchange 服务器。|
| `saveCallback`|对 `updateProperty` 和 `removeProperty` 函数中 **saveAsync** 方法调用的回调。|

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

## <a name="see-also"></a>另请参阅

- [保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)
- [初始化 Office 加载项](../develop/initialize-add-in.md)