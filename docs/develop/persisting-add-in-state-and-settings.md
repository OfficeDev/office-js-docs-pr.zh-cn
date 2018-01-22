
# <a name="persisting-add-in-state-and-settings"></a>保留加载项状态和设置

Office 外接程序实质上是运行在浏览器控件的无状态环境中的 Web 应用程序。因此，加载项可能需要保留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能具有在下一次初始化时保存和重新加载所需的自定义设置或其他值（例如用户的首选视图或默认位置）。

为此，您可以：


- 使用适用于 Office 的 JavaScript API 的成员，在加载项类型决定的位置中存储的属性包中，它们将数据作为名称/值对存储。
    
- 使用基础浏览器控件（浏览器 cookies 或 HTML5 Web 存储）提供的技术（ [localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) 或 [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)）。
    
本文重点介绍如何使用适用于 Office 的 JavaScript API 保留外接程序状态。有关使用浏览器 Cookie 和 Web 存储的示例，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>使用适用于 Office 的 JavaScript API 保留加载项状态和设置


适用于 Office 的 JavaScript API 为在各个会话中保存外接程序状态提供了 [Settings](http://dev.office.com/reference/add-ins/shared/settings)、 [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 和 [CustomProperties](http://dev.office.com/reference/add-ins/outlook/CustomProperties) 对象，如下表中所述。在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) 相关联。



|**对象**|**外接程序类型支持**|**存储位置**|**Office 主机支持**|
|:-----|:-----|:-----|:-----|
|[设置](http://dev.office.com/reference/add-ins/shared/settings)|内容和任务窗格|加载项使用的文档、电子表格或演示文稿。内容和任务窗格加载项设置供创建它们的加载项使用，且能从保存它们的文档访问。**重要说明：**不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。你应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。|Word、Excel 或 PowerPoint **注意：**Project 2013 的任务窗格外接程序不支持 **Settings** API 存储外接程序状态或设置。但对于在 Project（及其他 Office 主机应用程序）中运行的外接程序，可以使用浏览器 Cookies 或 Web 存储等技术。有关详细信息，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。 |
|[RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings)|Outlook|安装了加载项的用户 Exchange 服务器邮箱。由于这些设置存储在用户的服务器邮箱中，因此当加载项运行在任何访问该用户邮箱的受支持客户端主机应用程序或浏览器的上下文中时，这些设置可随用户"漫游"且可供加载项使用。 Outlook 外接程序漫游设置仅供创建它们的外接程序使用，并且只能从安装了外接程序的邮箱使用。|Outlook|
|[CustomProperties](http://dev.office.com/reference/add-ins/outlook/CustomProperties)|Outlook|加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。|Outlook|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>运行时设置数据在内存中托管


在内部，通过  **Settings**、 **CustomProperties** 或 **RoamingSettings** 对象访问的属性包中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 对象，包含名称/值对。每个值的名称（键）必须为 **string**，且存储的值可为 JavaScript  **string**、 **number**、 **date** 或 **object**，但不能为  **function**。

本属性包结构示例包含三个已定义  **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

设置属性包在上一加载项会话中进行保存后，在加载项当前会话期间，可以在加载项初始化时或其初始化后的任何时间点加载设置属性包。在会话期间，设置使用与您所创建的该类设置相对应的对象（ **Settings**、 **CustomProperties** 或 **RoamingSettings**）的 **get**、 **set** 和 **remove** 方法完全托管在内存中。 


 >**重要信息**  要将在加载项当前会话期间所添加、更新或删除的任何内容保存到存储位置，您必须调用与用于处理该类设置的对象相对应的  **saveAsync** 方法。 **get**、 **set** 和 **remove** 方法仅用于设置属性包的内存副本。如果您的加载项在没有调用 **saveAsync** 的情况下关闭，则在该会话期间对设置所做的任何更改将丢失。 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>如何按文档保留内容和任务窗格加载项的加载项状态和设置


要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](http://dev.office.com/reference/add-ins/shared/settings) 对象及其方法。使用 **Settings** 对象的方法创建的属性包仅供创建它的内容或任务窗格加载项的实例使用，并且只能从保存它的文档使用。

**Settings** 对象自动加载为 [Document](http://dev.office.com/reference/add-ins/shared/document) 对象的一部分，并在任务窗格或内容加载项激活时可用。实例化 **Document** 对象后，你可以使用 **Document** 对象的 [settings ](../../reference/shared/document.settings.md)属性访问 **Settings** 对象。在该会话的生命周期中，你只能使用 **Settings.get**、**Settings.set** 和 **Settings.remove** 方法从属性包的内存副本中读取、写入或删除保留的设置和加载项状态。

由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法。


### <a name="creating-or-updating-a-setting-value"></a>创建或更新设置值

以下代码示例演示如何使用 [Settings.set](../../reference/shared/settings.set.md) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。


```
Office.context.document.settings.set('themeColor', 'green');
```

 如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。使用 **Settings.saveAsync** 方法可将新的或更新的设置保留到文档中。


### <a name="getting-the-value-of-a-setting"></a>获取设置的值

下面的示例演示如何使用 [Settings.get](../../reference/shared/settings.get.md) 方法获取名为“themeColor”的设置值。**get** 方法的唯一参数是设置的 _name_（区分大小写）。


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 **get** 方法返回之前为传入的设置 _name_ 保存的值。如果不存在该设置，那么方法返回 **null**。


### <a name="removing-a-setting"></a>删除设置

下面的示例演示如何使用 [Settings.remove](../../reference/shared/settings.removehandlerasync.md) 方法删除名为“themeColor”的设置。**remove** 方法的唯一参数是设置的 _name_（区分大小写）。


```
Office.context.document.settings.remove('themeColor');
```

如果不存在该设置，则不执行任何操作。使用 **Settings.saveAsync** 方法可保留文档中设置的删除操作。


### <a name="saving-your-settings"></a>保存设置

若要保存当前会话中加载项对设置属性包内存副本所做的任何添加、更改或删除操作，必须调用 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法，将它们存储在文档中。**saveAsync** 方法的唯一参数是使用单个参数的回调函数 _callback_。 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

完成此操作后，将执行作为  **callback** 参数传入 _saveAsync_ 方法中的匿名函数。回调的 _asyncResult_ 参数提供对包含操作状态的 **AsyncResult** 对象的访问。在此示例中，函数将检查 **AsyncResult.status** 属性，以查看保存操作成功还是失败，然后在加载项页中显示结果。


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>如何将 Outlook 外接程序用户邮箱中的设置保存为漫游设置


Outlook 外接程序可以使用 [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 对象保存特定于用户邮箱的外接程序状态和设置数据。此数据只能由该 Outlook 外接程序代表运行外接程序的用户访问。此数据存储在用户的 Exchange Server 邮箱上，并可供该用户登录其帐户并运行 Outlook 外接程序时访问。


### <a name="loading-roaming-settings"></a>加载漫游设置


Outlook 外接程序通常在 [Office.initialize](../../reference/shared/office.initialize.md) 事件处理程序中加载漫游设置。以下 JavaScript 代码示例演示了如何加载现有漫游设置。


```
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


紧接着前面的示例，下面的  `setAppSetting` 函数演示如何使用 [RoamingSettings.set](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 方法通过当天的日期设置或更新名为 `cookie` 的设置。然后使用 [RoamingSettings.saveAsync](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 方法将所有漫游设置保存回 Exchange Server。


```
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

**saveAsync** 方法可异步保存漫游设置，并采用一个可选的回调函数。此代码示例将名为 `saveMyAppSettingsCallback` 的回调函数传递到 **saveAsync** 方法。异步调用返回时，`saveMyAppSettingsCallback` 函数的 _asyncResult_ 参数提供对 [AsyncResult](http://dev.office.com/reference/add-ins/outlook/simple-types) 对象的访问权限，该对象用于确定通过 **AsyncResult.status** 属性的操作是成功还是失败。


### <a name="removing-a-roaming-setting"></a>删除漫游设置


进一步展开前面的示例，以下  `removeAppSetting` 函数演示了如何使用 [RoamingSettings.remove](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>如何按项目将 Outlook 外接程序的设置保存为自定义属性


自定义属性允许 Outlook 外接程序存储其使用的有关项目的信息。例如，如果 Outlook 外接程序根据邮件中的会议建议创建约会，则可以使用自定义属性存储创建了会议的事实。这确保了如果再次打开邮件，Outlook 外接程序不再可供创建约会。

在您将自定义属性用于特定邮件、约会或会议请求项目之前，必须通过调用  [Item](../../reference/outlook/Office.context.mailbox.item.md) 对象的 **loadCustomPropertiesAsync** 方法将属性加载到内存中。如果为当前项目设置了任何自定义属性，此时会从 Exchanger Server 加载这些属性。在您加载了属性以后，可以使用 [CustomProperties](http://dev.office.com/reference/add-ins/outlook/CustomProperties) 对象的 [set](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 和 **get** 方法添加、更新和检索内存中的属性。要保存对于项目的自定义属性所做的任何更改，必须使用 [saveAsync](http://dev.office.com/reference/add-ins/outlook/CustomProperties) 方法在 Exchanger Server上保留对项目所做的更改。


### <a name="custom-properties-example"></a>自定义属性示例

下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的函数。可以将此示例用作使用自定义属性的 Outlook 外接程序的起点。 

使用这些函数的 Outlook 外接程序通过对  `_customProps` 变量调用 **get** 方法来检索任何自定义属性，如下面的示例所示。




```
var property = _customProps.get("propertyName");
```

此示例包括以下函数：



|**函数名称**|**说明**|
|:-----|:-----|
| `Office.initialize`|从 Exchange 服务器初始化外接程序并加载当前项目的自定义属性。|
| `customPropsCallback`|获取从 Exchange 服务器返回的自定义属性并将其保存以供后续之用。|
| `updateProperty`|设置或更新特定属性，然后将更改保存到 Exchange 服务器。|
| `removeProperty`|删除特定的属性，然后保留删除操作到 Exchange 服务器。|
| `saveCallback`|对  `updateProperty` 和 `removeProperty` 函数中 **saveAsync** 方法调用的回调。|



```
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


## <a name="additional-resources"></a>其他资源



- [了解适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
