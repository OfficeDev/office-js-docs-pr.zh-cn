---
title: 暂留加载项状态和设置
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 6092a93751825561f83cfea1671fe59e273f6142
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617015"
---
# <a name="persisting-add-in-state-and-settings"></a>暂留加载项状态和设置

Office 加载项实质上是在浏览器控件的无状态环境中运行的 Web 应用。因此，加载项可能需要暂留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能有需要在下一次初始化时保存和重新加载的自定义设置或其他值（如用户的首选视图或默认位置）。为此，可以执行下列操作：

- 使用适用于 Office 的 JavaScript API 成员，将数据存储为：
    -  在依赖加载项类型的位置上存储的属性包中的名称-数值对。
    -  在文档中存储的自定义 XML。

- 使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。

本文重点介绍如何使用适用于 Office 的 JavaScript API 保留外接程序状态。有关使用浏览器 Cookie 和 Web 存储的示例，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>使用适用于 Office 的 JavaScript API 保留加载项状态和设置

适用于 Office 的 JavaScript API 为在各个会话中保存外接程序状态提供了 [Settings](/javascript/api/office/office.settings)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) 和 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象，如下表中所述。在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](/office/dev/add-ins/reference/manifest/id) 相关联。

|**对象**|**外接程序类型支持**|**存储位置**|**Office 主机支持**|
|:-----|:-----|:-----|:-----|
|[Settings](/javascript/api/office/office.settings)|内容和任务窗格|加载项要使用的文档、电子表格或演示文稿。内容和任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**重要说明：** 不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。您应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。|Word、Excel 或 PowerPoint<br/><br/> **注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。不过，对于在 Project（及其他 Office 主机应用）中运行的加载项，可以使用浏览器 Cookie 或 Web 存储等技术。若要详细了解这些技术，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。 |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|安装了加载项的用户 Exchange 服务器邮箱。由于这些设置存储在用户的服务器邮箱中，因此如果加载项在任何访问用户邮箱的受支持客户端主机应用或浏览器的上下文中运行，这些设置可随用户“漫游”，且可供加载项使用。<br/><br/> Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|任务窗格|加载项要使用的文档、电子表格或演示文稿。任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**重要说明：** 请勿将密码和其他敏感的个人身份信息 (PII) 存储在自定义 XML 部分中。虽然保存的数据对最终用户不可见，但它存储为文档的一部分，可通过直接读取文档的文件格式进行访问。应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在服务器上，且服务器将加载项托管为用户保护资源。|Word（使用 Office JavaScript 常见 API）、Excel（使用主机专用 Excel JavaScript API）|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>设置数据在运行时托管在内存中

> [!NOTE]
> 下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。 主机专用 Excel JavaScript API 还提供对自定义设置的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel SettingCollection](/javascript/api/excel/excel.settingcollection)。

在内部，通过 **Settings**、 **CustomProperties** 或 **RoamingSettings** 对象访问的属性包中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 对象，包含名称/值对。 每个值的名称（键）必须为 **string**，且存储的值可为 JavaScript **string**、**number**、**date** 或 **object**，但不能为 **function**。

本属性包结构示例包含三个已定义 **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

在前一个加载项会话中保存设置属性包之后，可以在加载项的当前会话中初始化加载项时或在之后的任何时间加载该设置属性包。 在会话过程中，将会使用与所创建的种类设置（**Settings**、**CustomProperties** 或 **RoamingSettings**）对应的对象 的 **get**、**set** 和 **remove** 方法，将设置整体托管到内存中。 


> [!IMPORTANT]
> 若要将在加载项的当前会话中所做的任何添加、更新或删除暂留到存储位置，必须调用与此类设置搭配使用的对应对象的 **saveAsync** 方法。 **get**、**set** 和 **remove** 方法只可在设置属性包的内存副本上运行。 如果加载项在未调用 **saveAsync** 的情况下关闭，则在该会话过程中对设置所做的任何更改都会丢失。 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>如何按文档暂留内容和任务窗格加载项的加载项状态和设置


要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](/javascript/api/office/office.settings) 对象及其方法。使用 **Settings** 对象的方法创建的属性包仅供创建它的内容或任务窗格加载项的实例使用，并且只能从保存它的文档使用。

**Settings** 对象将作为 [Document](/javascript/api/office/office.document) 对象的一部分自动加载，并且在任务窗格或内容加载项激活时可用。 **Document** 对象实例化之后，可以通过 **Document** 对象的 [settings](/javascript/api/office/office.document#settings) 属性访问 **Settings** 对象。 在会话的生存期，可以只使用 **Settings.get**、**Settings.set** 和 **Settings.remove** 方法读取、写入或删除属性包内存副本中暂留的设置和加载项状态。

由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法。


### <a name="creating-or-updating-a-setting-value"></a>创建或更新设置值

以下代码示例演示如何使用 [Settings.set](/javascript/api/office/office.settings#set-name--value-) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。


```js
Office.context.document.settings.set('themeColor', 'green');
```

 如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。使用 **Settings.saveAsync** 方法可将新的或更新的设置保留到文档中。


### <a name="getting-the-value-of-a-setting"></a>获取设置的值

下面的示例演示如何使用 [Settings.get](/javascript/api/office/office.settings#get-name-) 方法获取名为“themeColor”的设置值。**get** 方法的唯一参数是设置的 _name_（区分大小写）。


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 **get** 方法返回之前为传入的设置 _name_ 保存的值。如果不存在该设置，那么方法返回 **null**。


### <a name="removing-a-setting"></a>删除设置

下面的示例演示如何使用 [Settings.remove](/javascript/api/office/office.settings#remove-name-) 方法删除名为“themeColor”的设置。**remove** 方法的唯一参数是设置的 _name_（区分大小写）。


```js
Office.context.document.settings.remove('themeColor');
```

如果不存在该设置，则不执行任何操作。 使用 **Settings.saveAsync** 方法可保留文档中设置的删除操作。


### <a name="saving-your-settings"></a>保存设置

若要保存当前会话中加载项对设置属性包内存副本所做的任何添加、更改或删除操作，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法，将它们存储在文档中。**saveAsync** 方法的唯一参数是使用单个参数的回调函数 _callback_。 


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

作为 _callback_ 参数传入 **saveAsync** 方法的匿名函数在操作完成时执行。 回调的 _asyncResult_ 参数提供对包含操作状态的 **AsyncResult** 对象的访问权限。 在此示例中，函数将检查 **AsyncResult.status** 属性，以确定保存操作是否成功，然后在加载项页显示结果。

## <a name="how-to-save-custom-xml-to-the-document"></a>如何将自定义 XML 保存到文档

> [!NOTE]
> 此部分是在 Word 中支持的 Office 常见 JavaScript API 上下文中介绍的自定义 XML 部分。 主机专用 Excel JavaScript API 还提供对自定义 XML 部分的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart)。

如果需要存储的信息超过文档设置的大小限制或有结构化字符，还有一个额外的存储选项。 可以在 Word 的任务窗格加载项中暂留自定义 XML 标记（对于 Excel，但请参阅本节顶部的注释）。 在 Word 中，可以使用 [CustomXmlPart](/javascript/api/office/office.customxmlpart) 对象及其方法（同样，请参阅上面的 Excel 注释）。 以下代码将创建自定义 XML 部件，并在页面的 divs 中显示其 ID 及内容。 请注意，XML 字符串中必须有一个 `xmlns` 属性。

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

若要检索自定义 XML 部分，请使用 [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) 方法，但 ID 是在创建 XML 部分时生成的 GUID，因此编码时无法知道 ID 是什么。 因此，最好是在创建 XML 部分时，立即将 XML 部分的 ID 存储为设置，并为它提供容易记住的密钥。 下面的方法展示了如何执行此操作。 （不过，处理自定义设置时，请参阅本文的前面部分，以详细了解相关信息和最佳做法）。

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

下面的代码展示了如何通过先从设置中获取 ID 来检索 XML 部分。

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```


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

**saveAsync** 方法将异步保存漫游设置，并采用可选回调函数。 此代码示例会将名为 `saveMyAppSettingsCallback` 的回调函数传递给 **saveAsync** 方法。 当异步调用返回时，`saveMyAppSettingsCallback` 函数的 _asyncResult_ 参数提供对 [AsyncResult](/javascript/api/outlook) 对象的访问权限，你可以使用该对象通过 **AsyncResult.status** 属性确定操作是否成功。


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

- [了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)
- [Outlook 加载项](/outlook/add-ins/)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
