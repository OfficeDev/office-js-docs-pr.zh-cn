---
title: 保留加载项状态和设置
description: 了解如何在浏览器控件的无状态环境中运行的 Office 外接程序 Web 应用程序中保存数据。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e2018e5ecf419744257cdceac31b8b1688fa65ff
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810006"
---
# <a name="persist-add-in-state-and-settings"></a>保留加载项状态和设置

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location.
To do that, you can:

- 使用 Office JavaScript API 的成员，该 API 将数据存储为以下任一：
  - 在依赖加载项类型的位置上存储的属性包中的名称-数值对。
  - 在文档中存储的自定义 XML。

- 使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。
    > [!NOTE]
    > 某些浏览器或用户的浏览器设置可能会阻止基于浏览器的存储技术。 应按照 [使用 Web 存储 API](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API) 中所述测试可用性。

本文重点介绍如何使用 Office JavaScript API 将加载项状态保存到当前文档。 如果需要跨文档保留状态，例如跟踪打开的任何文档的用户首选项，则需要使用其他方法。 例如，可以使用 [SSO](use-sso-to-get-office-signed-in-user-token.md) 获取用户标识，然后将用户 ID 及其设置保存到联机数据库。

## <a name="persist-add-in-state-and-settings-with-the-office-javascript-api"></a>使用 Office JavaScript API 保留加载项状态和设置

Office JavaScript API 提供 [Settings](/javascript/api/office/office.settings)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) 和 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象，用于跨会话保存加载项状态，如下表所述。 在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](/javascript/api/manifest/id) 相关联。

|Object|加载项类型支持|存储位置|Office 应用程序支持|
|:-----|:-----|:-----|:-----|
|[设置](/javascript/api/office/office.settings)|-内容<br>- 任务窗格|加载项使用的文档、电子表格或演示文稿。 内容和任务窗格加载项设置供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|-词<br>-Excel<br>-幻灯片<br/><br/> **注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。 但是，对于在 Project (中运行的加载项以及其他 Office 客户端应用程序) 可以使用浏览器 Cookie 或 Web 存储等技术。 有关详细信息，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。 |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|mail|安装了加载项的用户 Exchange 服务器邮箱。 由于这些设置存储在用户的服务器邮箱中，因此它们可以与用户一起“漫游”，并且当加载项在访问该用户的邮箱的任何受支持的 Office 客户端应用程序或浏览器的上下文中运行时，加载项可供使用。<br/><br/> Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|mail|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|任务窗格|The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|- 使用 Office JavaScript 通用 API 的 Word () <br>- 使用特定于应用程序的 Excel JavaScript API 的 Excel () |

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>设置数据在运行时托管在内存中

> [!NOTE]
> 下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。 特定于应用程序的 Excel JavaScript API 还提供对自定义设置的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel SettingCollection](/javascript/api/excel/excel.settingcollection)。

在内部，使用 `Settings`、 `CustomProperties`或 `RoamingSettings` 对象访问的属性包中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 对象，其中包含名称/值对。 每个值 (键) 的名称必须是 `string`，并且存储的值可以是 JavaScript `string`、 `number`、 `date`或 `object`，但不能是 **函数**。

本属性包结构示例包含三个已定义 **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

在前一个加载项会话中保存设置属性包之后，可以在加载项的当前会话中初始化加载项时或在之后的任何时间加载该设置属性包。 在会话期间，使用 `get`对象的 、 `set`和 `remove` 方法在内存中完全管理设置，这些对象对应于要创建的设置类型， (**Settings**、 **CustomProperties** 或 **RoamingSettings**) 。

> [!IMPORTANT]
> 若要将加载项当前会话期间进行的任何添加、更新或删除保存到存储位置，必须调用 `saveAsync` 用于处理此类设置的相应对象的 方法。 `get`、 `set`和 `remove` 方法仅在设置属性包的内存中副本上运行。 如果加载项在未调用 `saveAsync`的情况下关闭，则在该会话期间对设置所做的任何更改都将丢失。

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>如何按文档暂留内容和任务窗格加载项的加载项状态和设置

要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](/javascript/api/office/office.settings) 对象及其方法。 使用 对象方法 `Settings` 创建的属性包仅适用于创建它的内容或任务窗格加载项的实例，并且只能从保存它的文档中使用。

对象 `Settings` 作为 [Document](/javascript/api/office/office.document) 对象的一部分自动加载，并在激活任务窗格或内容加载项时可用。 实例 `Document` 化对象后，可以使用 对象的 `Settings` [settings](/javascript/api/office/office.document#office-office-document-settings-member) 属性 `Document` 访问 对象。 在会话的生存期内，只需使用 `Settings.get`、 `Settings.set`和 `Settings.remove` 方法来读取、写入或删除属性包的内存中副本中的持久化设置和加载项状态。

由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) 方法。

### <a name="creating-or-updating-a-setting-value"></a>创建或更新设置值

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。 `Settings.saveAsync`使用 方法可将新的或更新的设置保存到文档中。

### <a name="getting-the-value-of-a-setting"></a>获取设置的值

下面的示例演示如何使用 [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) 方法获取名为"themeColor"的设置值。 方法的唯一 `get` 参数是设置的区分大小写 _的名称_ 。

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 方法 `get` 返回之前为传入的设置 _名称_ 保存的值。 如果不存在该设置，那么方法返回 **null**。

### <a name="removing-a-setting"></a>删除设置

下面的示例演示如何使用 [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) 方法删除名为"themeColor"的设置。 方法的唯一 `remove` 参数是设置的区分大小写 _的名称_ 。

```js
Office.context.document.settings.remove('themeColor');
```

如果不存在该设置，则不执行任何操作。 `Settings.saveAsync`使用 方法持久删除文档中的设置。

### <a name="saving-your-settings"></a>保存设置

若要保存当前会话中加载项对设置属性包内存副本所做的任意添加、更改或删除操作，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) 方法将它们存储在文档中。 方法的唯一 `saveAsync` 参数是 _callback_，它是具有单个参数的回调函数。

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

在操作完成时，将作为 _回调_ 参数传入`saveAsync`方法的匿名函数执行。 回调的 _asyncResult_ 参数提供对 `AsyncResult` 包含操作状态的对象的访问权限。 在此示例中，函数检查 `AsyncResult.status` 属性以查看保存操作是成功还是失败，然后在加载项的页面中显示结果。

## <a name="how-to-save-custom-xml-to-the-document"></a>如何将自定义 XML 保存到文档

> [!NOTE]
> 此部分是在 Word 中支持的 Office 常见 JavaScript API 上下文中介绍的自定义 XML 部分。 特定于应用程序的 Excel JavaScript API 还提供对自定义 XML 部件的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart)。

当需要存储的信息超过文档“设置”的大小限制或具有结构化字符的信息时，还有一个额外的存储选项。 可以在 Word 的任务窗格加载项中暂留自定义 XML 标记（对于 Excel，但请参阅本节顶部的注释）。 在 Word 中，可以使用 [CustomXmlPart](/javascript/api/office/office.customxmlpart) 对象及其方法（同样，请参阅上面的 Excel 注释）。 以下代码将创建自定义 XML 部件，并在页面的 divs 中显示其 ID 及内容。 请注意，XML 字符串中必须有一个 `xmlns` 属性。

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

若要检索自定义 XML 部分，请使用 [getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) 方法，但 ID 是在创建 XML 部分时生成的 GUID，因此编码时无法知道 ID 是什么。 因此，最好是在创建 XML 部分时，立即将 XML 部分的 ID 存储为设置，并为它提供容易记住的密钥。 下面的方法展示了如何执行此操作。  (但请参阅本文的前面部分，了解使用自定义设置时的详细信息和最佳做法。) 

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>如何在 Outlook 加载项中保存设置

有关如何在 Outlook 外接程序中保存设置的信息，请参阅 [管理 Outlook 外接程序的状态和设置](../outlook/manage-state-and-settings-outlook.md)。

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Outlook 加载项](../outlook/outlook-add-ins-overview.md)
- [管理 Outlook 加载项的状态和设置](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
