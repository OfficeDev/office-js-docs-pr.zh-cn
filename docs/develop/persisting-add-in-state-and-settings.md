---
title: 暂留加载项状态和设置
description: 了解如何在浏览器控件的无状态环境中保存运行的 Office 外接程序 web 应用程序中的数据。
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: 81f149bdff540b236252a02a0c368799a11fed10
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609399"
---
# <a name="persisting-add-in-state-and-settings"></a>暂留加载项状态和设置

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office 加载项实质上是在浏览器控件的无状态环境中运行的 Web 应用。因此，加载项可能需要暂留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能有需要在下一次初始化时保存和重新加载的自定义设置或其他值（如用户的首选视图或默认位置）。为此，可以执行下列操作：

- 使用 Office JavaScript API 的成员，将数据存储为以下任意一种：
    -  在依赖加载项类型的位置上存储的属性包中的名称-数值对。
    -  在文档中存储的自定义 XML。

- 使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。

本文重点介绍如何使用 Office JavaScript API 来保留加载项状态。 有关使用浏览器 Cookie 和 Web 存储的示例，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a>使用 Office JavaScript API 保留加载项状态和设置

Office JavaScript API 提供了[设置](/javascript/api/office/office.settings)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)和[CustomProperties](/javascript/api/outlook/office.customproperties)对象，用于按下表所述在会话中保存外接程序状态。 在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](../reference/manifest/id.md) 相关联。

|**对象**|**外接程序类型支持**|**存储位置**|**Office 主机支持**|
|:-----|:-----|:-----|:-----|
|[Settings](/javascript/api/office/office.settings)|内容和任务窗格|加载项使用的文档、电子表格或演示文稿。 内容和任务窗格加载项设置供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**重要说明：** 不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。您应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。|Word、Excel 或 PowerPoint<br/><br/> **注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。不过，对于在 Project（及其他 Office 主机应用）中运行的加载项，可以使用浏览器 Cookie 或 Web 存储等技术。若要详细了解这些技术，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。 |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|安装了加载项的用户 Exchange 服务器邮箱。 由于这些设置存储在用户的服务器邮箱中，因此当加载项运行在任何访问该用户邮箱的受支持客户端主机应用程序或浏览器的上下文中时，这些设置可随用户“漫游”且可供加载项使用。<br/><br/> Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|任务窗格|加载项要使用的文档、电子表格或演示文稿。任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**重要说明：** 请勿将密码和其他敏感的个人身份信息 (PII) 存储在自定义 XML 部分中。虽然保存的数据对最终用户不可见，但它存储为文档的一部分，可通过直接读取文档的文件格式进行访问。应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在服务器上，且服务器将加载项托管为用户保护资源。|Word（使用 Office JavaScript 常见 API）、Excel（使用主机专用 Excel JavaScript API）|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>设置数据在运行时托管在内存中

> [!NOTE]
> 下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。 主机专用 Excel JavaScript API 还提供对自定义设置的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel SettingCollection](/javascript/api/excel/excel.settingcollection)。

在内部，使用、或对象访问的属性包中的数据 `Settings` `CustomProperties` `RoamingSettings` 存储为序列化的 JavaScript 对象表示法（JSON）对象，其中包含名称/值对。 每个值的名称（键）都必须是 `string` ，并且存储的值可以是 JavaScript `string` 、 `number` 、 `date` 或 `object` ，但不能是**函数**。

本属性包结构示例包含三个已定义 **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

在前一个加载项会话中保存设置属性包之后，可以在加载项的当前会话中初始化加载项时或在之后的任何时间加载该设置属性包。 在会话过程中，将使用 `get` 与所创建的设置类型相对应的对象的、和方法在完全内存中管理设置 `set` `remove` （**Settings**、 **CustomProperties**或**RoamingSettings**）。


> [!IMPORTANT]
> 若要将在外接程序的当前会话过程中所做的任何添加、更新或删除操作保存到存储位置，必须调用 `saveAsync` 与该类型的设置一起使用的相应对象的方法。 `get`、 `set` 和 `remove` 方法仅在设置属性包的内存中副本上运行。 如果外接程序在未呼叫的情况 `saveAsync` 下关闭，则在该会话期间对设置所做的任何更改都将丢失。


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>如何按文档暂留内容和任务窗格加载项的加载项状态和设置


要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](/javascript/api/office/office.settings) 对象及其方法。 使用对象的方法创建的属性包 `Settings` 仅可用于创建它的内容或任务窗格外接程序的实例，并且只能从保存它的文档中获取。

`Settings`对象将作为[Document](/javascript/api/office/office.document)对象的一部分自动加载，并在激活任务窗格或内容加载项时可用。 在 `Document` 实例化对象之后，可以 `Settings` 使用对象的[settings](/javascript/api/office/office.document#settings)属性访问对象 `Document` 。 在会话的生存期期间，您只需使用 `Settings.get` 、 `Settings.set` 和方法，即可 `Settings.remove` 从属性包的内存中副本中读取、写入或删除保留的设置和加载项状态。

由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法。


### <a name="creating-or-updating-a-setting-value"></a>创建或更新设置值

以下代码示例演示如何使用 [Settings.set](/javascript/api/office/office.settings#set-name--value-) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。


```js
Office.context.document.settings.set('themeColor', 'green');
```

 如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。 使用 `Settings.saveAsync` 方法可将新的或更新的设置保存到文档中。


### <a name="getting-the-value-of-a-setting"></a>获取设置的值

下面的示例演示如何使用 [Settings.get](/javascript/api/office/office.settings#get-name-) 方法获取名为"themeColor"的设置值。 方法的唯一参数 `get` 是设置的区分大小写的_名称_。


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 该 `get` 方法返回之前为传入的设置_名称_保存的值。 如果不存在该设置，那么方法返回 **null**。


### <a name="removing-a-setting"></a>删除设置

下面的示例演示如何使用 [Settings.remove](/javascript/api/office/office.settings#remove-name-) 方法删除名为"themeColor"的设置。 方法的唯一参数 `remove` 是设置的区分大小写的_名称_。


```js
Office.context.document.settings.remove('themeColor');
```

如果不存在该设置，则不执行任何操作。 使用 `Settings.saveAsync` 方法可将设置从文档中永久删除。


### <a name="saving-your-settings"></a>保存设置

若要保存当前会话中加载项对设置属性包内存副本所做的任意添加、更改或删除操作，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法将它们存储在文档中。 该方法的唯一参数 `saveAsync` 是_callback_，它是一个具有单个参数的回调函数。 


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

在 `saveAsync` 操作完成时，会执行匿名函数作为_callback_参数传入方法。 此回调的_asyncResult_参数提供对 `AsyncResult` 包含操作状态的对象的访问权限。 在此示例中，函数检查 `AsyncResult.status` 属性以查看保存操作是成功还是失败，然后在加载项页面中显示结果。

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>如何在 Outlook 加载项中保存设置

有关如何在 Outlook 外接程序中保存设置的信息，请参阅[Manage state and settings For outlook 外接程序](../outlook/manage-state-and-settings-outlook.md)。


## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Outlook 加载项](../outlook/outlook-add-ins-overview.md)
- [管理 Outlook 外接程序的状态和设置](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
