---
title: 保留加载项状态和设置
description: 了解如何将数据保留Office浏览器控件的无状态环境中运行的外接程序 Web 应用程序中。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: b09520d997354e5acc7ec68e3408d97230e4c9dc
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743680"
---
# <a name="persist-add-in-state-and-settings"></a>保留加载项状态和设置

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office 加载项实质上是在浏览器控件的无状态环境中运行的 Web 应用。因此，加载项可能需要暂留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能有需要在下一次初始化时保存和重新加载的自定义设置或其他值（如用户的首选视图或默认位置）。为此，可以执行下列操作：

- 使用存储数据的 Office JavaScript API 的成员：
  - 在依赖加载项类型的位置上存储的属性包中的名称-数值对。
  - 在文档中存储的自定义 XML。

- 使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。
    > [!NOTE]
    > 某些浏览器或用户的浏览器设置可能会阻止基于浏览器的存储技术。 应测试可用性，如[使用 Web 存储 API 中记录](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API)。

本文重点介绍如何使用 Office JavaScript API 将外接程序状态保留到当前文档。 如果需要跨文档保留状态，例如跨文档打开的任何文档跟踪用户首选项，则需要使用不同的方法。 例如，您可以使用 [SSO](use-sso-to-get-office-signed-in-user-token.md) 获取用户标识，然后将用户 ID 及其设置保存到联机数据库中。

## <a name="persist-add-in-state-and-settings-with-the-office-javascript-api"></a>使用 JavaScript API 保留Office状态和设置

JavaScript API Office提供了用于[跨](/javascript/api/office/office.settings)设置保存外接程序状态（如下表所述）的 设置、[RoamingSettings](/javascript/api/outlook/office.roamingsettings) 和 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象。 在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](../reference/manifest/id.md) 相关联。

|**对象**|**外接程序类型支持**|**存储位置**|**Office应用程序支持**|
|:-----|:-----|:-----|:-----|
|[Settings](/javascript/api/office/office.settings)|内容和任务窗格|加载项使用的文档、电子表格或演示文稿。 内容和任务窗格加载项设置供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**重要说明：** 不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。您应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。|Word、Excel 或 PowerPoint<br/><br/> **注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。 但是，对于在 Project (中运行的外接程序以及其他 Office 客户端) 可以使用浏览器 Cookie 或 Web 存储等技术。 有关详细信息，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。 |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|安装了加载项的用户 Exchange 服务器邮箱。 由于这些设置存储在用户的服务器邮箱中，因此它们可以随用户一起"漫游"，并且可在外接程序在任何访问该用户邮箱的受支持 Office 客户端应用程序或浏览器的上下文中运行时使用。<br/><br/> Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|任务窗格|加载项要使用的文档、电子表格或演示文稿。任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。<br/><br/>**重要说明：** 请勿将密码和其他敏感的个人身份信息 (PII) 存储在自定义 XML 部分中。虽然保存的数据对最终用户不可见，但它存储为文档的一部分，可通过直接读取文档的文件格式进行访问。应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在服务器上，且服务器将加载项托管为用户保护资源。|Word (JavaScript API Office应用程序) Excel (JavaScript API Excel JavaScript|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>设置数据在运行时托管在内存中

> [!NOTE]
> 下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。 特定于应用程序的应用程序Excel JavaScript API 还提供对自定义设置的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel SettingCollection](/javascript/api/excel/excel.settingcollection)。

`Settings``CustomProperties``RoamingSettings`在内部，使用 、 或 对象访问的属性包中数据存储为序列化的 JavaScript 对象表示法 (JSON) 对象，其中包含名称/值对。 每个 (`string`键) 的名称必须为 ，而存储的值可以是 JavaScript`string``number`、、 `date``object`或 ，但不能 **是函数**。

本属性包结构示例包含三个已定义 **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

在前一个加载项会话中保存设置属性包之后，可以在加载项的当前会话中初始化加载项时或在之后的任何时间加载该设置属性包。 `get``set``remove`在会话期间，使用 对象的 、 和 方法完全在内存中管理设置，这些对象对应于要创建 (**设置**、**CustomProperties** 或 **RoamingSettings**) 的设置类型。

> [!IMPORTANT]
> 若要将加载项当前 `saveAsync` 会话期间执行的任何添加、更新或删除操作保留到存储位置，必须调用用于处理此类设置的相应对象的 方法。 、 `get``set`和 `remove` 方法仅对设置属性包的内存副本进行操作。 如果加载项在未调用的情况下关闭， `saveAsync`则在此会话期间对设置进行的任何更改都将丢失。

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>如何按文档暂留内容和任务窗格加载项的加载项状态和设置

要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](/javascript/api/office/office.settings) 对象及其方法。 使用 对象的方法 `Settings` 创建的属性包仅可用于创建对象的内容或任务窗格加载项的实例，并且只能从保存它的文档使用。

该对象 `Settings` 将自动作为 [Document](/javascript/api/office/office.document) 对象的一部分加载，并且可在任务窗格或内容外接程序激活时使用。 `Document`实例化对象后，可以使用`Settings`对象的 [settings](/javascript/api/office/office.document#office-office-document-settings-member) 属性访问`Document`该对象。 在会话的`Settings.get``Settings.set``Settings.remove`生命周期中，只能使用 、 和 方法读取、写入或删除属性包的内存副本中的保留设置和加载项状态。

由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) 方法。

### <a name="creating-or-updating-a-setting-value"></a>创建或更新设置值

以下代码示例演示如何使用 [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。

```js
Office.context.document.settings.set('themeColor', 'green');
```

 如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。 `Settings.saveAsync`使用 方法将新的或更新的设置保留到文档中。

### <a name="getting-the-value-of-a-setting"></a>获取设置的值

下面的示例演示如何使用 [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) 方法获取名为"themeColor"的设置值。 该方法的唯一参数 `get` 是设置 _的名称（区分_ 大小写）。

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 方法 `get` 返回之前为传入的设置 _名称_ 保存的值。 如果不存在该设置，那么方法返回 **null**。

### <a name="removing-a-setting"></a>删除设置

下面的示例演示如何使用 [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) 方法删除名为"themeColor"的设置。 该方法的唯一参数 `remove` 是设置 _的名称（区分_ 大小写）。

```js
Office.context.document.settings.remove('themeColor');
```

如果不存在该设置，则不执行任何操作。 `Settings.saveAsync`使用 方法可保留文档中设置的删除操作。

### <a name="saving-your-settings"></a>保存设置

若要保存当前会话中加载项对设置属性包内存副本所做的任意添加、更改或删除操作，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) 方法将它们存储在文档中。 方法的唯一参数 `saveAsync` 是 _callback_，它是具有单个参数的回调函数。

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

作为 callback 参数传入`saveAsync`_方法的匿名_ 函数在操作完成时执行。 回调 _的 asyncResult_ 参数提供对包含 `AsyncResult` 操作状态的对象的访问。 在示例中，函数检查 `AsyncResult.status` 属性以查看保存操作是成功还是失败，然后在加载项页面中显示结果。

## <a name="how-to-save-custom-xml-to-the-document"></a>如何将自定义 XML 保存到文档

> [!NOTE]
> 此部分是在 Word 中支持的 Office 常见 JavaScript API 上下文中介绍的自定义 XML 部分。 特定于应用程序的应用程序Excel JavaScript API 还提供对自定义 XML 部件的访问权限。 Excel API 和编程模式有点不一样。 有关详细信息，请参阅 [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart)。

当需要存储的信息超过文档文档大小限制或具有结构化字符设置有一个额外的存储选项。 可以在 Word 的任务窗格加载项中暂留自定义 XML 标记（对于 Excel，但请参阅本节顶部的注释）。 在 Word 中，可以使用 [CustomXmlPart](/javascript/api/office/office.customxmlpart) 对象及其方法（同样，请参阅上面的 Excel 注释）。 以下代码将创建自定义 XML 部件，并在页面的 divs 中显示其 ID 及内容。 请注意，XML 字符串中必须有一个 `xmlns` 属性。

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

若要检索自定义 XML 部分，请使用 [getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) 方法，但 ID 是在创建 XML 部分时生成的 GUID，因此编码时无法知道 ID 是什么。 因此，最好是在创建 XML 部分时，立即将 XML 部分的 ID 存储为设置，并为它提供容易记住的密钥。 下面的方法展示了如何执行此操作。  (有关使用自定义设置的详细信息和最佳做法，请参阅本文的前面) 

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>如何在加载项中Outlook设置

若要了解如何在加载项中保存Outlook，请参阅管理加载项的状态[Outlook设置](../outlook/manage-state-and-settings-outlook.md)。

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Outlook 加载项](../outlook/outlook-add-ins-overview.md)
- [管理加载项的状态Outlook设置](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
