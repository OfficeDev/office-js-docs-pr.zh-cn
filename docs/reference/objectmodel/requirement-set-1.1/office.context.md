
# <a name="context"></a>context

### [Office](Office.md). context

Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。


##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

### <a name="namespaces"></a>命名空间

[mailbox](office.context.mailbox.md)：为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。

### <a name="members"></a>成员

####  <a name="displaylanguage-string"></a>displayLanguage :String

获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。

`displayLanguage` 值反映在 Office 主机应用程序中通过**文件 > 选项 > 语言**设置指定的当前**显示语言**。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)

获取一个对象，它表示保存到用户邮箱的邮件加载项的自定义设置或状态。

`RoamingSettings` 对象允许你存储和访问用户邮箱中存储的邮件加载项的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该加载项时，加载项可以使用数据。

##### <a name="type"></a>类型：

*   [RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|