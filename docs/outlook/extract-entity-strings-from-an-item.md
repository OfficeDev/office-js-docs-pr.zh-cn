---
title: 从 Outlook 项中提取实体字符串
description: 了解如何从 Outlook 加载项中的某个 Outlook 项中提取实体字符串。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d266795e3794cfa293d073dafc1ca714644faa5b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938867"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a>从 Outlook 项中提取实体字符串

本文介绍了如何创建“**显示实体**”Outlook 加载项，以从选定 Outlook 项的主题和正文中提取受支持的已知实体的字符串实例。此项可以是约会、电子邮件、会议请求、会议响应或会议取消。

受支持的实体包括：

- **地址**：美国通信地址，至少包含街道号码、街道名称、城市、州和邮政编码等部分元素。
    
- **联系人**：个人联系信息，在地址或公司名称等其他实体的上下文中。
    
- **电子邮件地址**：SMTP 电子邮件地址。
    
- **会议建议**：提及活动等会议建议。请注意，只有邮件（而不是约会）支持提取会议建议。
    
- **电话号码**：北美电话号码。
    
- **任务建议**：通常以可操作短语表述的任务建议。
    
- **URL**
    
大多数这些实体依赖于基于大量数据机器学习的自然语言识别。因此，识别是非确定性的，有时候依赖于 Outlook 项目中的上下文。

无论用户选择查看约会、电子邮件或会议要求、响应或取消，Outlook 均会激活实体外接程序。在初始化期间，示例实体外接程序从当前项读取受支持的实体的所有实例。 

外接程序为用户提供按钮以选择实体类型。当用户选择一个实体时，外接程序在外接程序窗格中显示所选实体的实例。以下各节列出了实体外接程序的 XML 清单及 HTML 和 JavaScript 文件，并突出显示支持各自实体提取的代码。

## <a name="xml-manifest"></a>XML 清单

实体外接程序具有两个由逻辑 OR 运算连接的激活规则。 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

这些规则指定 Outlook 应在阅读窗格或阅读检查器中的当前所选项目为约会或邮件（包括电子邮件、会议请求、响应或取消）时激活此加载项。

下面是实体外接程序的清单。它对 Office 外接程序清单使用架构的 1.1 版本。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## <a name="html-implementation"></a>HTML 实现

实体外接程序的 HTML 文件为用户指定按钮以选择每种类型的实体，另外还指定另一个按钮以清除显示的实体实例。它包括 JavaScript 文件 default_entities.js，这在下一节的 [JavaScript 实现](#javascript-implementation)中进行介绍。JavaScript 文件包括其中每个按钮的事件处理程序。

请注意，所有 Outlook 外接程序都必须包含 office.js。下面的 HTML 文件包含 CDN 上 office.js 的版本 1.1。 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## <a name="style-sheet"></a>样式表


实体外接程序使用可选 CSS 文件 default_entities.css 指定输出的布局。下面为 CSS 文件的列表。


```CSS
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## <a name="javascript-implementation"></a>JavaScript 实现

其余部分介绍此示例（default_entities.js 文件）如何从用户查看的邮件或约会的主题和正文中提取已知实体。

## <a name="extracting-entities-upon-initialization"></a>初始化时提取实体

[Office.initialize](/javascript/api/office#Office_initialize_reason_) 事件发生时，实体外接程序调用当前项目的 [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法。 `getEntities`方法返回全局变量 `_MyEntities` 受支持实体的实例数组。 以下为相关的 JavaScript 代码。


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## <a name="extracting-addresses"></a>提取地址


当用户单击“获取地址”按钮时，`myGetAddresses` 事件处理程序从 `_MyEntities` 对象的 [addressess](/javascript/api/outlook/office.entities#addresses) 属性获取一组地址（如果已提取任何地址的话）。 提取的每个地址都存储为数组中的字符串。 `myGetAddresses` 在 `htmlText` 中形成本地 HTML 字符串以显示提取的地址的列表。 以下是相关的 JavaScript 代码。


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-contact-information"></a>提取联系人信息


当用户单击"**获取** 联系人信息"按钮时，事件处理程序从对象的 contacts 属性获取一组联系人及其信息（ `myGetContacts` 如果已提取任何 [](/javascript/api/outlook/office.entities#contacts) `_MyEntities` 联系人）。 每个提取的联系人在数组中被存储为 [Contact](/javascript/api/outlook/office.contact) 对象。 获取有关每个联系人的更多数据。 请注意，上下文确定 Outlook 是否可以从项目中提取电子邮件末尾的签名的联系人，或者联系人附近至少必须存在以下部分 &mdash; 信息。


- 表示 [Contact.personName](/javascript/api/outlook/office.contact#personName) 属性中联系人名称的字符串。

- 表示 [Contact.businessName](/javascript/api/outlook/office.contact#businessName) 属性中与联系人关联的公司名称的字符串。

- [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phoneNumbers) 属性中与联系人关联的电话号码数组。每个电话号码都由一个 [PhoneNumber](/javascript/api/outlook/office.phonenumber) 对象表示。

- 对于电话号码数组中的每个 **PhoneNumber** 成员，表示 [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phoneString) 属性中电话号码的字符串。

- [Contact.urls](/javascript/api/outlook/office.contact#urls) 属性中与联系人关联的 URL 的数组。每个 URL 都表示为数组成员中的一个字符串。

- [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailAddresses) 属性中与联系人关联的电子邮件地址的数组。每个电子邮件地址都表示为数组成员中的一个字符串。

- [Contact.addresses](/javascript/api/outlook/office.contact#addresses) 属性中与联系人关联的通信地址的数组。每个通信地址都表示为数组成员中的一个字符串。

`myGetContacts` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示每个联系人的数据。以下为相关的 JavaScript 代码。




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-email-addresses"></a>提取电子邮件地址


当用户单击“获取电子邮件地址”按钮时，`myGetEmailAddresses` 事件处理程序从 `_MyEntities` 对象的 [emailAddresses](/javascript/api/outlook/office.entities#emailAddresses) 属性获取一组 SMTP 电子邮件地址（如果已提取任何电子邮件地址的话）。提取的每个电子邮件地址都存储为数组中的字符串。`myGetEmailAddresses` 在 `htmlText` 中构成本地 HTML 字符串，以列出提取的电子邮件地址。下面展示了相关 JavaScript 代码。


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-meeting-suggestions"></a>提取会议建议


当用户单击“获取会议建议”按钮时，`myGetMeetingSuggestions` 事件处理程序从 `_MyEntities` 对象的 [meetingSuggestions](/javascript/api/outlook/office.entities#meetingSuggestions) 属性获取一组会议建议（如果已提取任何会议建议的话）。


 > [!NOTE]
 > 只有邮件（而非约会）才支持 `MeetingSuggestion` 实体类型。

每个提取的会议建议都存储为数组中的一个 [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) 对象。`myGetMeetingSuggestions` 获取有关每个会议建议的更多数据：


- [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingString) 属性中已识别为会议建议的字符串。

- [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) 属性中会议参与者的数组。每个参与者都由一个 [EmailUser](/javascript/api/outlook/office.emailuser) 对象表示。

- 对于每个参与者，[EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayName) 属性中的名称。

- 对于每个参与者，[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailAddress) 属性中的 SMTP 地址。

- [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) 属性中表示会议建议位置的字符串。

- [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) 属性中表示会议建议主题的字符串。

- [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) 属性中表示会议建议开始时间的字符串。

- [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) 属性中表示会议建议结束时间的字符串。

`myGetMeetingSuggestions` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示其中每个会议建议的数据。以下是相关的 JavaScript 代码。




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-phone-numbers"></a>提取电话号码


当用户单击“获取电话号码”按钮时，`myGetPhoneNumbers` 事件处理程序从 `_MyEntities` 对象的 [phoneNumbers](/javascript/api/outlook/office.entities#phoneNumbers) 属性获取一组电话号码（如果已提取任何电话号码的话）。提取的每个电话号码都存储为数组中的 [PhoneNumber](/javascript/api/outlook/office.phonenumber) 对象。`myGetPhoneNumbers` 获取每个电话号码的更多数据：


- [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) 属性中表示电话号码种类的字符串（例如家庭电话号码）。

- [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phoneString) 属性中表示实际电话号码的字符串。

- [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalPhoneString) 属性中最初识别为电话号码的字符串。

`myGetPhoneNumbers` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示每个电话号码的数据。以下是相关的 JavaScript 代码。




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-task-suggestions"></a>提取任务建议


当用户单击“获取任务建议”按钮时，`myGetTaskSuggestions` 事件处理程序从 `_MyEntities` 对象的 [taskSuggestions](/javascript/api/outlook/office.entities#taskSuggestions) 属性获取一组任务建议（如果已提取任何任务建议的话）。提取每个的任务建议都存储为数组中的 [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) 对象。`myGetTaskSuggestions` 获取每个任务建议的更多数据：


- [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskString) 属性中最初识别为任务建议的字符串。

- [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) 属性中任务受托人的数组。每个受托人都由一个 [EmailUser](/javascript/api/outlook/office.emailuser) 对象表示。

- 对于每个受托人，[EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayName) 属性中的名称。

- 对于每个受托人，[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailAddress) 属性中的 SMTP 地址。

`myGetTaskSuggestions` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示每个任务建议的数据。以下为相关的 JavaScript 代码。




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="extracting-urls"></a>提取 URL


当用户单击“获取 URL”按钮时，`myGetUrls` 事件处理程序从 `_MyEntities` 对象的 [urls](/javascript/api/outlook/office.entities#urls) 属性获取一组 URL（如果已提取任何 URL 的话）。提取每个的 URL 都存储为数组中的字符串。`myGetUrls` 在 `htmlText` 中构成本地 HTML 字符串，以列出提取的 URL。


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="clearing-displayed-entity-strings"></a>清除显示的实体字符串


最后，实体外接程序指定一个 `myClearEntitiesBox` 事件处理程序，以清除任何显示的字符串。以下是相关的代码。


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a>JavaScript 列表


以下是 JavaScript 实现的完整列表。


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="see-also"></a>另请参阅

- [创建适用于阅读窗体的 Outlook 加载项](read-scenario.md)
- [将 Outlook 项目中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
- [item.getEntities 方法](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
