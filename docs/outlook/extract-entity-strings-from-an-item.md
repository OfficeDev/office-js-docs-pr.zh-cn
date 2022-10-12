---
title: 从 Outlook 项中提取实体字符串
description: 了解如何从 Outlook 加载项中的某个 Outlook 项中提取实体字符串。
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 512999272c720f5b87480c49d60c649bc6a886ad
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541273"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a>从 Outlook 项中提取实体字符串

This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.

> [!NOTE]
> 本文中介绍的 Outlook 外接程序功能使用激活规则，这些规则在使用 Office 外接程序的 Teams 清单的外接程序中不受支持 [ (预览版) ](../develop/json-manifest-overview.md)。

受支持的实体包括：

- **地址**：美国通信地址，至少包含街道号码、街道名称、城市、州和邮政编码等部分元素。

- **联系人**：个人联系信息，在地址或公司名称等其他实体的上下文中。

- **电子邮件地址**：SMTP 电子邮件地址。

- **Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.

- **电话号码**：北美电话号码。

- **任务建议**：通常以可操作短语表述的任务建议。

- **URL**

Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.

Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.

The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.

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

The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.

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

The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.

请注意，所有 Outlook 外接程序都必须包含 office.js。 下面的 HTML 文件包括内容分发网络 (CDN) 上office.js版本 1.1。

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

The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.

```CSS
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

[Office.initialize](/javascript/api/office#Office_initialize_reason_) 事件发生时，实体外接程序调用当前项目的 [getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法。 该 `getEntities` 方法返回全局变量 `_MyEntities` 支持的实体实例数组。 以下为相关的 JavaScript 代码。

```js
// Global variables
let _Item;
let _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    const _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}
```

## <a name="extracting-addresses"></a>提取地址

当用户单击“获取地址”按钮时，`myGetAddresses` 事件处理程序从 `_MyEntities` 对象的 [addressess](/javascript/api/outlook/office.entities#outlook-office-entities-addresses-member) 属性获取一组地址（如果已提取任何地址的话）。 提取的每个地址都存储为数组中的字符串。 `myGetAddresses` 在 `htmlText` 中形成本地 HTML 字符串以显示提取的地址的列表。 以下是相关的 JavaScript 代码。

```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    let htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    const addressesArray = _MyEntities.addresses;
    for (let i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-contact-information"></a>提取联系人信息

当用户单击 **“获取联系人信息”** 按钮时， `myGetContacts` 事件处理程序会从对象 [的联系](/javascript/api/outlook/office.entities#outlook-office-entities-contacts-member) 人属性 `_MyEntities` （如果提取了任何联系人）获取联系人数组及其信息。 每个提取的联系人在数组中被存储为 [Contact](/javascript/api/outlook/office.contact) 对象。 获取有关每个联系人的更多数据。 请注意，上下文确定 Outlook 是可以从邮件末尾的签名项目&mdash;中提取联系人，还是至少以下某些信息必须存在于联系人附近。

- 表示 [Contact.personName](/javascript/api/outlook/office.contact#outlook-office-contact-personname-member) 属性中联系人名称的字符串。

- 表示 [Contact.businessName](/javascript/api/outlook/office.contact#outlook-office-contact-businessname-member) 属性中与联系人关联的公司名称的字符串。

- The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#outlook-office-contact-phonenumbers-member) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.

- 对于电话号码数组中的每个 **PhoneNumber** 成员，表示 [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-phonestring-member) 属性中电话号码的字符串。

- The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#outlook-office-contact-urls-member) property. Each URL is represented as a string in an array member.

- The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#outlook-office-contact-emailaddresses-member) property. Each email address is represented as a string in an array member.

- The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#outlook-office-contact-addresses-member) property. Each postal address is represented as a string in an array member.

`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.

```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    let htmlText = "";

    // Gets an array of contacts and their information.
    const contactsArray = _MyEntities.contacts;
    for (let i = 0; i < contactsArray.length; i++)
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
        let phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (let j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        let urlsArray = contactsArray[i].urls;
        for (let j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        let emailAddressesArray = contactsArray[i].emailAddresses;
        for (let j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        let addressesArray = contactsArray[i].addresses;
        for (let j = 0; j < addressesArray.length; j++)
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

When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#outlook-office-entities-emailaddresses-member) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.

```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    let htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    const emailAddressesArray = _MyEntities.emailAddresses;
    for (let i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-meeting-suggestions"></a>提取会议建议

当用户单击“获取会议建议”按钮时，`myGetMeetingSuggestions` 事件处理程序从 `_MyEntities` 对象的 [meetingSuggestions](/javascript/api/outlook/office.entities#outlook-office-entities-meetingsuggestions-member) 属性获取一组会议建议（如果已提取任何会议建议的话）。

 > [!NOTE]
 > 只有消息而不支持约会支持 `MeetingSuggestion` 实体类型。

每个提取的会议建议都存储为数组中的一个 [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) 对象。 获取有关每个会议建议的更多数据：

- [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-meetingstring-member) 属性中已识别为会议建议的字符串。

- The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-attendees-member) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.

- 对于每个参与者，[EmailUser.displayName](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-displayname-member) 属性中的名称。

- 对于每个参与者，[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-emailaddress-member) 属性中的 SMTP 地址。

- [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-location-member) 属性中表示会议建议位置的字符串。

- [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-subject-member) 属性中表示会议建议主题的字符串。

- [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-start-member) 属性中表示会议建议开始时间的字符串。

- [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-end-member) 属性中表示会议建议结束时间的字符串。

`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.

```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    let htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    const meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (let i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        let attendeesArray = meetingsArray[i].attendees;
        for (let j = 0; j < attendeesArray.length; j++) {
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

When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#outlook-office-entities-phonenumbers-member) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:

- [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-type-member) 属性中表示电话号码种类的字符串（例如家庭电话号码）。

- [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-phonestring-member) 属性中表示实际电话号码的字符串。

- [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-originalphonestring-member) 属性中最初识别为电话号码的字符串。

`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.

```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    let htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    const phoneNumbersArray = _MyEntities.phoneNumbers;
    for (let i = 0; i < phoneNumbersArray.length; i++)
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

When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#outlook-office-entities-tasksuggestions-member) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:

- [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#outlook-office-tasksuggestion-taskstring-member) 属性中最初识别为任务建议的字符串。

- The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#outlook-office-tasksuggestion-assignees-member) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.

- 对于每个受托人，[EmailUser.displayName](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-displayname-member) 属性中的名称。

- 对于每个受托人，[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-emailaddress-member) 属性中的 SMTP 地址。

`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.

```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    let htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    const tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (let i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        let assigneesArray = tasksArray[i].assignees;
        for (let j = 0; j < assigneesArray.length; j++)
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

When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#outlook-office-entities-urls-member) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.

```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    let htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    const urlArray = _MyEntities.urls;
    for (let i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="clearing-displayed-entity-strings"></a>清除显示的实体字符串

Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.

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
let _Item;
let _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    const _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready method.
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
    let htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    const addressesArray = _MyEntities.addresses;
    for (let i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    let htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    const emailAddressesArray = _MyEntities.emailAddresses;
    for (let i = 0; i < emailAddressesArray.length; i++)
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
    let htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    const meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (let i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        let attendeesArray = meetingsArray[i].attendees;
        for (let j = 0; j < attendeesArray.length; j++)
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
    let htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    const phoneNumbersArray = _MyEntities.phoneNumbers;
    for (let i = 0; i < phoneNumbersArray.length; i++)
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
    let htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    const tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (let i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        let assigneesArray = tasksArray[i].assignees;
        for (let j = 0; j < assigneesArray.length; j++)
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
    let htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    const urlArray = _MyEntities.urls;
    for (let i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="see-also"></a>另请参阅

- [创建适用于阅读窗体的 Outlook 加载项](read-scenario.md)
- [将 Outlook 项目中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
- [item.getEntities 方法](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
