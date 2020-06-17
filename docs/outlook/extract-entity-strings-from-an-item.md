---
title: 从 Outlook 项目中提取实体字符串
description: 了解如何从 Outlook 加载项中的某个 Outlook 项中提取实体字符串。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: b15ad23427f79a333ae8ae9d342acdf28e6d010c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608941"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a><span data-ttu-id="1daaf-103">从 Outlook 项中提取实体字符串</span><span class="sxs-lookup"><span data-stu-id="1daaf-103">Extract entity strings from an Outlook item</span></span>

<span data-ttu-id="1daaf-p101">本文介绍了如何创建“**显示实体**”Outlook 加载项，以从选定 Outlook 项的主题和正文中提取受支持的已知实体的字符串实例。此项可以是约会、电子邮件、会议请求、会议响应或会议取消。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p101">This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.</span></span>

<span data-ttu-id="1daaf-106">受支持的实体包括：</span><span class="sxs-lookup"><span data-stu-id="1daaf-106">The supported entities include:</span></span>

- <span data-ttu-id="1daaf-107">**地址**：美国通信地址，至少包含街道号码、街道名称、城市、州和邮政编码等部分元素。</span><span class="sxs-lookup"><span data-stu-id="1daaf-107">**Address**: A United States postal address, that has at least a subset of the elements of a street number, street name, city, state, and zip code.</span></span>
    
- <span data-ttu-id="1daaf-108">**联系人**：个人联系信息，在地址或公司名称等其他实体的上下文中。</span><span class="sxs-lookup"><span data-stu-id="1daaf-108">**Contact**: A person's contact information, in the context of other entities such as an address or business name.</span></span>
    
- <span data-ttu-id="1daaf-109">**电子邮件地址**：SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="1daaf-109">**Email address**: An SMTP email address.</span></span>
    
- <span data-ttu-id="1daaf-p102">**会议建议**：提及活动等会议建议。请注意，只有邮件（而不是约会）支持提取会议建议。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p102">**Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.</span></span>
    
- <span data-ttu-id="1daaf-112">**电话号码**：北美电话号码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-112">**Phone number**: A North American phone number.</span></span>
    
- <span data-ttu-id="1daaf-113">**任务建议**：通常以可操作短语表述的任务建议。</span><span class="sxs-lookup"><span data-stu-id="1daaf-113">**Task suggestion**: A task suggestion, typically expressed in an actionable phrase.</span></span>
    
- <span data-ttu-id="1daaf-114">**URL**</span><span class="sxs-lookup"><span data-stu-id="1daaf-114">**URL**</span></span>
    
<span data-ttu-id="1daaf-p103">大多数这些实体依赖于基于大量数据机器学习的自然语言识别。因此，识别是非确定性的，有时候依赖于 Outlook 项目中的上下文。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p103">Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.</span></span>

<span data-ttu-id="1daaf-p104">无论用户选择查看约会、电子邮件或会议要求、响应或取消，Outlook 均会激活实体外接程序。在初始化期间，示例实体外接程序从当前项读取受支持的实体的所有实例。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p104">Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.</span></span> 

<span data-ttu-id="1daaf-p105">外接程序为用户提供按钮以选择实体类型。当用户选择一个实体时，外接程序在外接程序窗格中显示所选实体的实例。以下各节列出了实体外接程序的 XML 清单及 HTML 和 JavaScript 文件，并突出显示支持各自实体提取的代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p105">The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.</span></span>

## <a name="xml-manifest"></a><span data-ttu-id="1daaf-122">XML 清单</span><span class="sxs-lookup"><span data-stu-id="1daaf-122">XML manifest</span></span>

<span data-ttu-id="1daaf-123">实体外接程序具有两个由逻辑 OR 运算连接的激活规则。</span><span class="sxs-lookup"><span data-stu-id="1daaf-123">The entities add-in has two activation rules joined by a logical OR operation.</span></span> 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

<span data-ttu-id="1daaf-124">这些规则指定 Outlook 应在阅读窗格或阅读检查器中的当前所选项目为约会或邮件（包括电子邮件、会议请求、响应或取消）时激活此加载项。</span><span class="sxs-lookup"><span data-stu-id="1daaf-124">These rules specify that Outlook should activate this add-in when the currently selected item in the Reading Pane or read inspector is an appointment or message (including an email message, or meeting request, response, or cancellation).</span></span>

<span data-ttu-id="1daaf-p106">下面是实体外接程序的清单。它对 Office 外接程序清单使用架构的 1.1 版本。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p106">The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.</span></span>

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


## <a name="html-implementation"></a><span data-ttu-id="1daaf-127">HTML 实现</span><span class="sxs-lookup"><span data-stu-id="1daaf-127">HTML implementation</span></span>

<span data-ttu-id="1daaf-p107">实体外接程序的 HTML 文件为用户指定按钮以选择每种类型的实体，另外还指定另一个按钮以清除显示的实体实例。它包括 JavaScript 文件 default_entities.js，这在下一节的 [JavaScript 实现](#javascript-implementation)中进行介绍。JavaScript 文件包括其中每个按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p107">The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.</span></span>

<span data-ttu-id="1daaf-p108">请注意，所有 Outlook 外接程序都必须包含 office.js。下面的 HTML 文件包含 CDN 上 office.js 的版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p108">Note that all Outlook add-ins must include office.js. The HTML file that follows includes version 1.1 of office.js on the CDN.</span></span> 

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


## <a name="style-sheet"></a><span data-ttu-id="1daaf-133">样式表</span><span class="sxs-lookup"><span data-stu-id="1daaf-133">Style sheet</span></span>


<span data-ttu-id="1daaf-p109">实体外接程序使用可选 CSS 文件 default_entities.css 指定输出的布局。下面为 CSS 文件的列表。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p109">The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.</span></span>


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


## <a name="javascript-implementation"></a><span data-ttu-id="1daaf-136">JavaScript 实现</span><span class="sxs-lookup"><span data-stu-id="1daaf-136">JavaScript implementation</span></span>

<span data-ttu-id="1daaf-137">其余部分介绍此示例（default_entities.js 文件）如何从用户查看的邮件或约会的主题和正文中提取已知实体。</span><span class="sxs-lookup"><span data-stu-id="1daaf-137">The remaining sections describe how this sample (default_entities.js file) extracts well-known entities from the subject and body of the message or appointment that the user is viewing.</span></span>

## <a name="extracting-entities-upon-initialization"></a><span data-ttu-id="1daaf-138">初始化时提取实体</span><span class="sxs-lookup"><span data-stu-id="1daaf-138">Extracting entities upon initialization</span></span>

<span data-ttu-id="1daaf-139">[Office.initialize](/javascript/api/office#office-initialize-reason-) 事件发生时，实体外接程序调用当前项目的 [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法。</span><span class="sxs-lookup"><span data-stu-id="1daaf-139">Upon the [Office.initialize](/javascript/api/office#office-initialize-reason-) event, the entities add-in calls the [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method of the current item.</span></span> <span data-ttu-id="1daaf-140">该 `getEntities` 方法返回全局变量 `_MyEntities` 受支持实体的实例数组。</span><span class="sxs-lookup"><span data-stu-id="1daaf-140">The `getEntities` method returns the global variable `_MyEntities` an array of instances of supported entities.</span></span> <span data-ttu-id="1daaf-141">以下为相关的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-141">The following is the related JavaScript code.</span></span>


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


## <a name="extracting-addresses"></a><span data-ttu-id="1daaf-142">提取地址</span><span class="sxs-lookup"><span data-stu-id="1daaf-142">Extracting addresses</span></span>


<span data-ttu-id="1daaf-143">当用户单击“获取地址”\*\*\*\* 按钮时，`myGetAddresses` 事件处理程序从 `_MyEntities` 对象的 [addressess](/javascript/api/outlook/office.entities#addresses) 属性获取一组地址（如果已提取任何地址的话）。</span><span class="sxs-lookup"><span data-stu-id="1daaf-143">When the user clicks the **Get Addresses** button, the `myGetAddresses` event handler obtains an array of addresses from the [addresses](/javascript/api/outlook/office.entities#addresses) property of the `_MyEntities` object, if any address was extracted.</span></span> <span data-ttu-id="1daaf-144">提取的每个地址都存储为数组中的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-144">Each extracted address is stored as a string in the array.</span></span> <span data-ttu-id="1daaf-145">`myGetAddresses` 在 `htmlText` 中形成本地 HTML 字符串以显示提取的地址的列表。</span><span class="sxs-lookup"><span data-stu-id="1daaf-145">`myGetAddresses` forms a local HTML string in `htmlText` to display the list of extracted addresses.</span></span> <span data-ttu-id="1daaf-146">以下是相关的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-146">The following is the related JavaScript code.</span></span>


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


## <a name="extracting-contact-information"></a><span data-ttu-id="1daaf-147">提取联系人信息</span><span class="sxs-lookup"><span data-stu-id="1daaf-147">Extracting contact information</span></span>


<span data-ttu-id="1daaf-p112">当用户单击“获取联系人信息”\*\*\*\* 按钮时，`myGetContacts` 事件处理程序从 `_MyEntities` 对象的 [contacts](/javascript/api/outlook/office.entities#contacts) 属性获取一组联系人及其信息（如果已提取任何联系人的话）。提取的每个联系人都存储为数组中的 [Contact](/javascript/api/outlook/office.contact) 对象。`myGetContacts` 获取每个联系人的更多数据。请注意，上下文确定 Outlook 能否从项提取联系人，即通过电子邮件末尾的签名，或联系人附近至少必须有以下部分信息：</span><span class="sxs-lookup"><span data-stu-id="1daaf-p112">When the user clicks the **Get Contact Information** button, the `myGetContacts` event handler obtains an array of contacts together with their information from the [contacts](/javascript/api/outlook/office.entities#contacts) property of the `_MyEntities` object, if any was extracted. Each extracted contact is stored as a [Contact](/javascript/api/outlook/office.contact) object in the array. `myGetContacts` obtains further data about each contact. Note that the context determines whether Outlook can extract a contact from an item&mdash;a signature at the end of an email message, or at least some of the following information would have to exist in the vicinity of the contact:</span></span>


- <span data-ttu-id="1daaf-152">表示 [Contact.personName](/javascript/api/outlook/office.contact#personname) 属性中联系人名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-152">The string representing the contact's name from the [Contact.personName](/javascript/api/outlook/office.contact#personname) property.</span></span>

- <span data-ttu-id="1daaf-153">表示 [Contact.businessName](/javascript/api/outlook/office.contact#businessname) 属性中与联系人关联的公司名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-153">The string representing the company name associated with the contact from the [Contact.businessName](/javascript/api/outlook/office.contact#businessname) property.</span></span>

- <span data-ttu-id="1daaf-p113">[Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) 属性中与联系人关联的电话号码数组。每个电话号码都由一个 [PhoneNumber](/javascript/api/outlook/office.phonenumber) 对象表示。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p113">The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.</span></span>

- <span data-ttu-id="1daaf-156">对于电话号码数组中的每个 **PhoneNumber** 成员，表示 [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) 属性中电话号码的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-156">For each **PhoneNumber** member in the telephone numbers array, the string representing the telephone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="1daaf-p114">[Contact.urls](/javascript/api/outlook/office.contact#urls) 属性中与联系人关联的 URL 的数组。每个 URL 都表示为数组成员中的一个字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p114">The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#urls) property. Each URL is represented as a string in an array member.</span></span>

- <span data-ttu-id="1daaf-p115">[Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) 属性中与联系人关联的电子邮件地址的数组。每个电子邮件地址都表示为数组成员中的一个字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p115">The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) property. Each email address is represented as a string in an array member.</span></span>

- <span data-ttu-id="1daaf-p116">[Contact.addresses](/javascript/api/outlook/office.contact#addresses) 属性中与联系人关联的通信地址的数组。每个通信地址都表示为数组成员中的一个字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p116">The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#addresses) property. Each postal address is represented as a string in an array member.</span></span>

<span data-ttu-id="1daaf-p117">`myGetContacts` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示每个联系人的数据。以下为相关的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p117">`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-email-addresses"></a><span data-ttu-id="1daaf-165">提取电子邮件地址</span><span class="sxs-lookup"><span data-stu-id="1daaf-165">Extracting email addresses</span></span>


<span data-ttu-id="1daaf-p118">当用户单击“获取电子邮件地址”\*\*\*\* 按钮时，`myGetEmailAddresses` 事件处理程序从 `_MyEntities` 对象的 [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) 属性获取一组 SMTP 电子邮件地址（如果已提取任何电子邮件地址的话）。提取的每个电子邮件地址都存储为数组中的字符串。`myGetEmailAddresses` 在 `htmlText` 中构成本地 HTML 字符串，以列出提取的电子邮件地址。下面展示了相关 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p118">When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.</span></span>


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


## <a name="extracting-meeting-suggestions"></a><span data-ttu-id="1daaf-170">提取会议建议</span><span class="sxs-lookup"><span data-stu-id="1daaf-170">Extracting meeting suggestions</span></span>


<span data-ttu-id="1daaf-171">当用户单击“获取会议建议”\*\*\*\* 按钮时，`myGetMeetingSuggestions` 事件处理程序从 `_MyEntities` 对象的 [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) 属性获取一组会议建议（如果已提取任何会议建议的话）。</span><span class="sxs-lookup"><span data-stu-id="1daaf-171">When the user clicks the **Get Meeting Suggestions** button, the `myGetMeetingSuggestions` event handler obtains an array of meeting suggestions from the [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) property of the `_MyEntities` object, if any was extracted.</span></span>


 > [!NOTE]
 > <span data-ttu-id="1daaf-172">仅邮件而非约会支持 `MeetingSuggestion` 实体类型。</span><span class="sxs-lookup"><span data-stu-id="1daaf-172">Only messages but not appointments support the `MeetingSuggestion` entity type.</span></span>

<span data-ttu-id="1daaf-p119">每个提取的会议建议都存储为数组中的一个 [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) 对象。`myGetMeetingSuggestions` 获取有关每个会议建议的更多数据：</span><span class="sxs-lookup"><span data-stu-id="1daaf-p119">Each extracted meeting suggestion is stored as a [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object in the array. `myGetMeetingSuggestions` obtains further data about each meeting suggestion:</span></span>


- <span data-ttu-id="1daaf-175">[MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) 属性中已识别为会议建议的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-175">The string that was identified as a meeting suggestion from the [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) property.</span></span>

- <span data-ttu-id="1daaf-p120">[MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) 属性中会议参与者的数组。每个参与者都由一个 [EmailUser](/javascript/api/outlook/office.emailuser) 对象表示。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p120">The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="1daaf-178">对于每个参与者，[EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) 属性中的名称。</span><span class="sxs-lookup"><span data-stu-id="1daaf-178">For each attendee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="1daaf-179">对于每个参与者，[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) 属性中的 SMTP 地址。</span><span class="sxs-lookup"><span data-stu-id="1daaf-179">For each attendee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

- <span data-ttu-id="1daaf-180">[MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) 属性中表示会议建议位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-180">The string representing the location of the meeting suggestion from the [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) property.</span></span>

- <span data-ttu-id="1daaf-181">[MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) 属性中表示会议建议主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-181">The string representing the subject of the meeting suggestion from the [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) property.</span></span>

- <span data-ttu-id="1daaf-182">[MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) 属性中表示会议建议开始时间的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-182">The string representing the start time of the meeting suggestion from the [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) property.</span></span>

- <span data-ttu-id="1daaf-183">[MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) 属性中表示会议建议结束时间的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-183">The string representing the end time of the meeting suggestion from the [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) property.</span></span>

<span data-ttu-id="1daaf-p121">`myGetMeetingSuggestions` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示其中每个会议建议的数据。以下是相关的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p121">`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-phone-numbers"></a><span data-ttu-id="1daaf-186">提取电话号码</span><span class="sxs-lookup"><span data-stu-id="1daaf-186">Extracting phone numbers</span></span>


<span data-ttu-id="1daaf-p122">当用户单击“获取电话号码”\*\*\*\* 按钮时，`myGetPhoneNumbers` 事件处理程序从 `_MyEntities` 对象的 [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) 属性获取一组电话号码（如果已提取任何电话号码的话）。提取的每个电话号码都存储为数组中的 [PhoneNumber](/javascript/api/outlook/office.phonenumber) 对象。`myGetPhoneNumbers` 获取每个电话号码的更多数据：</span><span class="sxs-lookup"><span data-stu-id="1daaf-p122">When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:</span></span>


- <span data-ttu-id="1daaf-190">[PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) 属性中表示电话号码种类的字符串（例如家庭电话号码）。</span><span class="sxs-lookup"><span data-stu-id="1daaf-190">The string representing the kind of phone number, for example, home phone number, from the [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) property.</span></span>

- <span data-ttu-id="1daaf-191">[PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) 属性中表示实际电话号码的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-191">The string representing the actual phone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="1daaf-192">[PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) 属性中最初识别为电话号码的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-192">The string that was originally identified as the phone number from the [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) property.</span></span>

<span data-ttu-id="1daaf-p123">`myGetPhoneNumbers` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示每个电话号码的数据。以下是相关的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p123">`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-task-suggestions"></a><span data-ttu-id="1daaf-195">提取任务建议</span><span class="sxs-lookup"><span data-stu-id="1daaf-195">Extracting task suggestions</span></span>


<span data-ttu-id="1daaf-p124">当用户单击“获取任务建议”\*\*\*\* 按钮时，`myGetTaskSuggestions` 事件处理程序从 `_MyEntities` 对象的 [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) 属性获取一组任务建议（如果已提取任何任务建议的话）。提取每个的任务建议都存储为数组中的 [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) 对象。`myGetTaskSuggestions` 获取每个任务建议的更多数据：</span><span class="sxs-lookup"><span data-stu-id="1daaf-p124">When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:</span></span>


- <span data-ttu-id="1daaf-199">[TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) 属性中最初识别为任务建议的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daaf-199">The string that was originally identified a task suggestion from the [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) property.</span></span>

- <span data-ttu-id="1daaf-p125">[TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) 属性中任务受托人的数组。每个受托人都由一个 [EmailUser](/javascript/api/outlook/office.emailuser) 对象表示。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p125">The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="1daaf-202">对于每个受托人，[EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) 属性中的名称。</span><span class="sxs-lookup"><span data-stu-id="1daaf-202">For each assignee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="1daaf-203">对于每个受托人，[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) 属性中的 SMTP 地址。</span><span class="sxs-lookup"><span data-stu-id="1daaf-203">For each assignee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

<span data-ttu-id="1daaf-p126">`myGetTaskSuggestions` 在 `htmlText` 中形成一个本地 HTML 字符串，以显示每个任务建议的数据。以下为相关的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p126">`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.</span></span>




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


## <a name="extracting-urls"></a><span data-ttu-id="1daaf-206">提取 URL</span><span class="sxs-lookup"><span data-stu-id="1daaf-206">Extracting URLs</span></span>


<span data-ttu-id="1daaf-p127">当用户单击“获取 URL”\*\*\*\* 按钮时，`myGetUrls` 事件处理程序从 `_MyEntities` 对象的 [urls](/javascript/api/outlook/office.entities#urls) 属性获取一组 URL（如果已提取任何 URL 的话）。提取每个的 URL 都存储为数组中的字符串。`myGetUrls` 在 `htmlText` 中构成本地 HTML 字符串，以列出提取的 URL。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p127">When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#urls) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.</span></span>


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


## <a name="clearing-displayed-entity-strings"></a><span data-ttu-id="1daaf-210">清除显示的实体字符串</span><span class="sxs-lookup"><span data-stu-id="1daaf-210">Clearing displayed entity strings</span></span>


<span data-ttu-id="1daaf-p128">最后，实体外接程序指定一个 `myClearEntitiesBox` 事件处理程序，以清除任何显示的字符串。以下是相关的代码。</span><span class="sxs-lookup"><span data-stu-id="1daaf-p128">Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.</span></span>


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a><span data-ttu-id="1daaf-213">JavaScript 列表</span><span class="sxs-lookup"><span data-stu-id="1daaf-213">JavaScript listing</span></span>


<span data-ttu-id="1daaf-214">以下是 JavaScript 实现的完整列表。</span><span class="sxs-lookup"><span data-stu-id="1daaf-214">The following is the complete listing of the JavaScript implementation.</span></span>


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


## <a name="see-also"></a><span data-ttu-id="1daaf-215">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1daaf-215">See also</span></span>

- [<span data-ttu-id="1daaf-216">创建适用于阅读窗体的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="1daaf-216">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="1daaf-217">将 Outlook 项目中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="1daaf-217">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="1daaf-218">item.getEntities 方法</span><span class="sxs-lookup"><span data-stu-id="1daaf-218">item.getEntities method</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
