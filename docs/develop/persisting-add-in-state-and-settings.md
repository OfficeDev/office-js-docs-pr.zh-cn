---
title: 暂留加载项状态和设置
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: ee65d6b1f033b012a548bc685b9228679bec8c5e
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925568"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="40db9-102">暂留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="40db9-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="40db9-p101">Office 加载项实质上是在浏览器控件的无状态环境中运行的 Web 应用。因此，加载项可能需要暂留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能有需要在下一次初始化时保存和重新加载的自定义设置或其他值（如用户的首选视图或默认位置）。为此，可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="40db9-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="40db9-107">使用适用于 Office 的 JavaScript API 成员，将数据存储为：</span><span class="sxs-lookup"><span data-stu-id="40db9-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="40db9-108">在依赖加载项类型的位置上存储的属性包中的名称-数值对。</span><span class="sxs-lookup"><span data-stu-id="40db9-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="40db9-109">在文档中存储的自定义 XML。</span><span class="sxs-lookup"><span data-stu-id="40db9-109">Custom XML stored in the document.</span></span>
    
- <span data-ttu-id="40db9-110">使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。</span><span class="sxs-lookup"><span data-stu-id="40db9-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>
    
<span data-ttu-id="40db9-p102">本文重点介绍如何使用适用于 Office 的 JavaScript API 保留外接程序状态。有关使用浏览器 Cookie 和 Web 存储的示例，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。</span><span class="sxs-lookup"><span data-stu-id="40db9-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="40db9-113">使用适用于 Office 的 JavaScript API 保留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="40db9-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="40db9-p103">适用于 Office 的 JavaScript API 为在各个会话中保存外接程序状态提供了 [Settings](https://dev.office.com/reference/add-ins/shared/settings)、 [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 和 [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) 对象，如下表中所述。在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](https://dev.office.com/reference/add-ins/manifest/id) 相关联。</span><span class="sxs-lookup"><span data-stu-id="40db9-p103">The JavaScript API for Office provides the [Settings](https://dev.office.com/reference/add-ins/shared/settings), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings), and [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](https://dev.office.com/reference/add-ins/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="40db9-116">**对象**</span><span class="sxs-lookup"><span data-stu-id="40db9-116">**Object**</span></span>|<span data-ttu-id="40db9-117">**外接程序类型支持**</span><span class="sxs-lookup"><span data-stu-id="40db9-117">**Add-in type support**</span></span>|<span data-ttu-id="40db9-118">**存储位置**</span><span class="sxs-lookup"><span data-stu-id="40db9-118">**Storage location**</span></span>|<span data-ttu-id="40db9-119">**Office 主机支持**</span><span class="sxs-lookup"><span data-stu-id="40db9-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="40db9-120">设置</span><span class="sxs-lookup"><span data-stu-id="40db9-120">Settings</span></span>](https://dev.office.com/reference/add-ins/shared/settings)|<span data-ttu-id="40db9-121">内容和任务窗格</span><span class="sxs-lookup"><span data-stu-id="40db9-121">content and task pane</span></span>|<span data-ttu-id="40db9-122">加载项要使用的文档、电子表格或演示文稿。内容和任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。</span><span class="sxs-lookup"><span data-stu-id="40db9-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="40db9-p104">**重要说明：** 不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。您应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。</span><span class="sxs-lookup"><span data-stu-id="40db9-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="40db9-126">Word、Excel 或 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="40db9-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="40db9-p105">**注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。不过，对于在 Project（及其他 Office 主机应用）中运行的加载项，可以使用浏览器 Cookie 或 Web 存储等技术。若要详细了解这些技术，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。</span><span class="sxs-lookup"><span data-stu-id="40db9-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="40db9-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="40db9-130">RoamingSettings</span></span>](https://dev.office.com/reference/add-ins/outlook/RoamingSettings)|<span data-ttu-id="40db9-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="40db9-131">Outlook</span></span>|<span data-ttu-id="40db9-132">安装了加载项的用户 Exchange 服务器邮箱。由于这些设置存储在用户的服务器邮箱中，因此如果加载项在任何访问用户邮箱的受支持客户端主机应用或浏览器的上下文中运行，这些设置可随用户“漫游”，且可供加载项使用。</span><span class="sxs-lookup"><span data-stu-id="40db9-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="40db9-133">Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。</span><span class="sxs-lookup"><span data-stu-id="40db9-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="40db9-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="40db9-134">Outlook</span></span>|
|[<span data-ttu-id="40db9-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="40db9-135">CustomProperties</span></span>](https://dev.office.com/reference/add-ins/outlook/CustomProperties)|<span data-ttu-id="40db9-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="40db9-136">Outlook</span></span>|<span data-ttu-id="40db9-p106">加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。</span><span class="sxs-lookup"><span data-stu-id="40db9-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="40db9-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="40db9-139">Outlook</span></span>|
|[<span data-ttu-id="40db9-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="40db9-140">CustomXMLParts</span></span>](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts)|<span data-ttu-id="40db9-141">任务窗格</span><span class="sxs-lookup"><span data-stu-id="40db9-141">task pane</span></span>|<span data-ttu-id="40db9-p107">加载项要使用的文档、电子表格或演示文稿。任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。</span><span class="sxs-lookup"><span data-stu-id="40db9-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="40db9-p108">**重要说明：** 请勿将密码和其他敏感的个人身份信息 (PII) 存储在自定义 XML 部分中。虽然保存的数据对最终用户不可见，但它存储为文档的一部分，可通过直接读取文档的文件格式进行访问。应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在服务器上，且服务器将加载项托管为用户保护资源。</span><span class="sxs-lookup"><span data-stu-id="40db9-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="40db9-147">Word（使用 Office JavaScript 常见 API）、Excel（使用主机专用 Excel JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="40db9-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="40db9-148">设置数据在运行时托管在内存中</span><span class="sxs-lookup"><span data-stu-id="40db9-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="40db9-149">下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。</span><span class="sxs-lookup"><span data-stu-id="40db9-149">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="40db9-150">主机专用 Excel JavaScript API 还提供对自定义设置的访问权限。</span><span class="sxs-lookup"><span data-stu-id="40db9-150">The host-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="40db9-151">Excel API 和编程模式有点不一样。</span><span class="sxs-lookup"><span data-stu-id="40db9-151">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="40db9-152">有关详细信息，请参阅 [Excel SettingCollection](https://dev.office.com/reference/add-ins/excel/settingcollection)。</span><span class="sxs-lookup"><span data-stu-id="40db9-152">For more information, see [Excel SettingCollection](https://dev.office.com/reference/add-ins/excel/settingcollection).</span></span>

<span data-ttu-id="40db9-p110">在内部，通过  **Settings**、 **CustomProperties** 或 **RoamingSettings** 对象访问的属性包中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 对象，包含名称/值对。每个值的名称（键）必须为 **string**，且存储的值可为 JavaScript  **string**、 **number**、 **date** 或 **object**，但不能为  **function**。</span><span class="sxs-lookup"><span data-stu-id="40db9-p110">Internally, the data in the property bag accessed with the  **Settings**,  **CustomProperties**, or  **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a **string**, and the stored value can be a JavaScript  **string**,  **number**,  **date**, or  **object**, but not a  **function**.</span></span>

<span data-ttu-id="40db9-155">本属性包结构示例包含三个已定义  **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。</span><span class="sxs-lookup"><span data-stu-id="40db9-155">This example of the property bag structure contains three defined  **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="40db9-p111">设置属性包在上一加载项会话中进行保存后，在加载项当前会话期间，可以在加载项初始化时或其初始化后的任何时间点加载设置属性包。在会话期间，设置使用与您所创建的该类设置相对应的对象（ **Settings**、 **CustomProperties** 或 **RoamingSettings**）的 **get**、 **set** 和 **remove** 方法完全托管在内存中。</span><span class="sxs-lookup"><span data-stu-id="40db9-p111">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session. During the session, the settings are managed in entirely in memory using the  **get**,  **set**, and  **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**,  **CustomProperties**, or  **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="40db9-p112">若要将在加载项当前会话期间添加、更新或删除的任何内容暂留到存储位置，必须调用用于处理此类设置的相应对象的 **saveAsync** 方法。**get**、**set** 和 **remove** 方法仅对 settings 属性包的内存副本执行操作。如果加载项在没有调用 **saveAsync** 的情况下关闭，那么在相应会话期间对设置所做的任何更改都会丢失。</span><span class="sxs-lookup"><span data-stu-id="40db9-p112">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the  **saveAsync** method of the corresponding object used to work with that kind of settings. The **get**,  **set**, and  **remove** methods operate only on the in-memory copy of the settings property bag. If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="40db9-161">如何每文档暂留内容和任务窗格加载项的加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="40db9-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="40db9-p113">要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](https://dev.office.com/reference/add-ins/shared/settings) 对象及其方法。使用 **Settings** 对象的方法创建的属性包仅供创建它的内容或任务窗格加载项的实例使用，并且只能从保存它的文档使用。</span><span class="sxs-lookup"><span data-stu-id="40db9-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](https://dev.office.com/reference/add-ins/shared/settings) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="40db9-p114">**Settings** 对象自动加载为 [Document](https://dev.office.com/reference/add-ins/shared/document) 对象的一部分，并在任务窗格或内容加载项激活时可用。实例化 **Document** 对象后，你可以使用 **Document** 对象的 [settings ](https://dev.office.com/reference/add-ins/shared/document.settings)属性访问 **Settings** 对象。在该会话的生命周期中，你只能使用 **Settings.get**、**Settings.set** 和 **Settings.remove** 方法从属性包的内存副本中读取、写入或删除保留的设置和加载项状态。</span><span class="sxs-lookup"><span data-stu-id="40db9-p114">The  **Settings** object is automatically loaded as part of the [Document](https://dev.office.com/reference/add-ins/shared/document) object, and is available when the task pane or content add-in is activated. After the **Document** object is instantiated, you can access the **Settings** object with the [settings](https://dev.office.com/reference/add-ins/shared/document.settings) property of the **Document** object. During the lifetime of the session, you can just use the **Settings.get**,  **Settings.set**, and  **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="40db9-167">由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) 方法。</span><span class="sxs-lookup"><span data-stu-id="40db9-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="40db9-168">创建或更新设置值</span><span class="sxs-lookup"><span data-stu-id="40db9-168">Creating or updating a setting value</span></span>

<span data-ttu-id="40db9-p115">以下代码示例演示如何使用 [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。</span><span class="sxs-lookup"><span data-stu-id="40db9-p115">The following code example shows how to use the [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="40db9-p116">如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。使用 **Settings.saveAsync** 方法可将新的或更新的设置保留到文档中。</span><span class="sxs-lookup"><span data-stu-id="40db9-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="40db9-174">获取设置的值</span><span class="sxs-lookup"><span data-stu-id="40db9-174">Getting the value of a setting</span></span>

<span data-ttu-id="40db9-p117">下面的示例演示如何使用 [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) 方法获取名为“themeColor”的设置值。**get** 方法的唯一参数是设置的 _name_（区分大小写）。</span><span class="sxs-lookup"><span data-stu-id="40db9-p117">The following example shows how use the [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="40db9-p118">**get** 方法返回之前为传入的设置 _name_ 保存的值。如果不存在该设置，那么方法返回 **null**。</span><span class="sxs-lookup"><span data-stu-id="40db9-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="40db9-179">删除设置</span><span class="sxs-lookup"><span data-stu-id="40db9-179">Removing a setting</span></span>

<span data-ttu-id="40db9-p119">下面的示例演示如何使用 [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) 方法删除名为“themeColor”的设置。**remove** 方法的唯一参数是设置的 _name_（区分大小写）。</span><span class="sxs-lookup"><span data-stu-id="40db9-p119">The following example shows how to use the [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="40db9-p120">如果不存在该设置，则不执行任何操作。使用 **Settings.saveAsync** 方法可保留文档中设置的删除操作。</span><span class="sxs-lookup"><span data-stu-id="40db9-p120">Nothing will happen if the setting does not exist. Use the  **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="40db9-184">保存设置</span><span class="sxs-lookup"><span data-stu-id="40db9-184">Saving your settings</span></span>

<span data-ttu-id="40db9-p121">若要保存当前会话中加载项对设置属性包内存副本所做的任何添加、更改或删除操作，必须调用 [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) 方法，将它们存储在文档中。**saveAsync** 方法的唯一参数是使用单个参数的回调函数 _callback_。</span><span class="sxs-lookup"><span data-stu-id="40db9-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="40db9-p122">完成此操作后，将执行作为  **callback** 参数传入 _saveAsync_ 方法中的匿名函数。回调的 _asyncResult_ 参数提供对包含操作状态的 **AsyncResult** 对象的访问。在此示例中，函数将检查 **AsyncResult.status** 属性，以查看保存操作成功还是失败，然后在加载项页中显示结果。</span><span class="sxs-lookup"><span data-stu-id="40db9-p122">The anonymous function passed into the  **saveAsync** method as the _callback_ parameter is executed when the operation is completed. The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation. In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="40db9-190">如何将自定义 XML 保存到文档</span><span class="sxs-lookup"><span data-stu-id="40db9-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="40db9-191">此部分是在 Word 中支持的 Office 常见 JavaScript API 上下文中介绍的自定义 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="40db9-191">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="40db9-192">主机专用 Excel JavaScript API 还提供对自定义 XML 部分的访问权限。</span><span class="sxs-lookup"><span data-stu-id="40db9-192">The host-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="40db9-193">Excel API 和编程模式有点不一样。</span><span class="sxs-lookup"><span data-stu-id="40db9-193">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="40db9-194">有关详细信息，请参阅 [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart)。</span><span class="sxs-lookup"><span data-stu-id="40db9-194">For more information, see [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart).</span></span>

<span data-ttu-id="40db9-195">如果需要存储的信息超过文档设置的大小限制或有结构化字符，还有一个额外的存储选项。</span><span class="sxs-lookup"><span data-stu-id="40db9-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="40db9-196">可以在 Word 任务窗格加载项中暂留自定义 XML 标记（而对于 Excel，则请参阅本部分顶部的注释）。</span><span class="sxs-lookup"><span data-stu-id="40db9-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="40db9-197">在 Word 中，使用 [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) 对象及其方法（同样，对于 Excel，请参阅上面的注释）。下面的代码创建自定义 XML 部分，并在页面的 div 中显示它的 ID 和内容。</span><span class="sxs-lookup"><span data-stu-id="40db9-197">In Word, you use the [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) object and its methods (Again, see the note above for Excel.) The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="40db9-198">请注意，XML 字符串中必须有 `xmlns` 属性。</span><span class="sxs-lookup"><span data-stu-id="40db9-198">Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);                    
                }
            );
        }
    );
}
```

<span data-ttu-id="40db9-199">若要检索自定义 XML 部分，请使用 [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) 方法，但 ID 是在创建 XML 部分时生成的 GUID，因此编码时无法知道 ID 是什么。</span><span class="sxs-lookup"><span data-stu-id="40db9-199">To retrieve a custom XML part, you use the [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is.</span></span> <span data-ttu-id="40db9-200">因此，最好是在创建 XML 部分时，立即将 XML 部分的 ID 存储为设置，并为它提供容易记住的密钥。</span><span class="sxs-lookup"><span data-stu-id="40db9-200">For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key.</span></span> <span data-ttu-id="40db9-201">下面的方法展示了如何执行此操作。</span><span class="sxs-lookup"><span data-stu-id="40db9-201">The following method shows how to do this.</span></span> <span data-ttu-id="40db9-202">（不过，处理自定义设置时，请参阅本文的前面部分，以详细了解相关信息和最佳做法）。</span><span class="sxs-lookup"><span data-stu-id="40db9-202">(But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="40db9-203">下面的代码展示了如何通过先从设置中获取 ID 来检索 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="40db9-203">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID'));
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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="40db9-204">如何在 Outlook 加载项用户邮箱中将设置保存为漫游设置</span><span class="sxs-lookup"><span data-stu-id="40db9-204">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="40db9-205">Outlook 外接程序可以使用 [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 对象来保存特定于用户邮箱的外接程序状态和设置数据。</span><span class="sxs-lookup"><span data-stu-id="40db9-205">An Outlook add-in can use the [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="40db9-206">此数据只能由该 Outlook 外接程序代表运行外接程序的用户访问。</span><span class="sxs-lookup"><span data-stu-id="40db9-206">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="40db9-207">数据存储在用户的 Exchange Server 邮箱中，并且可以在该用户登录到其帐户并运行 Outlook 外接程序时访问。</span><span class="sxs-lookup"><span data-stu-id="40db9-207">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="40db9-208">加载漫游设置</span><span class="sxs-lookup"><span data-stu-id="40db9-208">Loading roaming settings</span></span>


<span data-ttu-id="40db9-p127">Outlook 外接程序通常在 [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) 事件处理程序中加载漫游设置。以下 JavaScript 代码示例演示了如何加载现有漫游设置。</span><span class="sxs-lookup"><span data-stu-id="40db9-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="40db9-211">创建或分配漫游设置</span><span class="sxs-lookup"><span data-stu-id="40db9-211">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="40db9-p128">紧接着前面的示例，下面的  `setAppSetting` 函数演示如何使用 [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 方法通过当天的日期设置或更新名为 `cookie` 的设置。然后使用 [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 方法将所有漫游设置保存回 Exchange Server。</span><span class="sxs-lookup"><span data-stu-id="40db9-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method.</span></span>


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

<span data-ttu-id="40db9-p129">**saveAsync** 方法可异步保存漫游设置，并采用一个可选的回调函数。此代码示例将名为 `saveMyAppSettingsCallback` 的回调函数传递到 **saveAsync** 方法。异步调用返回时，`saveMyAppSettingsCallback` 函数的 _asyncResult_ 参数提供对 [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) 对象的访问权限，该对象用于确定通过 **AsyncResult.status** 属性的操作是成功还是失败。</span><span class="sxs-lookup"><span data-stu-id="40db9-p129">The  **saveAsync** method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="40db9-217">删除漫游设置</span><span class="sxs-lookup"><span data-stu-id="40db9-217">Removing a roaming setting</span></span>


<span data-ttu-id="40db9-218">进一步展开前面的示例，以下  `removeAppSetting` 函数演示了如何使用 [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。</span><span class="sxs-lookup"><span data-stu-id="40db9-218">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="40db9-219">如何按项目将 Outlook 外接程序的设置保存为自定义属性</span><span class="sxs-lookup"><span data-stu-id="40db9-219">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="40db9-p130">自定义属性允许 Outlook 外接程序存储其使用的有关项目的信息。例如，如果 Outlook 外接程序根据邮件中的会议建议创建约会，则可以使用自定义属性存储创建了会议的事实。这确保了如果再次打开邮件，Outlook 外接程序不再可供创建约会。</span><span class="sxs-lookup"><span data-stu-id="40db9-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="40db9-p131">在您将自定义属性用于特定邮件、约会或会议请求项目之前，必须通过调用  [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) 对象的 **loadCustomPropertiesAsync** 方法将属性加载到内存中。如果为当前项目设置了任何自定义属性，此时会从 Exchanger Server 加载这些属性。在您加载了属性以后，可以使用 [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) 对象的 [set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 和 **get** 方法添加、更新和检索内存中的属性。要保存对于项目的自定义属性所做的任何更改，必须使用 [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) 方法在 Exchanger Server上保留对项目所做的更改。</span><span class="sxs-lookup"><span data-stu-id="40db9-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](https://dev.office.com/reference/add-ins/outlook/CustomProperties) and [get](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="40db9-227">自定义属性示例</span><span class="sxs-lookup"><span data-stu-id="40db9-227">Custom properties example</span></span>

<span data-ttu-id="40db9-p132">下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的函数。可以将此示例用作使用自定义属性的 Outlook 外接程序的起点。</span><span class="sxs-lookup"><span data-stu-id="40db9-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="40db9-230">使用这些函数的 Outlook 外接程序通过对  `_customProps` 变量调用 **get** 方法来检索任何自定义属性，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="40db9-230">An Outlook add-in that uses these functions retrieves any custom properties by calling the  **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="40db9-231">此示例包括以下函数：</span><span class="sxs-lookup"><span data-stu-id="40db9-231">This example includes the following functions:</span></span>



|<span data-ttu-id="40db9-232">**函数名称**</span><span class="sxs-lookup"><span data-stu-id="40db9-232">**Function name**</span></span>|<span data-ttu-id="40db9-233">**说明**</span><span class="sxs-lookup"><span data-stu-id="40db9-233">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="40db9-234">从 Exchange 服务器初始化外接程序并加载当前项目的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="40db9-234">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="40db9-235">获取从 Exchange 服务器返回的自定义属性并将其保存以供后续之用。</span><span class="sxs-lookup"><span data-stu-id="40db9-235">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="40db9-236">设置或更新特定属性，然后将更改保存到 Exchange 服务器。</span><span class="sxs-lookup"><span data-stu-id="40db9-236">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="40db9-237">删除特定的属性，然后保留删除操作到 Exchange 服务器。</span><span class="sxs-lookup"><span data-stu-id="40db9-237">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="40db9-238">回调 `updateProperty` 和 `removeProperty` 函数中的 **saveAsync** 方法。</span><span class="sxs-lookup"><span data-stu-id="40db9-238">Callback for calls to the  **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="40db9-239">另请参阅</span><span class="sxs-lookup"><span data-stu-id="40db9-239">See also</span></span>

- [<span data-ttu-id="40db9-240">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="40db9-240">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="40db9-241">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="40db9-241">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="40db9-242">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="40db9-242">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
