---
title: ??????????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b4d1cdf2ce127d140153b6db02bc9a337a37bb5d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="5b2ad-102">??????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="5b2ad-p101">Office ??????????????????????? Web ????????????????????????????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="5b2ad-107">????? Office ? JavaScript API ??????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="5b2ad-108">??????????????????????-????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="5b2ad-109">?????????? XML?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-109">Custom XML stored in the document.</span></span>
    
- <span data-ttu-id="5b2ad-110">?????????????????? Cookie ? HTML5 Web ???[localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) ? [sessionStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/sessionStorage)??</span><span class="sxs-lookup"><span data-stu-id="5b2ad-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/sessionStorage)).</span></span>
    
<span data-ttu-id="5b2ad-p102">????????????? Office ? JavaScript API ???????????????? Cookie ? Web ????????? [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="5b2ad-113">????? Office ? JavaScript API ??????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="5b2ad-p103">??? Office ? JavaScript API ?????????????????? [Settings](https://dev.office.com/reference/add-ins/shared/settings)? [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ? [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) ?????????????????????????????????? [Id](https://dev.office.com/reference/add-ins/manifest/id) ????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p103">The JavaScript API for Office provides the [Settings](https://dev.office.com/reference/add-ins/shared/settings), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings), and [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](https://dev.office.com/reference/add-ins/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="5b2ad-116">**??**</span><span class="sxs-lookup"><span data-stu-id="5b2ad-116">**Object**</span></span>|<span data-ttu-id="5b2ad-117">**????????**</span><span class="sxs-lookup"><span data-stu-id="5b2ad-117">**Add-in type support**</span></span>|<span data-ttu-id="5b2ad-118">**????**</span><span class="sxs-lookup"><span data-stu-id="5b2ad-118">**Storage location**</span></span>|<span data-ttu-id="5b2ad-119">**Office ????**</span><span class="sxs-lookup"><span data-stu-id="5b2ad-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5b2ad-120">??</span><span class="sxs-lookup"><span data-stu-id="5b2ad-120">Settings</span></span>](https://dev.office.com/reference/add-ins/shared/settings)|<span data-ttu-id="5b2ad-121">???????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-121">content and task pane</span></span>|<span data-ttu-id="5b2ad-122">??????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="5b2ad-p104">**?????**???? **Settings** ?????????????????? (PII)??????????????????????????????????????????????????????? PII ??????????????? PII ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="5b2ad-126">Word?Excel ? PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5b2ad-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="5b2ad-p105">**???** Project 2013 ??????????????????????? **Settings** API??????? Project???? Office ???????????????????? Cookie ? Web ???????????????????? [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="5b2ad-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="5b2ad-130">RoamingSettings</span></span>](https://dev.office.com/reference/add-ins/outlook/RoamingSettings)|<span data-ttu-id="5b2ad-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b2ad-131">Outlook</span></span>|<span data-ttu-id="5b2ad-132">????????? Exchange ??????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="5b2ad-133">Outlook ?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="5b2ad-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b2ad-134">Outlook</span></span>|
|[<span data-ttu-id="5b2ad-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="5b2ad-135">CustomProperties</span></span>](https://dev.office.com/reference/add-ins/outlook/CustomProperties)|<span data-ttu-id="5b2ad-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b2ad-136">Outlook</span></span>|<span data-ttu-id="5b2ad-p106">??????????????????? Outlook ????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="5b2ad-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="5b2ad-139">Outlook</span></span>|
|[<span data-ttu-id="5b2ad-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5b2ad-140">CustomXMLParts</span></span>](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts)|<span data-ttu-id="5b2ad-141">????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-141">task pane</span></span>|<span data-ttu-id="5b2ad-p107">???????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="5b2ad-p108">**?????**????????????????? (PII) ?????? XML ?????????????????????????????????????????????????????????? PII ??????????????? PII ??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="5b2ad-147">Word??? Office JavaScript ?? API??Excel??????? Excel JavaScript API?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="5b2ad-148">??????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="5b2ad-149">??????? Office ?? JavaScript API ??????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-149">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="5b2ad-150">???? Excel JavaScript API ???????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-150">The host-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="5b2ad-151">Excel API ???????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-151">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="5b2ad-152">?????????? [Excel SettingCollection](https://dev.office.com/reference/add-ins/excel/settingcollection)?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-152">For more information, see [Excel SettingCollection](https://dev.office.com/reference/add-ins/excel/settingcollection).</span></span>

<span data-ttu-id="5b2ad-p110">??????  **Settings**? **CustomProperties** ? **RoamingSettings** ??????????????????? JavaScript ????? (JSON) ???????/??????????????? **string**???????? JavaScript  **string**? **number**? **date** ? **object**?????  **function**?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p110">Internally, the data in the property bag accessed with the  **Settings**,  **CustomProperties**, or  **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a **string**, and the stored value can be a JavaScript  **string**,  **number**,  **date**, or  **object**, but not a  **function**.</span></span>

<span data-ttu-id="5b2ad-155">???????????????  **string** ????? `firstName`? `location` ? `defaultView`?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-155">This example of the property bag structure contains three defined  **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="5b2ad-p111">???????????????????????????????????????????????????????????????????????????????????????? **Settings**? **CustomProperties** ? **RoamingSettings**?? **get**? **set** ? **remove** ???????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p111">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session. During the session, the settings are managed in entirely in memory using the  **get**,  **set**, and  **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**,  **CustomProperties**, or  **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="5b2ad-p112">???????????????????????????????????????????????????? **saveAsync** ???**get**?**set** ? **remove** ???? settings ??????????????????????? **saveAsync** ???????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p112">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the  **saveAsync** method of the corresponding object used to work with that kind of settings. The **get**,  **set**, and  **remove** methods operate only on the in-memory copy of the settings property bag. If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="5b2ad-161">??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="5b2ad-p113">??? Word?Excel ? PowerPoint ???????????????????????? [Settings](https://dev.office.com/reference/add-ins/shared/settings) ????????? **Settings** ???????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](https://dev.office.com/reference/add-ins/shared/settings) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="5b2ad-p114">**Settings** ??????? [Document](https://dev.office.com/reference/add-ins/shared/document) ???????????????????????????? **Document** ????????? **Document** ??? [settings ](https://dev.office.com/reference/add-ins/shared/document.settings)???? **Settings** ??????????????????? **Settings.get**?**Settings.set** ? **Settings.remove** ????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p114">The  **Settings** object is automatically loaded as part of the [Document](https://dev.office.com/reference/add-ins/shared/document) object, and is available when the task pane or content add-in is activated. After the **Document** object is instantiated, you can access the **Settings** object with the [settings](https://dev.office.com/reference/add-ins/shared/document.settings) property of the **Document** object. During the lifetime of the session, you can just use the **Settings.get**,  **Settings.set**, and  **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="5b2ad-167">?? set ? remove ??????????????????????????????????????????? [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) ???</span><span class="sxs-lookup"><span data-stu-id="5b2ad-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="5b2ad-168">????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-168">Creating or updating a setting value</span></span>

<span data-ttu-id="5b2ad-p115">???????????? [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) ?????? `'themeColor'` ??? `'green'` ????set ??????????????????? _name_ (Id)????????????????? _value_?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p115">The following code example shows how to use the [Settings.set](https://dev.office.com/reference/add-ins/shared/settings.set) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="5b2ad-p116">????????????????????????????????????????? **Settings.saveAsync** ???????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="5b2ad-174">??????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-174">Getting the value of a setting</span></span>

<span data-ttu-id="5b2ad-p117">??????????? [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) ???????themeColor??????**get** ??????????? _name_????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p117">The following example shows how use the [Settings.get](https://dev.office.com/reference/add-ins/shared/settings.get) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="5b2ad-p118">**get** ???????????? _name_ ???????????????????? **null**?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="5b2ad-179">????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-179">Removing a setting</span></span>

<span data-ttu-id="5b2ad-p119">??????????? [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) ???????themeColor?????**remove** ??????????? _name_????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p119">The following example shows how to use the [Settings.remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="5b2ad-p120">???????????????????? **Settings.saveAsync** ????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p120">Nothing will happen if the setting does not exist. Use the  **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="5b2ad-184">????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-184">Saving your settings</span></span>

<span data-ttu-id="5b2ad-p121">?????????????????????????????????????????? [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) ?????????????**saveAsync** ??????????????????? _callback_?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="5b2ad-p122">????????????  **callback** ???? _saveAsync_ ???????????? _asyncResult_ ???????????? **AsyncResult** ????????????????? **AsyncResult.status** ??????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p122">The anonymous function passed into the  **saveAsync** method as the _callback_ parameter is executed when the operation is completed. The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation. In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="5b2ad-190">?????? XML ?????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="5b2ad-191">????? Word ???? Office ?? JavaScript API ?????????? XML ???</span><span class="sxs-lookup"><span data-stu-id="5b2ad-191">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="5b2ad-192">???? Excel JavaScript API ??????? XML ????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-192">The host-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="5b2ad-193">Excel API ???????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-193">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="5b2ad-194">?????????? [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart)?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-194">For more information, see [Excel CustomXmlPart](https://dev.office.com/reference/add-ins/excel/customxmlpart).</span></span>

<span data-ttu-id="5b2ad-195">????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="5b2ad-196">??? Word ????????????? XML ?????? Excel???????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="5b2ad-197">? Word ???? [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) ???????????? Excel????????????????????? XML ???????? div ????? ID ????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-197">In Word, you use the [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) object and its methods (Again, see the note above for Excel.) The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="5b2ad-198">????XML ??????? `xmlns` ???</span><span class="sxs-lookup"><span data-stu-id="5b2ad-198">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="5b2ad-199">??????? XML ?????? [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) ???? ID ???? XML ?????? GUID?????????? ID ????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-199">To retrieve a custom XML part, you use the [getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is.</span></span> <span data-ttu-id="5b2ad-200">????????? XML ??????? XML ??? ID ???????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-200">For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key.</span></span> <span data-ttu-id="5b2ad-201">????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-201">The following method shows how to do this.</span></span> <span data-ttu-id="5b2ad-202">????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-202">(But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="5b2ad-203">??????????????????? ID ??? XML ???</span><span class="sxs-lookup"><span data-stu-id="5b2ad-203">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="5b2ad-204">????????? Outlook ?????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-204">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="5b2ad-205"> Outlook ???????? [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-205">An Outlook add-in can use the [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="5b2ad-206">??????? Outlook ??????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-206">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="5b2ad-207">???????? Exchange Server ????????????????????? Outlook ????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-207">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="5b2ad-208">??????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-208">Loading roaming settings</span></span>


<span data-ttu-id="5b2ad-p127">Outlook ??????? [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) ???????????????? JavaScript ??????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="5b2ad-211">?????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-211">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="5b2ad-p128">????????????  `setAppSetting` ???????? [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ???????????????? `cookie` ???????? [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ???????????? Exchange Server?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method.</span></span>


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

<span data-ttu-id="5b2ad-p129">**saveAsync** ????????????????????????????????? `saveMyAppSettingsCallback` ???????? **saveAsync** ???????????`saveMyAppSettingsCallback` ??? _asyncResult_ ????? [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) ????????????????? **AsyncResult.status** ?????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p129">The  **saveAsync** method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](https://dev.office.com/reference/add-ins/outlook/simple-types) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="5b2ad-217">??????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-217">Removing a roaming setting</span></span>


<span data-ttu-id="5b2ad-218">?????????????  `removeAppSetting` ????????? [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ???? `cookie` ????????????? Exchange Server?</span><span class="sxs-lookup"><span data-stu-id="5b2ad-218">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="5b2ad-219">?????? Outlook ???????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-219">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="5b2ad-p130">??????? Outlook ??????????????????????? Outlook ?????????????????????????????????????????????????????Outlook ?????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="5b2ad-p131">?????????????????????????????????  [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) ??? **loadCustomPropertiesAsync** ?????????????????????????????????? Exchanger Server ????????????????????? [CustomProperties](https://dev.office.com/reference/add-ins/outlook/CustomProperties) ??? [set](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ? **get** ?????????????????????????????????????????? [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) ??? Exchanger Server????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](https://dev.office.com/reference/add-ins/outlook/CustomProperties) and [get](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](https://dev.office.com/reference/add-ins/outlook/CustomProperties) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="5b2ad-227">???????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-227">Custom properties example</span></span>

<span data-ttu-id="5b2ad-p132">??????????????? Outlook ????????????????????????????? Outlook ????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="5b2ad-230">??????? Outlook ???????  `_customProps` ???? **get** ??????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-230">An Outlook add-in that uses these functions retrieves any custom properties by calling the  **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="5b2ad-231">??????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-231">This example includes the following functions:</span></span>



|<span data-ttu-id="5b2ad-232">**????**</span><span class="sxs-lookup"><span data-stu-id="5b2ad-232">**Function name**</span></span>|<span data-ttu-id="5b2ad-233">**??**</span><span class="sxs-lookup"><span data-stu-id="5b2ad-233">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="5b2ad-234">? Exchange ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-234">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="5b2ad-235">??? Exchange ???????????????????????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-235">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="5b2ad-236">?????????????????? Exchange ????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-236">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="5b2ad-237">????????????????? Exchange ????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-237">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="5b2ad-238">?? `updateProperty` ? `removeProperty` ???? **saveAsync** ???</span><span class="sxs-lookup"><span data-stu-id="5b2ad-238">Callback for calls to the  **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="5b2ad-239">????</span><span class="sxs-lookup"><span data-stu-id="5b2ad-239">See also</span></span>

- [<span data-ttu-id="5b2ad-240">????? Office ? JavaScript API</span><span class="sxs-lookup"><span data-stu-id="5b2ad-240">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="5b2ad-241">Outlook ???</span><span class="sxs-lookup"><span data-stu-id="5b2ad-241">Outlook add-ins</span></span>](https://docs.microsoft.com/en-us/outlook/add-ins/)
- [<span data-ttu-id="5b2ad-242">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="5b2ad-242">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
