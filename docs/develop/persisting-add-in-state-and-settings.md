---
title: 暂留加载项状态和设置
description: 了解如何在浏览器控件的无状态环境中保存运行的 Office 外接程序 web 应用程序中的数据。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 90e072d638a3a598610c4bcbb2e6af07f1196467
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087950"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="2542f-103">暂留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="2542f-103">Persisting add-in state and settings</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="2542f-p101">Office 加载项实质上是在浏览器控件的无状态环境中运行的 Web 应用。因此，加载项可能需要暂留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能有需要在下一次初始化时保存和重新加载的自定义设置或其他值（如用户的首选视图或默认位置）。为此，可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="2542f-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="2542f-108">使用 Office JavaScript API 的成员，将数据存储为以下任意一种：</span><span class="sxs-lookup"><span data-stu-id="2542f-108">Use members of the Office JavaScript API that store data as either:</span></span>
  - <span data-ttu-id="2542f-109">在依赖加载项类型的位置上存储的属性包中的名称-数值对。</span><span class="sxs-lookup"><span data-stu-id="2542f-109">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
  - <span data-ttu-id="2542f-110">在文档中存储的自定义 XML。</span><span class="sxs-lookup"><span data-stu-id="2542f-110">Custom XML stored in the document.</span></span>

- <span data-ttu-id="2542f-111">使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。</span><span class="sxs-lookup"><span data-stu-id="2542f-111">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="2542f-112">本文重点介绍如何使用 Office JavaScript API 来保留加载项状态。</span><span class="sxs-lookup"><span data-stu-id="2542f-112">This article focuses on how to use the Office JavaScript API to persist add-in state.</span></span> <span data-ttu-id="2542f-113">有关使用浏览器 Cookie 和 Web 存储的示例，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。</span><span class="sxs-lookup"><span data-stu-id="2542f-113">For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a><span data-ttu-id="2542f-114">使用 Office JavaScript API 保留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="2542f-114">Persisting add-in state and settings with the Office JavaScript API</span></span>

<span data-ttu-id="2542f-115">Office JavaScript API 提供了 [设置](/javascript/api/office/office.settings)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)和 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象，用于按下表所述在会话中保存外接程序状态。</span><span class="sxs-lookup"><span data-stu-id="2542f-115">The Office JavaScript API provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="2542f-116">在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](../reference/manifest/id.md) 相关联。</span><span class="sxs-lookup"><span data-stu-id="2542f-116">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="2542f-117">**对象**</span><span class="sxs-lookup"><span data-stu-id="2542f-117">**Object**</span></span>|<span data-ttu-id="2542f-118">**外接程序类型支持**</span><span class="sxs-lookup"><span data-stu-id="2542f-118">**Add-in type support**</span></span>|<span data-ttu-id="2542f-119">**存储位置**</span><span class="sxs-lookup"><span data-stu-id="2542f-119">**Storage location**</span></span>|<span data-ttu-id="2542f-120">**Office 应用程序支持**</span><span class="sxs-lookup"><span data-stu-id="2542f-120">**Office application support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2542f-121">Settings</span><span class="sxs-lookup"><span data-stu-id="2542f-121">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="2542f-122">内容和任务窗格</span><span class="sxs-lookup"><span data-stu-id="2542f-122">content and task pane</span></span>|<span data-ttu-id="2542f-123">加载项使用的文档、电子表格或演示文稿。</span><span class="sxs-lookup"><span data-stu-id="2542f-123">The document, spreadsheet, or presentation the add-in is working with.</span></span> <span data-ttu-id="2542f-124">内容和任务窗格加载项设置供创建它们的加载项使用，且能从保存它们的文档访问。</span><span class="sxs-lookup"><span data-stu-id="2542f-124">Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="2542f-p105">**重要说明：** 不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。您应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。</span><span class="sxs-lookup"><span data-stu-id="2542f-p105">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="2542f-128">Word、Excel 或 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2542f-128">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="2542f-129">**注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。</span><span class="sxs-lookup"><span data-stu-id="2542f-129">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings.</span></span> <span data-ttu-id="2542f-130">但是，对于在 Project (以及其他 Office 客户端应用程序中运行的加载项) ，您可以使用浏览器 cookie 或 web 存储等技术。</span><span class="sxs-lookup"><span data-stu-id="2542f-130">However, for add-ins running in Project (as well as other Office client applications) you can use techniques such as browser cookies or web storage.</span></span> <span data-ttu-id="2542f-131">有关详细信息，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。</span><span class="sxs-lookup"><span data-stu-id="2542f-131">For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="2542f-132">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2542f-132">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="2542f-133">Outlook</span><span class="sxs-lookup"><span data-stu-id="2542f-133">Outlook</span></span>|<span data-ttu-id="2542f-134">安装了加载项的用户 Exchange 服务器邮箱。</span><span class="sxs-lookup"><span data-stu-id="2542f-134">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="2542f-135">由于这些设置存储在用户的服务器邮箱中，因此它们可以与用户 "漫游"，并在加载项运行在任何受支持的 Office 客户端应用程序或浏览器访问该用户的邮箱的上下文中时，使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="2542f-135">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported Office client application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="2542f-136">Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。</span><span class="sxs-lookup"><span data-stu-id="2542f-136">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="2542f-137">Outlook</span><span class="sxs-lookup"><span data-stu-id="2542f-137">Outlook</span></span>|
|[<span data-ttu-id="2542f-138">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="2542f-138">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="2542f-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="2542f-139">Outlook</span></span>|<span data-ttu-id="2542f-p108">加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。</span><span class="sxs-lookup"><span data-stu-id="2542f-p108">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="2542f-142">Outlook</span><span class="sxs-lookup"><span data-stu-id="2542f-142">Outlook</span></span>|
|[<span data-ttu-id="2542f-143">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2542f-143">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="2542f-144">任务窗格</span><span class="sxs-lookup"><span data-stu-id="2542f-144">task pane</span></span>|<span data-ttu-id="2542f-p109">加载项要使用的文档、电子表格或演示文稿。任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。</span><span class="sxs-lookup"><span data-stu-id="2542f-p109">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="2542f-p110">**重要说明：** 请勿将密码和其他敏感的个人身份信息 (PII) 存储在自定义 XML 部分中。虽然保存的数据对最终用户不可见，但它存储为文档的一部分，可通过直接读取文档的文件格式进行访问。应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在服务器上，且服务器将加载项托管为用户保护资源。</span><span class="sxs-lookup"><span data-stu-id="2542f-p110">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="2542f-150">Word (使用 Office JavaScript 通用 API) Excel (使用特定于应用程序的 Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="2542f-150">Word (using the Office JavaScript Common API) Excel (using the application-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="2542f-151">设置数据在运行时托管在内存中</span><span class="sxs-lookup"><span data-stu-id="2542f-151">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="2542f-152">下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。</span><span class="sxs-lookup"><span data-stu-id="2542f-152">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="2542f-153">应用程序特定的 Excel JavaScript API 还提供对自定义设置的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2542f-153">The application-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="2542f-154">Excel API 和编程模式有点不一样。</span><span class="sxs-lookup"><span data-stu-id="2542f-154">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="2542f-155">有关详细信息，请参阅 [Excel SettingCollection](/javascript/api/excel/excel.settingcollection)。</span><span class="sxs-lookup"><span data-stu-id="2542f-155">For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="2542f-156">在内部，使用、或对象访问的属性包中的数据 `Settings` `CustomProperties` `RoamingSettings` 存储为序列化的 JavaScript 对象表示法 (JSON) 对象，其中包含名称/值对。</span><span class="sxs-lookup"><span data-stu-id="2542f-156">Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="2542f-157">每个值 (键) 的名称必须为 `string` ，并且存储的值可以是 JavaScript `string` 、 `number` 、 `date` 或 `object` ，但不能是 **函数**。</span><span class="sxs-lookup"><span data-stu-id="2542f-157">The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.</span></span>

<span data-ttu-id="2542f-158">本属性包结构示例包含三个已定义 **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。</span><span class="sxs-lookup"><span data-stu-id="2542f-158">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="2542f-159">在前一个加载项会话中保存设置属性包之后，可以在加载项的当前会话中初始化加载项时或在之后的任何时间加载该设置属性包。</span><span class="sxs-lookup"><span data-stu-id="2542f-159">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="2542f-160">在会话过程中，将使用与所创建的设置类型相对应的对象的、和方法来完全在内存中管理这些设置 `get` `set` `remove` (**设置**、 **CustomProperties** 或 **RoamingSettings**) 。</span><span class="sxs-lookup"><span data-stu-id="2542f-160">During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you are creating (**Settings**, **CustomProperties**, or **RoamingSettings**).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2542f-161">若要将在外接程序的当前会话过程中所做的任何添加、更新或删除操作保存到存储位置，必须调用 `saveAsync` 与该类型的设置一起使用的相应对象的方法。</span><span class="sxs-lookup"><span data-stu-id="2542f-161">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="2542f-162">`get`、 `set` 和 `remove` 方法仅在设置属性包的内存中副本上运行。</span><span class="sxs-lookup"><span data-stu-id="2542f-162">The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="2542f-163">如果外接程序在未呼叫的情况 `saveAsync` 下关闭，则在该会话期间对设置所做的任何更改都将丢失。</span><span class="sxs-lookup"><span data-stu-id="2542f-163">If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.</span></span>

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="2542f-164">如何按文档暂留内容和任务窗格加载项的加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="2542f-164">How to save add-in state and settings per document for content and task pane add-ins</span></span>

<span data-ttu-id="2542f-165">要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](/javascript/api/office/office.settings) 对象及其方法。</span><span class="sxs-lookup"><span data-stu-id="2542f-165">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods.</span></span> <span data-ttu-id="2542f-166">使用对象的方法创建的属性包 `Settings` 仅可用于创建它的内容或任务窗格外接程序的实例，并且只能从保存它的文档中获取。</span><span class="sxs-lookup"><span data-stu-id="2542f-166">The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="2542f-167">`Settings`对象将作为[Document](/javascript/api/office/office.document)对象的一部分自动加载，并在激活任务窗格或内容加载项时可用。</span><span class="sxs-lookup"><span data-stu-id="2542f-167">The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="2542f-168">在 `Document` 实例化对象之后，可以 `Settings` 使用对象的 [settings](/javascript/api/office/office.document#settings) 属性访问对象 `Document` 。</span><span class="sxs-lookup"><span data-stu-id="2542f-168">After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#settings) property of the `Document` object.</span></span> <span data-ttu-id="2542f-169">在会话的生存期期间，您只需使用 `Settings.get` 、 `Settings.set` 和方法，即可 `Settings.remove` 从属性包的内存中副本中读取、写入或删除保留的设置和加载项状态。</span><span class="sxs-lookup"><span data-stu-id="2542f-169">During the lifetime of the session, you can just use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="2542f-170">由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法。</span><span class="sxs-lookup"><span data-stu-id="2542f-170">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>

### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="2542f-171">创建或更新设置值</span><span class="sxs-lookup"><span data-stu-id="2542f-171">Creating or updating a setting value</span></span>

<span data-ttu-id="2542f-p117">以下代码示例演示如何使用 [Settings.set](/javascript/api/office/office.settings#set-name--value-) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。</span><span class="sxs-lookup"><span data-stu-id="2542f-p117">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>

```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="2542f-175">如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。</span><span class="sxs-lookup"><span data-stu-id="2542f-175">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist.</span></span> <span data-ttu-id="2542f-176">使用 `Settings.saveAsync` 方法可将新的或更新的设置保存到文档中。</span><span class="sxs-lookup"><span data-stu-id="2542f-176">Use the `Settings.saveAsync` method to persist the new or updated settings to the document.</span></span>

### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="2542f-177">获取设置的值</span><span class="sxs-lookup"><span data-stu-id="2542f-177">Getting the value of a setting</span></span>

<span data-ttu-id="2542f-178">下面的示例演示如何使用 [Settings.get](/javascript/api/office/office.settings#get-name-) 方法获取名为"themeColor"的设置值。</span><span class="sxs-lookup"><span data-stu-id="2542f-178">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor".</span></span> <span data-ttu-id="2542f-179">方法的唯一参数 `get` 是设置的区分大小写的 _名称_ 。</span><span class="sxs-lookup"><span data-stu-id="2542f-179">The only parameter of the `get` method is the case-sensitive _name_ of the setting.</span></span>

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 <span data-ttu-id="2542f-180">该 `get` 方法返回之前为传入的设置 _名称_ 保存的值。</span><span class="sxs-lookup"><span data-stu-id="2542f-180">The `get` method returns the value that was previously saved for the setting _name_ that was passed in.</span></span> <span data-ttu-id="2542f-181">如果不存在该设置，那么方法返回 **null**。</span><span class="sxs-lookup"><span data-stu-id="2542f-181">If the setting doesn't exist, the method returns **null**.</span></span>

### <a name="removing-a-setting"></a><span data-ttu-id="2542f-182">删除设置</span><span class="sxs-lookup"><span data-stu-id="2542f-182">Removing a setting</span></span>

<span data-ttu-id="2542f-183">下面的示例演示如何使用 [Settings.remove](/javascript/api/office/office.settings#remove-name-) 方法删除名为"themeColor"的设置。</span><span class="sxs-lookup"><span data-stu-id="2542f-183">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor".</span></span> <span data-ttu-id="2542f-184">方法的唯一参数 `remove` 是设置的区分大小写的 _名称_ 。</span><span class="sxs-lookup"><span data-stu-id="2542f-184">The only parameter of the `remove` method is the case-sensitive _name_ of the setting.</span></span>

```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="2542f-185">如果不存在该设置，则不执行任何操作。</span><span class="sxs-lookup"><span data-stu-id="2542f-185">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="2542f-186">使用 `Settings.saveAsync` 方法可将设置从文档中永久删除。</span><span class="sxs-lookup"><span data-stu-id="2542f-186">Use the `Settings.saveAsync` method to persist removal of the setting from the document.</span></span>

### <a name="saving-your-settings"></a><span data-ttu-id="2542f-187">保存设置</span><span class="sxs-lookup"><span data-stu-id="2542f-187">Saving your settings</span></span>

<span data-ttu-id="2542f-188">若要保存当前会话中加载项对设置属性包内存副本所做的任意添加、更改或删除操作，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法将它们存储在文档中。</span><span class="sxs-lookup"><span data-stu-id="2542f-188">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document.</span></span> <span data-ttu-id="2542f-189">该方法的唯一参数 `saveAsync` 是 _callback_，它是一个具有单个参数的回调函数。</span><span class="sxs-lookup"><span data-stu-id="2542f-189">The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.</span></span>

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

<span data-ttu-id="2542f-190">在 `saveAsync` 操作完成时，会执行匿名函数作为 _callback_ 参数传入方法。</span><span class="sxs-lookup"><span data-stu-id="2542f-190">The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="2542f-191">此回调的 _asyncResult_ 参数提供对 `AsyncResult` 包含操作状态的对象的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2542f-191">The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation.</span></span> <span data-ttu-id="2542f-192">在此示例中，函数检查 `AsyncResult.status` 属性以查看保存操作是成功还是失败，然后在加载项页面中显示结果。</span><span class="sxs-lookup"><span data-stu-id="2542f-192">In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="2542f-193">如何将自定义 XML 保存到文档</span><span class="sxs-lookup"><span data-stu-id="2542f-193">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="2542f-194">此部分是在 Word 中支持的 Office 常见 JavaScript API 上下文中介绍的自定义 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="2542f-194">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="2542f-195">特定于应用程序的 Excel JavaScript API 还提供对自定义 XML 部件的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2542f-195">The application-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="2542f-196">Excel API 和编程模式有点不一样。</span><span class="sxs-lookup"><span data-stu-id="2542f-196">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="2542f-197">有关详细信息，请参阅 [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart)。</span><span class="sxs-lookup"><span data-stu-id="2542f-197">For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="2542f-198">当您需要存储超出文档设置大小限制或具有结构化字符的信息时，还会有一个额外的存储选项。</span><span class="sxs-lookup"><span data-stu-id="2542f-198">There is an additional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="2542f-199">可以在 Word 的任务窗格加载项中暂留自定义 XML 标记（对于 Excel，但请参阅本节顶部的注释）。</span><span class="sxs-lookup"><span data-stu-id="2542f-199">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="2542f-200">在 Word 中，可以使用 [CustomXmlPart](/javascript/api/office/office.customxmlpart) 对象及其方法（同样，请参阅上面的 Excel 注释）。</span><span class="sxs-lookup"><span data-stu-id="2542f-200">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="2542f-201">以下代码将创建自定义 XML 部件，并在页面的 divs 中显示其 ID 及内容。</span><span class="sxs-lookup"><span data-stu-id="2542f-201">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="2542f-202">请注意，XML 字符串中必须有一个 `xmlns` 属性。</span><span class="sxs-lookup"><span data-stu-id="2542f-202">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="2542f-p127">若要检索自定义 XML 部分，请使用 [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) 方法，但 ID 是在创建 XML 部分时生成的 GUID，因此编码时无法知道 ID 是什么。 因此，最好是在创建 XML 部分时，立即将 XML 部分的 ID 存储为设置，并为它提供容易记住的密钥。 下面的方法展示了如何执行此操作。 （不过，处理自定义设置时，请参阅本文的前面部分，以详细了解相关信息和最佳做法）。</span><span class="sxs-lookup"><span data-stu-id="2542f-p127">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="2542f-207">下面的代码展示了如何通过先从设置中获取 ID 来检索 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="2542f-207">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a><span data-ttu-id="2542f-208">如何在 Outlook 加载项中保存设置</span><span class="sxs-lookup"><span data-stu-id="2542f-208">How to save settings in an Outlook add-in</span></span>

<span data-ttu-id="2542f-209">有关如何在 Outlook 外接程序中保存设置的信息，请参阅 [Manage state and settings For outlook 外接程序](../outlook/manage-state-and-settings-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="2542f-209">For information about how to save settings in an Outlook add-in, see [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2542f-210">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2542f-210">See also</span></span>

- [<span data-ttu-id="2542f-211">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="2542f-211">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="2542f-212">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="2542f-212">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="2542f-213">管理 Outlook 外接程序的状态和设置</span><span class="sxs-lookup"><span data-stu-id="2542f-213">Manage state and settings for an Outlook add-in</span></span>](../outlook/manage-state-and-settings-outlook.md)
- [<span data-ttu-id="2542f-214">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="2542f-214">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
