---
title: 暂留加载项状态和设置
description: 了解如何将数据保留Office浏览器控件的无状态环境中运行的外接程序 Web 应用程序中。
ms.date: 03/23/2021
localization_priority: Normal
ms.openlocfilehash: a5a54a07abfeefda39d24e635773bfd808b59c25
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349768"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="bfd28-103">暂留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-103">Persisting add-in state and settings</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="bfd28-p101">Office 加载项实质上是在浏览器控件的无状态环境中运行的 Web 应用。因此，加载项可能需要暂留数据，以维护各个使用加载项的会话中某些操作或功能的连续性。例如，加载项可能有需要在下一次初始化时保存和重新加载的自定义设置或其他值（如用户的首选视图或默认位置）。为此，可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="bfd28-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="bfd28-108">使用存储数据的 Office JavaScript API 的成员：</span><span class="sxs-lookup"><span data-stu-id="bfd28-108">Use members of the Office JavaScript API that store data as either:</span></span>
  - <span data-ttu-id="bfd28-109">在依赖加载项类型的位置上存储的属性包中的名称-数值对。</span><span class="sxs-lookup"><span data-stu-id="bfd28-109">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
  - <span data-ttu-id="bfd28-110">在文档中存储的自定义 XML。</span><span class="sxs-lookup"><span data-stu-id="bfd28-110">Custom XML stored in the document.</span></span>

- <span data-ttu-id="bfd28-111">使用基础浏览器控件提供的技术：浏览器 Cookie 或 HTML5 Web 存储（[localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)）。</span><span class="sxs-lookup"><span data-stu-id="bfd28-111">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>
    > [!NOTE]
    > <span data-ttu-id="bfd28-112">用户可以阻止基于浏览器的存储技术，具体取决于他们选择的设置。</span><span class="sxs-lookup"><span data-stu-id="bfd28-112">The user can block browser-based storage techniques depending on the settings they choose.</span></span>

<span data-ttu-id="bfd28-113">本文重点介绍如何使用 Office JavaScript API 将外接程序状态保留到当前文档。</span><span class="sxs-lookup"><span data-stu-id="bfd28-113">This article focuses on how to use the Office JavaScript API to persist add-in state to the current document.</span></span> <span data-ttu-id="bfd28-114">如果需要跨文档保留状态（例如，跨文档打开的任何文档跟踪用户首选项）需要使用不同的方法。</span><span class="sxs-lookup"><span data-stu-id="bfd28-114">If you need to persist state across documents, such as tracking user preferences across any documents they open, you will need to use a different approach.</span></span> <span data-ttu-id="bfd28-115">例如，您可以使用 [SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) 获取用户标识，然后将用户 ID 及其设置保存到联机数据库中。</span><span class="sxs-lookup"><span data-stu-id="bfd28-115">For example, you could use [SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) to obtain the user identity, and then save the user ID and their settings to an online database.</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a><span data-ttu-id="bfd28-116">使用 JavaScript API 保留Office状态和设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-116">Persisting add-in state and settings with the Office JavaScript API</span></span>

<span data-ttu-id="bfd28-117">JavaScript API Office提供了 设置、RoamingSettings[](/javascript/api/outlook/office.roamingsettings)和[CustomProperties](/javascript/api/outlook/office.customproperties)对象，用于跨会话保存外接程序状态，如下表所述。 [](/javascript/api/office/office.settings)</span><span class="sxs-lookup"><span data-stu-id="bfd28-117">The Office JavaScript API provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="bfd28-118">在所有情况下，保存的设置值仅与创建它们的外接程序 [Id](../reference/manifest/id.md) 相关联。</span><span class="sxs-lookup"><span data-stu-id="bfd28-118">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="bfd28-119">**对象**</span><span class="sxs-lookup"><span data-stu-id="bfd28-119">**Object**</span></span>|<span data-ttu-id="bfd28-120">**外接程序类型支持**</span><span class="sxs-lookup"><span data-stu-id="bfd28-120">**Add-in type support**</span></span>|<span data-ttu-id="bfd28-121">**存储位置**</span><span class="sxs-lookup"><span data-stu-id="bfd28-121">**Storage location**</span></span>|<span data-ttu-id="bfd28-122">**Office应用程序支持**</span><span class="sxs-lookup"><span data-stu-id="bfd28-122">**Office application support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="bfd28-123">Settings</span><span class="sxs-lookup"><span data-stu-id="bfd28-123">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="bfd28-124">内容和任务窗格</span><span class="sxs-lookup"><span data-stu-id="bfd28-124">content and task pane</span></span>|<span data-ttu-id="bfd28-125">加载项使用的文档、电子表格或演示文稿。</span><span class="sxs-lookup"><span data-stu-id="bfd28-125">The document, spreadsheet, or presentation the add-in is working with.</span></span> <span data-ttu-id="bfd28-126">内容和任务窗格加载项设置供创建它们的加载项使用，且能从保存它们的文档访问。</span><span class="sxs-lookup"><span data-stu-id="bfd28-126">Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="bfd28-p105">**重要说明：** 不要使用 **Settings** 对象保存密码和其他敏感的个人身份信息 (PII)。保存的数据对最终用户不可见，但它作为文档的一部分存储，可通过直接读取文档的文件格式进行访问。您应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。</span><span class="sxs-lookup"><span data-stu-id="bfd28-p105">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="bfd28-130">Word、Excel 或 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bfd28-130">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="bfd28-131">**注意：** Project 2013 任务窗格加载项不支持用于存储加载项状态或设置的 **Settings** API。</span><span class="sxs-lookup"><span data-stu-id="bfd28-131">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings.</span></span> <span data-ttu-id="bfd28-132">但是，对于在 Project (和其他 Office 客户端应用程序中运行的) 您可以使用浏览器 Cookie 或 Web 存储等技术。</span><span class="sxs-lookup"><span data-stu-id="bfd28-132">However, for add-ins running in Project (as well as other Office client applications) you can use techniques such as browser cookies or web storage.</span></span> <span data-ttu-id="bfd28-133">有关详细信息，请参阅 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。</span><span class="sxs-lookup"><span data-stu-id="bfd28-133">For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="bfd28-134">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="bfd28-134">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="bfd28-135">Outlook</span><span class="sxs-lookup"><span data-stu-id="bfd28-135">Outlook</span></span>|<span data-ttu-id="bfd28-136">安装了加载项的用户 Exchange 服务器邮箱。</span><span class="sxs-lookup"><span data-stu-id="bfd28-136">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="bfd28-137">由于这些设置存储在用户的服务器邮箱中，因此它们可以随用户"漫游"，并且可在外接程序在任何访问该用户邮箱的受支持 Office 客户端应用程序或浏览器的上下文中运行时使用。</span><span class="sxs-lookup"><span data-stu-id="bfd28-137">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported Office client application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="bfd28-138">Outlook 加载项漫游设置只可供创建它们的加载项使用，且只能从安装了加载项的邮箱访问。</span><span class="sxs-lookup"><span data-stu-id="bfd28-138">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="bfd28-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="bfd28-139">Outlook</span></span>|
|[<span data-ttu-id="bfd28-140">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="bfd28-140">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="bfd28-141">Outlook</span><span class="sxs-lookup"><span data-stu-id="bfd28-141">Outlook</span></span>|<span data-ttu-id="bfd28-p108">加载项使用的邮件、约会或会议请求项目。 Outlook 外接程序项目自定义属性仅供创建它们的外接程序使用，并且只能从保存它们的项目使用。</span><span class="sxs-lookup"><span data-stu-id="bfd28-p108">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="bfd28-144">Outlook</span><span class="sxs-lookup"><span data-stu-id="bfd28-144">Outlook</span></span>|
|[<span data-ttu-id="bfd28-145">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfd28-145">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="bfd28-146">任务窗格</span><span class="sxs-lookup"><span data-stu-id="bfd28-146">task pane</span></span>|<span data-ttu-id="bfd28-p109">加载项要使用的文档、电子表格或演示文稿。任务窗格加载项设置可供创建它们的加载项使用，且能从保存它们的文档访问。</span><span class="sxs-lookup"><span data-stu-id="bfd28-p109">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="bfd28-p110">**重要说明：** 请勿将密码和其他敏感的个人身份信息 (PII) 存储在自定义 XML 部分中。虽然保存的数据对最终用户不可见，但它存储为文档的一部分，可通过直接读取文档的文件格式进行访问。应限制加载项对 PII 的使用，并仅将加载项所需的任何 PII 存储在服务器上，且服务器将加载项托管为用户保护资源。</span><span class="sxs-lookup"><span data-stu-id="bfd28-p110">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="bfd28-152">Word (JavaScript OFFICE JavaScript API) Excel (JavaScript API Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="bfd28-152">Word (using the Office JavaScript Common API) Excel (using the application-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="bfd28-153">设置数据在运行时托管在内存中</span><span class="sxs-lookup"><span data-stu-id="bfd28-153">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="bfd28-154">下面两部分是在 Office 常见 JavaScript API 上下文中介绍的设置。</span><span class="sxs-lookup"><span data-stu-id="bfd28-154">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="bfd28-155">特定于应用程序的应用程序Excel JavaScript API 还提供对自定义设置的访问权限。</span><span class="sxs-lookup"><span data-stu-id="bfd28-155">The application-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="bfd28-156">Excel API 和编程模式有点不一样。</span><span class="sxs-lookup"><span data-stu-id="bfd28-156">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="bfd28-157">有关详细信息，请参阅 [Excel SettingCollection](/javascript/api/excel/excel.settingcollection)。</span><span class="sxs-lookup"><span data-stu-id="bfd28-157">For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="bfd28-158">在内部，使用 、 或 对象访问的属性包中数据存储为序列化的 JavaScript 对象表示法 `Settings` `CustomProperties` (JSON) 对象，其中包含名称/值对 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="bfd28-158">Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="bfd28-159">每个 (键) 的名称必须为 ，并且存储的值可以是 JavaScript 、、 或 ， `string` `string` `number` `date` `object` 但不能是 **函数**。</span><span class="sxs-lookup"><span data-stu-id="bfd28-159">The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.</span></span>

<span data-ttu-id="bfd28-160">本属性包结构示例包含三个已定义 **string** 值，分别为 `firstName`、 `location` 和 `defaultView`。</span><span class="sxs-lookup"><span data-stu-id="bfd28-160">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="bfd28-161">在前一个加载项会话中保存设置属性包之后，可以在加载项的当前会话中初始化加载项时或在之后的任何时间加载该设置属性包。</span><span class="sxs-lookup"><span data-stu-id="bfd28-161">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="bfd28-162">在会话期间，使用 对象的 、 和 方法完全在内存中管理设置，这些对象对应于要创建 `get` `set` (`remove` 设置、CustomProperties 或 **RoamingSettings** ) 的设置类型。</span><span class="sxs-lookup"><span data-stu-id="bfd28-162">During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you are creating (**Settings**, **CustomProperties**, or **RoamingSettings**).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bfd28-163">若要将加载项当前会话期间执行的任何添加、更新或删除操作保留到存储位置，必须调用用于处理此类设置的相应对象的 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="bfd28-163">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="bfd28-164">`get`、 `set` 和 `remove` 方法仅对设置属性包的内存副本进行操作。</span><span class="sxs-lookup"><span data-stu-id="bfd28-164">The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="bfd28-165">如果加载项在未调用的情况下关闭，则在此会话期间对设置进行 `saveAsync` 的任何更改都将丢失。</span><span class="sxs-lookup"><span data-stu-id="bfd28-165">If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.</span></span>

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="bfd28-166">如何按文档暂留内容和任务窗格加载项的加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-166">How to save add-in state and settings per document for content and task pane add-ins</span></span>

<span data-ttu-id="bfd28-167">要保留 Word、Excel 或 PowerPoint 的内容或任务窗格加载项的状态或自定义设置，可使用 [Settings](/javascript/api/office/office.settings) 对象及其方法。</span><span class="sxs-lookup"><span data-stu-id="bfd28-167">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods.</span></span> <span data-ttu-id="bfd28-168">使用 对象的方法创建的属性包仅可用于创建对象的内容或任务窗格加载项的实例，并且只能从保存它 `Settings` 的文档使用。</span><span class="sxs-lookup"><span data-stu-id="bfd28-168">The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="bfd28-169">该对象将自动作为 Document 对象的一部分加载，并且可在任务窗格或内容 `Settings` 外接程序激活时使用。 [](/javascript/api/office/office.document)</span><span class="sxs-lookup"><span data-stu-id="bfd28-169">The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="bfd28-170">实例 `Document` 化对象后，可以使用 `Settings` 对象的 [settings](/javascript/api/office/office.document#settings) 属性访问 `Document` 该对象。</span><span class="sxs-lookup"><span data-stu-id="bfd28-170">After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#settings) property of the `Document` object.</span></span> <span data-ttu-id="bfd28-171">在会话的生命周期中，只能使用 、 和 方法读取、写入或删除属性包的内存副本中的保留设置和 `Settings.get` `Settings.set` `Settings.remove` 加载项状态。</span><span class="sxs-lookup"><span data-stu-id="bfd28-171">During the lifetime of the session, you can just use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="bfd28-172">由于 set 和 remove 方法仅针对设置属性包的内存副本，若要将新的或更改的设置保存回加载项关联的文档，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法。</span><span class="sxs-lookup"><span data-stu-id="bfd28-172">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>

### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="bfd28-173">创建或更新设置值</span><span class="sxs-lookup"><span data-stu-id="bfd28-173">Creating or updating a setting value</span></span>

<span data-ttu-id="bfd28-p117">以下代码示例演示如何使用 [Settings.set](/javascript/api/office/office.settings#set-name--value-) 方法创建名为 `'themeColor'` 且值为 `'green'` 的设置。set 方法的第一个参数是要设置或创建的设置的 _name_ (Id)（区分大小写）。第二个参数是设置的 _value_。</span><span class="sxs-lookup"><span data-stu-id="bfd28-p117">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>

```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="bfd28-177">如果具有指定名称的设置尚不存在，则创建此设置，如果此设置存在，则对值进行更新。</span><span class="sxs-lookup"><span data-stu-id="bfd28-177">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist.</span></span> <span data-ttu-id="bfd28-178">使用 `Settings.saveAsync` 方法将新的或更新的设置保留到文档中。</span><span class="sxs-lookup"><span data-stu-id="bfd28-178">Use the `Settings.saveAsync` method to persist the new or updated settings to the document.</span></span>

### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="bfd28-179">获取设置的值</span><span class="sxs-lookup"><span data-stu-id="bfd28-179">Getting the value of a setting</span></span>

<span data-ttu-id="bfd28-180">下面的示例演示如何使用 [Settings.get](/javascript/api/office/office.settings#get-name-) 方法获取名为"themeColor"的设置值。</span><span class="sxs-lookup"><span data-stu-id="bfd28-180">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor".</span></span> <span data-ttu-id="bfd28-181">该方法的唯一 `get` 参数是设置 _的名称（区分_ 大小写）。</span><span class="sxs-lookup"><span data-stu-id="bfd28-181">The only parameter of the `get` method is the case-sensitive _name_ of the setting.</span></span>

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 <span data-ttu-id="bfd28-182">`get`方法返回之前为传入的设置 _名称_ 保存的值。</span><span class="sxs-lookup"><span data-stu-id="bfd28-182">The `get` method returns the value that was previously saved for the setting _name_ that was passed in.</span></span> <span data-ttu-id="bfd28-183">如果不存在该设置，那么方法返回 **null**。</span><span class="sxs-lookup"><span data-stu-id="bfd28-183">If the setting doesn't exist, the method returns **null**.</span></span>

### <a name="removing-a-setting"></a><span data-ttu-id="bfd28-184">删除设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-184">Removing a setting</span></span>

<span data-ttu-id="bfd28-185">下面的示例演示如何使用 [Settings.remove](/javascript/api/office/office.settings#remove-name-) 方法删除名为"themeColor"的设置。</span><span class="sxs-lookup"><span data-stu-id="bfd28-185">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor".</span></span> <span data-ttu-id="bfd28-186">该方法的唯一 `remove` 参数是设置 _的名称（区分_ 大小写）。</span><span class="sxs-lookup"><span data-stu-id="bfd28-186">The only parameter of the `remove` method is the case-sensitive _name_ of the setting.</span></span>

```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="bfd28-187">如果不存在该设置，则不执行任何操作。</span><span class="sxs-lookup"><span data-stu-id="bfd28-187">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="bfd28-188">使用 `Settings.saveAsync` 方法可保留文档中设置的删除操作。</span><span class="sxs-lookup"><span data-stu-id="bfd28-188">Use the `Settings.saveAsync` method to persist removal of the setting from the document.</span></span>

### <a name="saving-your-settings"></a><span data-ttu-id="bfd28-189">保存设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-189">Saving your settings</span></span>

<span data-ttu-id="bfd28-190">若要保存当前会话中加载项对设置属性包内存副本所做的任意添加、更改或删除操作，必须调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法将它们存储在文档中。</span><span class="sxs-lookup"><span data-stu-id="bfd28-190">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document.</span></span> <span data-ttu-id="bfd28-191">方法的唯一 `saveAsync` 参数是 _callback_，它是具有单个参数的回调函数。</span><span class="sxs-lookup"><span data-stu-id="bfd28-191">The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.</span></span>

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

<span data-ttu-id="bfd28-192">作为 callback 参数传入 `saveAsync` 方法的匿名函数在操作完成时执行。</span><span class="sxs-lookup"><span data-stu-id="bfd28-192">The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="bfd28-193">回调 _的 asyncResult_ 参数提供对包含操作 `AsyncResult` 状态的对象的访问。</span><span class="sxs-lookup"><span data-stu-id="bfd28-193">The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation.</span></span> <span data-ttu-id="bfd28-194">在示例中，函数检查 属性以查看保存操作是成功还是失败，然后在加载项页面中 `AsyncResult.status` 显示结果。</span><span class="sxs-lookup"><span data-stu-id="bfd28-194">In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="bfd28-195">如何将自定义 XML 保存到文档</span><span class="sxs-lookup"><span data-stu-id="bfd28-195">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="bfd28-196">此部分是在 Word 中支持的 Office 常见 JavaScript API 上下文中介绍的自定义 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="bfd28-196">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="bfd28-197">特定于应用程序的应用程序Excel JavaScript API 还提供对自定义 XML 部件的访问权限。</span><span class="sxs-lookup"><span data-stu-id="bfd28-197">The application-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="bfd28-198">Excel API 和编程模式有点不一样。</span><span class="sxs-lookup"><span data-stu-id="bfd28-198">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="bfd28-199">有关详细信息，请参阅 [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart)。</span><span class="sxs-lookup"><span data-stu-id="bfd28-199">For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="bfd28-200">当需要存储的信息超过文档文档大小限制或具有结构化字符设置存在附加存储选项。</span><span class="sxs-lookup"><span data-stu-id="bfd28-200">There is an additional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="bfd28-201">可以在 Word 的任务窗格加载项中暂留自定义 XML 标记（对于 Excel，但请参阅本节顶部的注释）。</span><span class="sxs-lookup"><span data-stu-id="bfd28-201">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="bfd28-202">在 Word 中，可以使用 [CustomXmlPart](/javascript/api/office/office.customxmlpart) 对象及其方法（同样，请参阅上面的 Excel 注释）。</span><span class="sxs-lookup"><span data-stu-id="bfd28-202">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="bfd28-203">以下代码将创建自定义 XML 部件，并在页面的 divs 中显示其 ID 及内容。</span><span class="sxs-lookup"><span data-stu-id="bfd28-203">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="bfd28-204">请注意，XML 字符串中必须有一个 `xmlns` 属性。</span><span class="sxs-lookup"><span data-stu-id="bfd28-204">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="bfd28-205">若要检索自定义 XML 部分，请使用 [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) 方法，但 ID 是在创建 XML 部分时生成的 GUID，因此编码时无法知道 ID 是什么。</span><span class="sxs-lookup"><span data-stu-id="bfd28-205">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is.</span></span> <span data-ttu-id="bfd28-206">因此，最好是在创建 XML 部分时，立即将 XML 部分的 ID 存储为设置，并为它提供容易记住的密钥。</span><span class="sxs-lookup"><span data-stu-id="bfd28-206">For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key.</span></span> <span data-ttu-id="bfd28-207">下面的方法展示了如何执行此操作。</span><span class="sxs-lookup"><span data-stu-id="bfd28-207">The following method shows how to do this.</span></span> <span data-ttu-id="bfd28-208"> (请参阅本文的前面部分，了解使用自定义设置时的详细信息和) </span><span class="sxs-lookup"><span data-stu-id="bfd28-208">(But see earlier sections of this article for details and best practices when working with custom settings.)</span></span>

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

<span data-ttu-id="bfd28-209">下面的代码展示了如何通过先从设置中获取 ID 来检索 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="bfd28-209">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a><span data-ttu-id="bfd28-210">如何在加载项中Outlook设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-210">How to save settings in an Outlook add-in</span></span>

<span data-ttu-id="bfd28-211">若要了解如何在加载项中保存Outlook，请参阅管理加载项的状态[Outlook设置](../outlook/manage-state-and-settings-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="bfd28-211">For information about how to save settings in an Outlook add-in, see [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="bfd28-212">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bfd28-212">See also</span></span>

- [<span data-ttu-id="bfd28-213">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="bfd28-213">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="bfd28-214">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="bfd28-214">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="bfd28-215">管理加载项的状态Outlook设置</span><span class="sxs-lookup"><span data-stu-id="bfd28-215">Manage state and settings for an Outlook add-in</span></span>](../outlook/manage-state-and-settings-outlook.md)
- [<span data-ttu-id="bfd28-216">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="bfd28-216">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
