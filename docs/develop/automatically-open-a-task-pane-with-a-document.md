---
title: 随文档自动打开任务窗格
description: ''
ms.date: 05/02/2018
ms.openlocfilehash: 7022381eb24d5b0566936b068f2940703f0eef9f
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533747"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a><span data-ttu-id="0e370-102">随文档自动打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="0e370-102">Automatically open a task pane with a document</span></span>

<span data-ttu-id="0e370-p101">可以使用 Office 外接程序中的外接程序命令，通过将按钮添加到 Office 功能区来扩展 Office UI。当用户单击命令按钮时，会执行一个操作，如打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0e370-p101">You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office ribbon. When users click your command button, an action occurs, such as opening a task pane.</span></span>

<span data-ttu-id="0e370-105">某些情况下，需要在文档打开时自动打开一个任务窗格，而无需进行显式用户交互。</span><span class="sxs-lookup"><span data-stu-id="0e370-105">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction.</span></span> <span data-ttu-id="0e370-106">可以使用 Addincommand 1.1 要求集中引入的 AutoOpen 任务窗格功能，以在情况需要时自动打开一个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0e370-106">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction. You can use the autoopen taskpane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it.</span></span>


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a><span data-ttu-id="0e370-107">AutoOpen 功能与插入任务窗格有何不同？</span><span class="sxs-lookup"><span data-stu-id="0e370-107">How is the autoopen feature different from inserting a task pane?</span></span>

<span data-ttu-id="0e370-p103">如果用户启动不使用外接程序命令的外接程序（例如，在 Office 2013 中运行的外接程序），外接程序会插入并保留在文档中。因此，当其他用户打开文档时，系统会提示他们安装外接程序，随后会打开任务窗格。这种模型面临的挑战是，在许多情况下，用户不希望外接程序保留在文档中。例如，在 Word 文档中使用字典外接的学生可能不希望系统他们的同学或老师在打开该文档时提示他们安装该外接程序。</span><span class="sxs-lookup"><span data-stu-id="0e370-p103">When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document. As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens. The challenge with this model is that in many cases, users don’t want the add-in to persist in the document. For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.</span></span>

<span data-ttu-id="0e370-112">使用 Autoopen 功能，可以显式定义或允许用户定义特定任务窗格外接程序是否保留在特定文档中。</span><span class="sxs-lookup"><span data-stu-id="0e370-112">With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document.</span></span>

## <a name="support-and-availability"></a><span data-ttu-id="0e370-113">支持和可用性</span><span class="sxs-lookup"><span data-stu-id="0e370-113">Support and availability</span></span>

<span data-ttu-id="0e370-114">Autoopen 功能目前<!-- in **developer preview** and it is only -->在以下产品和平台中受支持。</span><span class="sxs-lookup"><span data-stu-id="0e370-114">The autoopen feature is currently <!-- in **developer preview** and it is only --> supported in the following products and platforms.</span></span>

|<span data-ttu-id="0e370-115">**产品**</span><span class="sxs-lookup"><span data-stu-id="0e370-115">**Products**</span></span>|<span data-ttu-id="0e370-116">**平台**</span><span class="sxs-lookup"><span data-stu-id="0e370-116">**Platforms**</span></span>|
|:-----------|:------------|
|<ul><li><span data-ttu-id="0e370-117">Word</span><span class="sxs-lookup"><span data-stu-id="0e370-117">Word</span></span></li><li><span data-ttu-id="0e370-118">Excel</span><span class="sxs-lookup"><span data-stu-id="0e370-118">Excel</span></span></li><li><span data-ttu-id="0e370-119">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0e370-119">PowerPoint</span></span></li></ul>|<span data-ttu-id="0e370-120">所有产品的支持平台：</span><span class="sxs-lookup"><span data-stu-id="0e370-120">Supported platforms for all products:</span></span><ul><li><span data-ttu-id="0e370-p104">Office for Windows Desktop（内部版本 16.0.8121.1000+）</span><span class="sxs-lookup"><span data-stu-id="0e370-p104">Office for Windows Desktop. Build 16.0.8121.1000+</span></span></li><li><span data-ttu-id="0e370-p105">Office for Mac（生成号 15.34.17051500+）</span><span class="sxs-lookup"><span data-stu-id="0e370-p105">Office for Mac. Build 15.34.17051500+</span></span></li><li><span data-ttu-id="0e370-125">Office Online</span><span class="sxs-lookup"><span data-stu-id="0e370-125"> (Office Online)</span></span></li></ul>|


## <a name="best-practices"></a><span data-ttu-id="0e370-126">最佳做法</span><span class="sxs-lookup"><span data-stu-id="0e370-126">Best practices</span></span>

<span data-ttu-id="0e370-127">在使用 Autoopen 功能时应用下面的最佳做法：</span><span class="sxs-lookup"><span data-stu-id="0e370-127">Apply the following best practices when you use the autoopen feature:</span></span>

- <span data-ttu-id="0e370-128">当 Autoopen 功能可帮助外接程序用户工作更高效时使用此功能，如：</span><span class="sxs-lookup"><span data-stu-id="0e370-128">Use the autoopen feature when it will help make your add-in users more efficient, such as:</span></span>
  - <span data-ttu-id="0e370-p106">当文档需要外接程序才能正常工作时。例如，包括由外接程序定期刷新的股票值的电子表格。外接程序应在电子表格打开时自动打开，以保持值处于最新状态。</span><span class="sxs-lookup"><span data-stu-id="0e370-p106">When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span></span>
  - <span data-ttu-id="0e370-p107">当用户很可能始终将外接程序与某个特定文档一同使用时。例如，可帮助用户通过从后台系统中获取信息来填写或更改文档中数据的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0e370-p107">When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span></span>
- <span data-ttu-id="0e370-p108">允许用户打开或关闭 Autoopen 功能。用户可以选择 UI 中包含的一个选项来停止自动打开外接程序任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0e370-p108">Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span></span>  
- <span data-ttu-id="0e370-136">使用要求集检测来确定 Autoopen 功能是否可用，如果不可用则提供回退行为。</span><span class="sxs-lookup"><span data-stu-id="0e370-136">Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn’t.</span></span>
- <span data-ttu-id="0e370-p109">不要使用 Autoopen 功能来人为地增加外接程序的使用率。如果外接程序在某些文档中自动打开没有任何意义，那么这个功能就会令用户生厌。</span><span class="sxs-lookup"><span data-stu-id="0e370-p109">Don't use the autoopen feature to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to open automatically with certain documents, this feature can annoy users.</span></span>

    > [!NOTE]
    > <span data-ttu-id="0e370-139">如果 Microsoft 检测到滥用 AutoOpen 功能，加载项可能会从 AppSource 下架。</span><span class="sxs-lookup"><span data-stu-id="0e370-139">If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource.</span></span>

- <span data-ttu-id="0e370-p110">请勿使用此功能来固定多个任务窗格。只能设置一个外接程序窗格随文档自动打开。</span><span class="sxs-lookup"><span data-stu-id="0e370-p110">Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.</span></span>  

## <a name="implementation"></a><span data-ttu-id="0e370-142">实现</span><span class="sxs-lookup"><span data-stu-id="0e370-142">Implementation</span></span>

<span data-ttu-id="0e370-143">要实现 Autoopen 功能，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="0e370-143">To implement the autoopen feature:</span></span>

- <span data-ttu-id="0e370-144">指定要自动打开的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0e370-144">Specify the task pane to be opened automatically.</span></span>
- <span data-ttu-id="0e370-145">标记要自动打开任务窗格的文档。</span><span class="sxs-lookup"><span data-stu-id="0e370-145">Tag the document to automatically open the task pane.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0e370-p111">只有在用户设备上已安装加载项时，才能打开指定为自动打开的窗格。如果在打开文档时用户未安装加载项，那么 AutoOpen 功能将不起作用，而且设置也会被忽略。如果还要求加载项与文档一起分发，需要将“visibility”属性设置为 1；只能使用 OpenXML 完成此操作，本文稍后将提供示例。</span><span class="sxs-lookup"><span data-stu-id="0e370-p111">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span></span>

### <a name="step-1-specify-the-task-pane-to-open"></a><span data-ttu-id="0e370-149">第 1 步：指定要打开的任务窗格</span><span class="sxs-lookup"><span data-stu-id="0e370-149">Step 1: Specify the task pane to open</span></span>

<span data-ttu-id="0e370-p112">若要指定要自动打开的任务窗格，请将 [TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) 值设置为 **Office.AutoShowTaskpaneWithDocument**。只能在一个任务窗格上设置此值。如果在多个任务窗格上设置此值，将识别值的第一个匹配项，而忽略其他。</span><span class="sxs-lookup"><span data-stu-id="0e370-p112">To specify the task pane to open automatically, set the [TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span></span>

<span data-ttu-id="0e370-153">在下面的示例中，TaskPaneId 值设置为 Office.AutoShowTaskpaneWithDocument。</span><span class="sxs-lookup"><span data-stu-id="0e370-153">The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a><span data-ttu-id="0e370-154">第 2 步：将文档标记为自动打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="0e370-154">Step 2: Tag the document to automatically open the task pane</span></span>

<span data-ttu-id="0e370-155">可以通过下面的两种方法之一，将文档标记为触发自动打开功能。</span><span class="sxs-lookup"><span data-stu-id="0e370-155">You can tag the document to trigger the autoopen feature in one of two ways.</span></span> <span data-ttu-id="0e370-156">选择最适合自己应用场景的备选方法。</span><span class="sxs-lookup"><span data-stu-id="0e370-156">Pick the alternative that works best for your scenario.</span></span>  


#### <a name="tag-the-document-on-the-client-side"></a><span data-ttu-id="0e370-157">在客户端上标记文档</span><span class="sxs-lookup"><span data-stu-id="0e370-157">Tag the document on the client side</span></span>

<span data-ttu-id="0e370-158">使用 Office.js [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) 方法将 **Office.AutoShowTaskpaneWithDocument** 设置为“**true**”，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="0e370-158">Use the Office.js [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.</span></span>

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

<span data-ttu-id="0e370-159">如果需要将文档标记为外接程序交互的一部分（例如，在用户创建一个绑定，或选择一个选项来表示他们希望窗格自动打开时），则使用此方法。</span><span class="sxs-lookup"><span data-stu-id="0e370-159">Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).</span></span>

#### <a name="use-open-xml-to-tag-the-document"></a><span data-ttu-id="0e370-160">使用 Open XML 标记文档</span><span class="sxs-lookup"><span data-stu-id="0e370-160">Use Open XML to tag the document</span></span>

<span data-ttu-id="0e370-p114">可以使用 Open XML 来创建或修改文档，并添加适当的 Open Office XML 标记来触发 Autoopen 功能。有关演示如何执行此操作的示例，请参阅 [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)。</span><span class="sxs-lookup"><span data-stu-id="0e370-p114">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span></span>

<span data-ttu-id="0e370-163">向文档添加两个 Open XML 部件：</span><span class="sxs-lookup"><span data-stu-id="0e370-163">Add two Open XML parts to the document:</span></span>

- <span data-ttu-id="0e370-164">一个 `webextension` 部件</span><span class="sxs-lookup"><span data-stu-id="0e370-164">A `webextension` part</span></span>
- <span data-ttu-id="0e370-165">一个 `taskpane` 部件</span><span class="sxs-lookup"><span data-stu-id="0e370-165">A `taskpane` part</span></span>

<span data-ttu-id="0e370-166">以下示例演示如何添加 `webextension` 部件。</span><span class="sxs-lookup"><span data-stu-id="0e370-166">The following example shows how to add the taskpane part.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
   <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="0e370-167">`webextension` 部件包含一个属性包，以及必须设置为 `true` 的 **Office.AutoShowTaskpaneWithDocument** 属性。</span><span class="sxs-lookup"><span data-stu-id="0e370-167">The webextension part includes a property bag and a property named Office.AutoShowTaskpaneWithDocument that must be set to .</span></span>

<span data-ttu-id="0e370-168">`webextension` 部件还包括对具有 `id`、`storeType`、`store` 和 `version` 的属性的应用商店或目录的引用。</span><span class="sxs-lookup"><span data-stu-id="0e370-168">The `webextension` part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`.</span></span> <span data-ttu-id="0e370-169">在 `storeType` 值中，只有四个与 AutoOpen 功能相关。</span><span class="sxs-lookup"><span data-stu-id="0e370-169">Of the `storeType` values, only four are relevant to the autoopen feature.</span></span> <span data-ttu-id="0e370-170">其他三个属性的值取决于 `storeType` 的值，如下表所示。</span><span class="sxs-lookup"><span data-stu-id="0e370-170">The values for the other three attributes depend on the value for `storeType`, as shown in the following table.</span></span>

| <span data-ttu-id="0e370-171">**`storeType` 值**</span><span class="sxs-lookup"><span data-stu-id="0e370-171">**`storeType` value**</span></span> | <span data-ttu-id="0e370-172">**`id` 值**</span><span class="sxs-lookup"><span data-stu-id="0e370-172">**`id` value**</span></span>    |<span data-ttu-id="0e370-173">**`store` 值**</span><span class="sxs-lookup"><span data-stu-id="0e370-173">**`store` value**</span></span> | <span data-ttu-id="0e370-174">**`version` 值**</span><span class="sxs-lookup"><span data-stu-id="0e370-174">**`version` value**</span></span>|
|:---------------|:---------------|:---------------|:---------------|
|<span data-ttu-id="0e370-175">OMEX (AppSource)</span><span class="sxs-lookup"><span data-stu-id="0e370-175">OMEX (AppSource)</span></span>|<span data-ttu-id="0e370-176">加载项的 AppSource 资产 ID（请参阅“注意”）</span><span class="sxs-lookup"><span data-stu-id="0e370-176">The AppSource asset ID of the add-in (see Note)</span></span>|<span data-ttu-id="0e370-177">AppSource 的区域设置；例如，“en-us”。</span><span class="sxs-lookup"><span data-stu-id="0e370-177">The locale of AppSource; for example, "en-us".</span></span>|<span data-ttu-id="0e370-178">AppSource 目录中的版本（请参阅“注意”）</span><span class="sxs-lookup"><span data-stu-id="0e370-178">The version in the AppSource catalog (see Note)</span></span>|
|<span data-ttu-id="0e370-179">FileSystem（网络共享）</span><span class="sxs-lookup"><span data-stu-id="0e370-179">FileSystem (a network share)</span></span>|<span data-ttu-id="0e370-180">外接程序清单中外接程序的 GUID。</span><span class="sxs-lookup"><span data-stu-id="0e370-180">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="0e370-181">网络共享路径。例如，“\\\\MyComputer\\MySharedFolder”。</span><span class="sxs-lookup"><span data-stu-id="0e370-181">The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span>|<span data-ttu-id="0e370-182">外接程序清单中的版本。</span><span class="sxs-lookup"><span data-stu-id="0e370-182">The version in the add-in manifest.</span></span>|
|<span data-ttu-id="0e370-183">EXCatalog（通过 Exchange 服务器部署）</span><span class="sxs-lookup"><span data-stu-id="0e370-183">EXCatalog (deployment via the Exchange server)</span></span> |<span data-ttu-id="0e370-184">外接程序清单中外接程序的 GUID。</span><span class="sxs-lookup"><span data-stu-id="0e370-184">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="0e370-185">“EXCatalog”。</span><span class="sxs-lookup"><span data-stu-id="0e370-185">"EXCatalog"</span></span> <span data-ttu-id="0e370-186">EXCatalog 行是与在 Office 365 管理中心中使用集中式部署的外接程序一起使用的行。</span><span class="sxs-lookup"><span data-stu-id="0e370-186">EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Office 365 admin center.</span></span>|<span data-ttu-id="0e370-187">外接程序清单中的版本。</span><span class="sxs-lookup"><span data-stu-id="0e370-187">The version in the add-in manifest.</span></span>
|<span data-ttu-id="0e370-188">Registry（系统注册表）</span><span class="sxs-lookup"><span data-stu-id="0e370-188">Registry (System registry)</span></span>|<span data-ttu-id="0e370-189">外接程序清单中外接程序的 GUID。</span><span class="sxs-lookup"><span data-stu-id="0e370-189">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="0e370-190">“developer”</span><span class="sxs-lookup"><span data-stu-id="0e370-190">"developer"</span></span>|<span data-ttu-id="0e370-191">加载项清单中的版本。</span><span class="sxs-lookup"><span data-stu-id="0e370-191">The version in the add-in manifest.</span></span>|

> [!NOTE]
> <span data-ttu-id="0e370-p117">若要查找 AppSource 中加载项的资产 ID 和版本，请转到加载项的 AppSource 登陆页面。资产 ID 显示在浏览器的地址栏中。版本在页面的“详细信息”\*\*\*\* 部分中列出。</span><span class="sxs-lookup"><span data-stu-id="0e370-p117">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.</span></span>

<span data-ttu-id="0e370-195">若要详细了解 webextension 标记，请参阅 [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)。</span><span class="sxs-lookup"><span data-stu-id="0e370-195">For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span></span>

<span data-ttu-id="0e370-196">以下示例演示如何添加 `taskpane` 部件。</span><span class="sxs-lookup"><span data-stu-id="0e370-196">The following example shows how to add the taskpane part.</span></span>

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

<span data-ttu-id="0e370-197">请注意，在本例中，`visibility` 属性设置为“0”。</span><span class="sxs-lookup"><span data-stu-id="0e370-197">Note that in this example, the `visibility` attribute is set to "0".</span></span> <span data-ttu-id="0e370-198">这意味着在添加 webextension 部件和 `taskpane` 部件之后，第一次打开文档时，用户还必须从功能区上的“外接程序”\*\*\*\* 按钮安装该外接程序。</span><span class="sxs-lookup"><span data-stu-id="0e370-198">This means that after the webextension and `taskpane` parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon.</span></span> <span data-ttu-id="0e370-199">此后，外接程序任务窗格将在打开该文件时自动打开。</span><span class="sxs-lookup"><span data-stu-id="0e370-199">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="0e370-200">此外，在将 `visibility` 设置为“0”时，可以使用 Office.js 让用户打开或关闭 AutoOpen 功能。</span><span class="sxs-lookup"><span data-stu-id="0e370-200">Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="0e370-201">具体来说，脚本会将 **Office.AutoShowTaskpaneWithDocument** 文档设置为 `true` 或 `false`。</span><span class="sxs-lookup"><span data-stu-id="0e370-201">Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`.</span></span> <span data-ttu-id="0e370-202">（有关详细信息，请参阅[在客户端上标记文档](#tag-the-document-on-the-client-side)。）</span><span class="sxs-lookup"><span data-stu-id="0e370-202">(For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).)</span></span>

<span data-ttu-id="0e370-p119">如果 `visibility` 设置为“1”，任务窗格将在文件第一次打开时自动打开。系统会提示用户信任该外接程序，授予信任后，将打开外接程序。此后，外接程序任务窗格将在打开该文件时自动打开。但是，当 `visibility` 设置为“1”时，则不能使用 Office.js 让用户打开或关闭 Autoopen 功能。</span><span class="sxs-lookup"><span data-stu-id="0e370-p119">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span></span>

<span data-ttu-id="0e370-207">当外接程序和模板或文档内容紧密集成以致用户不会选择退出 Autoopen 功能时，将 `visibility` 设置为“1”是一个不错的选择。</span><span class="sxs-lookup"><span data-stu-id="0e370-207">Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature.</span></span>

> [!NOTE]
> <span data-ttu-id="0e370-p120">若要将加载项与文档一起分发，以便提示用户进行安装，必须将“visibility”属性设置为 1。只能通过 Open XML 执行此操作。</span><span class="sxs-lookup"><span data-stu-id="0e370-p120">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.</span></span>

<span data-ttu-id="0e370-p121">编写 XML 的一个简单方法是首先运行外接程序并[标记客户端上的文档](#tag-the-document-on-the-client-side)以写入值，然后保存该文档并检查生成的 XML。Office 将检测并提供适当的属性值。还可以使用 [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) 工具生成 C# 代码以编程方式添加基于生成的 XML 的标记。</span><span class="sxs-lookup"><span data-stu-id="0e370-p121">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span></span>

## <a name="test-and-verify-opening-task-panes"></a><span data-ttu-id="0e370-213">对打开任务窗格进行测试和验证</span><span class="sxs-lookup"><span data-stu-id="0e370-213">Test and verify opening task panes</span></span>

<span data-ttu-id="0e370-214">可以通过 Office 365 管理中心使用集中部署，来部署将自动打开任务窗格的测试版外接程序。</span><span class="sxs-lookup"><span data-stu-id="0e370-214">You can deploy a test version of your add-in that will automatically open a task pane using Centralized Deployment via the Office 365 admin center.</span></span> <span data-ttu-id="0e370-215">以下示例演示如何使用 EXCatalog 应用商店版本从集中部署目录插入外接程序。</span><span class="sxs-lookup"><span data-stu-id="0e370-215">The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="0e370-216">若要测试以前示例，请考虑加入 [Office 365 开发人员计划](https://docs.microsoft.com/office/developer-program/office-365-developer-program)，并注册以获取一个 [Office 365 开发人员帐户](https://developer.microsoft.com/office/dev-program)（如果还没有 Office 365 订阅）。</span><span class="sxs-lookup"><span data-stu-id="0e370-216">To test the previous example, please consider joining the [Office 365 Developer Program](https://docs.microsoft.com/office/developer-program/office-365-developer-program) and signing up for an [Office 365 developer account](https://developer.microsoft.com/office/dev-program) if you don't already own an Office 365 subscription.</span></span> <span data-ttu-id="0e370-217">可以实际测试驱动集中部署，并验证外接程序是否按预期工作。</span><span class="sxs-lookup"><span data-stu-id="0e370-217">You can actually test drive Centralized Deployment and verify that your add-in works as expected.</span></span>


## <a name="see-also"></a><span data-ttu-id="0e370-218">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0e370-218">See also</span></span>

<span data-ttu-id="0e370-219">有关演示如何使用 AutoOpen 功能的示例，请参阅 [Office 外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)。</span><span class="sxs-lookup"><span data-stu-id="0e370-219">For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).</span></span>
<span data-ttu-id="0e370-220">[加入 Office 365 开发人员计划](https://docs.microsoft.com/office/developer-program/office-365-developer-program)。</span><span class="sxs-lookup"><span data-stu-id="0e370-220">[Join the Office 365 Developer Program](https://docs.microsoft.com/office/developer-program/office-365-developer-program)</span></span>
