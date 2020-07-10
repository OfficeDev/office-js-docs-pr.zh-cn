---
title: 随文档自动打开任务窗格
description: 了解如何将 Office 外接程序配置为在文档打开时自动打开。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 85b421a569ccb83c3d07f0f10fd4767929332f96
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093705"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a><span data-ttu-id="45254-103">随文档自动打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="45254-103">Automatically open a task pane with a document</span></span>

<span data-ttu-id="45254-104">您可以使用 Office 外接程序中的外接程序命令，通过向 Office 应用程序功能区添加按钮来扩展 Office UI。</span><span class="sxs-lookup"><span data-stu-id="45254-104">You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office app ribbon.</span></span> <span data-ttu-id="45254-105">当用户单击命令按钮时，会执行一个操作，如打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="45254-105">When users click your command button, an action occurs, such as opening a task pane.</span></span>

<span data-ttu-id="45254-106">某些情况下，需要在文档打开时自动打开一个任务窗格，而无需进行显式用户交互。</span><span class="sxs-lookup"><span data-stu-id="45254-106">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction.</span></span> <span data-ttu-id="45254-107">可以使用 Addincommand 1.1 要求集中引入的 AutoOpen 任务窗格功能，以在情况需要时自动打开一个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="45254-107">You can use the autoopen task pane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it.</span></span>


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a><span data-ttu-id="45254-108">AutoOpen 功能与插入任务窗格有何不同？</span><span class="sxs-lookup"><span data-stu-id="45254-108">How is the autoopen feature different from inserting a task pane?</span></span>

<span data-ttu-id="45254-109">如果用户启动不使用外接程序命令的外接程序（例如，在 Office 2013 中运行的外接程序），外接程序会插入并保留在文档中。</span><span class="sxs-lookup"><span data-stu-id="45254-109">When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document.</span></span> <span data-ttu-id="45254-110">因此，当其他用户打开文档时，系统会提示他们安装外接程序，随后会打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="45254-110">As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens.</span></span> <span data-ttu-id="45254-111">此模型面临的挑战在于，在很多情况下，用户不希望外接程序在文档中保持。</span><span class="sxs-lookup"><span data-stu-id="45254-111">The challenge with this model is that in many cases, users don't want the add-in to persist in the document.</span></span> <span data-ttu-id="45254-112">例如，在 Word 文档中使用字典外接的学生可能不希望系统他们的同学或老师在打开该文档时提示他们安装该外接程序。</span><span class="sxs-lookup"><span data-stu-id="45254-112">For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.</span></span>

<span data-ttu-id="45254-113">使用 Autoopen 功能，可以显式定义或允许用户定义特定任务窗格外接程序是否保留在特定文档中。</span><span class="sxs-lookup"><span data-stu-id="45254-113">With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document.</span></span>

## <a name="support-and-availability"></a><span data-ttu-id="45254-114">支持和可用性</span><span class="sxs-lookup"><span data-stu-id="45254-114">Support and availability</span></span>

<span data-ttu-id="45254-115">目前支持在以下产品和平台中</span><span class="sxs-lookup"><span data-stu-id="45254-115">The autoopen feature is currently</span></span> <!-- in **developer preview** and it is only --> <span data-ttu-id="45254-116">使用 Autoopen 功能。</span><span class="sxs-lookup"><span data-stu-id="45254-116">supported in the following products and platforms.</span></span>

|<span data-ttu-id="45254-117">**产品**</span><span class="sxs-lookup"><span data-stu-id="45254-117">**Products**</span></span>|<span data-ttu-id="45254-118">**平台**</span><span class="sxs-lookup"><span data-stu-id="45254-118">**Platforms**</span></span>|
|:-----------|:------------|
|<ul><li><span data-ttu-id="45254-119">Word</span><span class="sxs-lookup"><span data-stu-id="45254-119">Word</span></span></li><li><span data-ttu-id="45254-120">Excel</span><span class="sxs-lookup"><span data-stu-id="45254-120">Excel</span></span></li><li><span data-ttu-id="45254-121">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="45254-121">PowerPoint</span></span></li></ul>|<span data-ttu-id="45254-122">所有产品的支持平台：</span><span class="sxs-lookup"><span data-stu-id="45254-122">Supported platforms for all products:</span></span><ul><li><span data-ttu-id="45254-123">Office on Windows Desktop.</span><span class="sxs-lookup"><span data-stu-id="45254-123">Office on Windows Desktop.</span></span> <span data-ttu-id="45254-124">Build 16.0.8121.1000+</span><span class="sxs-lookup"><span data-stu-id="45254-124">Build 16.0.8121.1000+</span></span></li><li><span data-ttu-id="45254-125">Office on Mac.</span><span class="sxs-lookup"><span data-stu-id="45254-125">Office on Mac.</span></span> <span data-ttu-id="45254-126">Build 15.34.17051500+</span><span class="sxs-lookup"><span data-stu-id="45254-126">Build 15.34.17051500+</span></span></li><li><span data-ttu-id="45254-127">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="45254-127">Office on the web</span></span></li></ul>|


## <a name="best-practices"></a><span data-ttu-id="45254-128">最佳做法</span><span class="sxs-lookup"><span data-stu-id="45254-128">Best practices</span></span>

<span data-ttu-id="45254-129">在使用 Autoopen 功能时应用下面的最佳做法：</span><span class="sxs-lookup"><span data-stu-id="45254-129">Apply the following best practices when you use the autoopen feature:</span></span>

- <span data-ttu-id="45254-130">当 Autoopen 功能可帮助外接程序用户工作更高效时使用此功能，如：</span><span class="sxs-lookup"><span data-stu-id="45254-130">Use the autoopen feature when it will help make your add-in users more efficient, such as:</span></span>
  - <span data-ttu-id="45254-131">When the document needs the add-in in order to function properly.</span><span class="sxs-lookup"><span data-stu-id="45254-131">When the document needs the add-in in order to function properly.</span></span> <span data-ttu-id="45254-132">For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in.</span><span class="sxs-lookup"><span data-stu-id="45254-132">For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in.</span></span> <span data-ttu-id="45254-133">The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span><span class="sxs-lookup"><span data-stu-id="45254-133">The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span></span>
  - <span data-ttu-id="45254-134">When the user will most likely always use the add-in with a particular document.</span><span class="sxs-lookup"><span data-stu-id="45254-134">When the user will most likely always use the add-in with a particular document.</span></span> <span data-ttu-id="45254-135">For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span><span class="sxs-lookup"><span data-stu-id="45254-135">For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span></span>
- <span data-ttu-id="45254-136">Allow users to turn on or turn off the autoopen feature.</span><span class="sxs-lookup"><span data-stu-id="45254-136">Allow users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="45254-137">Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span><span class="sxs-lookup"><span data-stu-id="45254-137">Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span></span>  
- <span data-ttu-id="45254-138">使用要求集检测来确定 autoopen 功能是否可用，如果不存在，则提供回退行为。</span><span class="sxs-lookup"><span data-stu-id="45254-138">Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn't.</span></span>
- <span data-ttu-id="45254-139">不要使用 Autoopen 功能来人为地增加外接程序的使用率。</span><span class="sxs-lookup"><span data-stu-id="45254-139">Don't use the autoopen feature to artificially increase usage of your add-in.</span></span> <span data-ttu-id="45254-140">如果你的外接程序无法在某些文档中自动打开，此功能可能会给用户增加烦恼。</span><span class="sxs-lookup"><span data-stu-id="45254-140">If it doesn't make sense for your add-in to open automatically with certain documents, this feature can annoy users.</span></span>

    > [!NOTE]
    > <span data-ttu-id="45254-141">如果 Microsoft 检测到滥用 AutoOpen 功能，加载项可能会从 AppSource 下架。</span><span class="sxs-lookup"><span data-stu-id="45254-141">If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource.</span></span>

- <span data-ttu-id="45254-142">Don't use this feature to pin multiple task panes.</span><span class="sxs-lookup"><span data-stu-id="45254-142">Don't use this feature to pin multiple task panes.</span></span> <span data-ttu-id="45254-143">You can only set one pane of your add-in to open automatically with a document.</span><span class="sxs-lookup"><span data-stu-id="45254-143">You can only set one pane of your add-in to open automatically with a document.</span></span>  

## <a name="implementation"></a><span data-ttu-id="45254-144">实现</span><span class="sxs-lookup"><span data-stu-id="45254-144">Implementation</span></span>

<span data-ttu-id="45254-145">要实现 Autoopen 功能，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="45254-145">To implement the autoopen feature:</span></span>

- <span data-ttu-id="45254-146">指定要自动打开的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="45254-146">Specify the task pane to be opened automatically.</span></span>
- <span data-ttu-id="45254-147">标记要自动打开任务窗格的文档。</span><span class="sxs-lookup"><span data-stu-id="45254-147">Tag the document to automatically open the task pane.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="45254-148">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device.</span><span class="sxs-lookup"><span data-stu-id="45254-148">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device.</span></span> <span data-ttu-id="45254-149">If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored.</span><span class="sxs-lookup"><span data-stu-id="45254-149">If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored.</span></span> <span data-ttu-id="45254-150">If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span><span class="sxs-lookup"><span data-stu-id="45254-150">If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span></span>

### <a name="step-1-specify-the-task-pane-to-open"></a><span data-ttu-id="45254-151">第 1 步：指定要打开的任务窗格</span><span class="sxs-lookup"><span data-stu-id="45254-151">Step 1: Specify the task pane to open</span></span>

<span data-ttu-id="45254-152">To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**.</span><span class="sxs-lookup"><span data-stu-id="45254-152">To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**.</span></span> <span data-ttu-id="45254-153">You can only set this value on one task pane.</span><span class="sxs-lookup"><span data-stu-id="45254-153">You can only set this value on one task pane.</span></span> <span data-ttu-id="45254-154">If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span><span class="sxs-lookup"><span data-stu-id="45254-154">If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span></span>

<span data-ttu-id="45254-155">在下面的示例中，TaskPaneId 值设置为 Office.AutoShowTaskpaneWithDocument。</span><span class="sxs-lookup"><span data-stu-id="45254-155">The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a><span data-ttu-id="45254-156">第 2 步：将文档标记为自动打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="45254-156">Step 2: Tag the document to automatically open the task pane</span></span>

<span data-ttu-id="45254-157">You can tag the document to trigger the autoopen feature in one of two ways.</span><span class="sxs-lookup"><span data-stu-id="45254-157">You can tag the document to trigger the autoopen feature in one of two ways.</span></span> <span data-ttu-id="45254-158">Pick the alternative that works best for your scenario.</span><span class="sxs-lookup"><span data-stu-id="45254-158">Pick the alternative that works best for your scenario.</span></span>  


#### <a name="tag-the-document-on-the-client-side"></a><span data-ttu-id="45254-159">在客户端上标记文档</span><span class="sxs-lookup"><span data-stu-id="45254-159">Tag the document on the client side</span></span>

<span data-ttu-id="45254-160">使用 Office.js [settings.set](/javascript/api/office/office.settings) 方法将 **Office.AutoShowTaskpaneWithDocument** 设置为“**true**”，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="45254-160">Use the Office.js [settings.set](/javascript/api/office/office.settings) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.</span></span>

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

<span data-ttu-id="45254-161">如果需要将文档标记为外接程序交互的一部分（例如，在用户创建一个绑定，或选择一个选项来表示他们希望窗格自动打开时），则使用此方法。</span><span class="sxs-lookup"><span data-stu-id="45254-161">Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).</span></span>

#### <a name="use-open-xml-to-tag-the-document"></a><span data-ttu-id="45254-162">使用 Open XML 标记文档</span><span class="sxs-lookup"><span data-stu-id="45254-162">Use Open XML to tag the document</span></span>

<span data-ttu-id="45254-163">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature.</span><span class="sxs-lookup"><span data-stu-id="45254-163">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature.</span></span> <span data-ttu-id="45254-164">For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span><span class="sxs-lookup"><span data-stu-id="45254-164">For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span></span>

<span data-ttu-id="45254-165">向文档添加两个 Open XML 部件：</span><span class="sxs-lookup"><span data-stu-id="45254-165">Add two Open XML parts to the document:</span></span>

- <span data-ttu-id="45254-166">一个 `webextension` 部件</span><span class="sxs-lookup"><span data-stu-id="45254-166">A `webextension` part</span></span>
- <span data-ttu-id="45254-167">一个 `taskpane` 部件</span><span class="sxs-lookup"><span data-stu-id="45254-167">A `taskpane` part</span></span>

<span data-ttu-id="45254-168">以下示例演示如何添加 `webextension` 部件。</span><span class="sxs-lookup"><span data-stu-id="45254-168">The following example shows how to add the `webextension` part.</span></span>

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

<span data-ttu-id="45254-169">`webextension` 部件包含一个属性包，以及必须设置为 `true` 的 **Office.AutoShowTaskpaneWithDocument** 属性。</span><span class="sxs-lookup"><span data-stu-id="45254-169">The `webextension` part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.</span></span>

<span data-ttu-id="45254-170">`webextension` 部件还包括对具有 `id`、`storeType`、`store` 和 `version` 的属性的应用商店或目录的引用。</span><span class="sxs-lookup"><span data-stu-id="45254-170">The `webextension` part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`.</span></span> <span data-ttu-id="45254-171">在 `storeType` 值中，只有四个与 AutoOpen 功能相关。</span><span class="sxs-lookup"><span data-stu-id="45254-171">Of the `storeType` values, only four are relevant to the autoopen feature.</span></span> <span data-ttu-id="45254-172">其他三个属性的值取决于 `storeType` 的值，如下表所示。</span><span class="sxs-lookup"><span data-stu-id="45254-172">The values for the other three attributes depend on the value for `storeType`, as shown in the following table.</span></span>

| <span data-ttu-id="45254-173">**`storeType` 值**</span><span class="sxs-lookup"><span data-stu-id="45254-173">**`storeType` value**</span></span> | <span data-ttu-id="45254-174">**`id` 值**</span><span class="sxs-lookup"><span data-stu-id="45254-174">**`id` value**</span></span>    |<span data-ttu-id="45254-175">**`store` 值**</span><span class="sxs-lookup"><span data-stu-id="45254-175">**`store` value**</span></span> | <span data-ttu-id="45254-176">**`version` 值**</span><span class="sxs-lookup"><span data-stu-id="45254-176">**`version` value**</span></span>|
|:---------------|:---------------|:---------------|:---------------|
|<span data-ttu-id="45254-177">OMEX (AppSource)</span><span class="sxs-lookup"><span data-stu-id="45254-177">OMEX (AppSource)</span></span>|<span data-ttu-id="45254-178">加载项的 AppSource 资产 ID（请参阅“注意”）</span><span class="sxs-lookup"><span data-stu-id="45254-178">The AppSource asset ID of the add-in (see Note)</span></span>|<span data-ttu-id="45254-179">AppSource 的区域设置；例如，“en-us”。</span><span class="sxs-lookup"><span data-stu-id="45254-179">The locale of AppSource; for example, "en-us".</span></span>|<span data-ttu-id="45254-180">AppSource 目录中的版本（请参阅“注意”）</span><span class="sxs-lookup"><span data-stu-id="45254-180">The version in the AppSource catalog (see Note)</span></span>|
|<span data-ttu-id="45254-181">FileSystem（网络共享）</span><span class="sxs-lookup"><span data-stu-id="45254-181">FileSystem (a network share)</span></span>|<span data-ttu-id="45254-182">外接程序清单中外接程序的 GUID。</span><span class="sxs-lookup"><span data-stu-id="45254-182">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="45254-183">网络共享路径。例如，“\\\\MyComputer\\MySharedFolder”。</span><span class="sxs-lookup"><span data-stu-id="45254-183">The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span>|<span data-ttu-id="45254-184">外接程序清单中的版本。</span><span class="sxs-lookup"><span data-stu-id="45254-184">The version in the add-in manifest.</span></span>|
|<span data-ttu-id="45254-185">EXCatalog（通过 Exchange 服务器部署）</span><span class="sxs-lookup"><span data-stu-id="45254-185">EXCatalog (deployment via the Exchange server)</span></span> |<span data-ttu-id="45254-186">外接程序清单中外接程序的 GUID。</span><span class="sxs-lookup"><span data-stu-id="45254-186">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="45254-187">“EXCatalog”。</span><span class="sxs-lookup"><span data-stu-id="45254-187">"EXCatalog".</span></span> <span data-ttu-id="45254-188">EXCatalog row 是在 Microsoft 365 管理中心使用集中部署的外接程序要使用的行。</span><span class="sxs-lookup"><span data-stu-id="45254-188">EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Microsoft 365 admin center.</span></span>|<span data-ttu-id="45254-189">外接程序清单中的版本。</span><span class="sxs-lookup"><span data-stu-id="45254-189">The version in the add-in manifest.</span></span>
|<span data-ttu-id="45254-190">Registry（系统注册表）</span><span class="sxs-lookup"><span data-stu-id="45254-190">Registry (System registry)</span></span>|<span data-ttu-id="45254-191">外接程序清单中外接程序的 GUID。</span><span class="sxs-lookup"><span data-stu-id="45254-191">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="45254-192">“developer”</span><span class="sxs-lookup"><span data-stu-id="45254-192">"developer"</span></span>|<span data-ttu-id="45254-193">加载项清单中的版本。</span><span class="sxs-lookup"><span data-stu-id="45254-193">The version in the add-in manifest.</span></span>|

> [!NOTE]
> <span data-ttu-id="45254-194">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in.</span><span class="sxs-lookup"><span data-stu-id="45254-194">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in.</span></span> <span data-ttu-id="45254-195">The asset ID appears in the address bar in the browser.</span><span class="sxs-lookup"><span data-stu-id="45254-195">The asset ID appears in the address bar in the browser.</span></span> <span data-ttu-id="45254-196">The version is listed in the **Details** section of the page.</span><span class="sxs-lookup"><span data-stu-id="45254-196">The version is listed in the **Details** section of the page.</span></span>

<span data-ttu-id="45254-197">若要详细了解 webextension 标记，请参阅 [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)。</span><span class="sxs-lookup"><span data-stu-id="45254-197">For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span></span>

<span data-ttu-id="45254-198">以下示例演示如何添加 `taskpane` 部件。</span><span class="sxs-lookup"><span data-stu-id="45254-198">The following example shows how to add the `taskpane` part.</span></span>

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

<span data-ttu-id="45254-199">请注意，在本例中，`visibility` 属性设置为“0”。</span><span class="sxs-lookup"><span data-stu-id="45254-199">Note that in this example, the `visibility` attribute is set to "0".</span></span> <span data-ttu-id="45254-200">这意味着在添加 webextension 部件和 `taskpane` 部件之后，第一次打开文档时，用户还必须从功能区上的“外接程序”\*\*\*\* 按钮安装该外接程序。</span><span class="sxs-lookup"><span data-stu-id="45254-200">This means that after the webextension and `taskpane` parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon.</span></span> <span data-ttu-id="45254-201">此后，外接程序任务窗格将在打开该文件时自动打开。</span><span class="sxs-lookup"><span data-stu-id="45254-201">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="45254-202">此外，在将 `visibility` 设置为“0”时，可以使用 Office.js 让用户打开或关闭 AutoOpen 功能。</span><span class="sxs-lookup"><span data-stu-id="45254-202">Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="45254-203">具体来说，脚本会将 **Office.AutoShowTaskpaneWithDocument** 文档设置为 `true` 或 `false`。</span><span class="sxs-lookup"><span data-stu-id="45254-203">Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`.</span></span> <span data-ttu-id="45254-204">（有关详细信息，请参阅[在客户端上标记文档](#tag-the-document-on-the-client-side)。）</span><span class="sxs-lookup"><span data-stu-id="45254-204">(For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).)</span></span>

<span data-ttu-id="45254-205">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened.</span><span class="sxs-lookup"><span data-stu-id="45254-205">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened.</span></span> <span data-ttu-id="45254-206">The user is prompted to trust the add-in, and when trust is granted, the add-in opens.</span><span class="sxs-lookup"><span data-stu-id="45254-206">The user is prompted to trust the add-in, and when trust is granted, the add-in opens.</span></span> <span data-ttu-id="45254-207">Thereafter, the add-in task pane opens automatically when the file is opened.</span><span class="sxs-lookup"><span data-stu-id="45254-207">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="45254-208">However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span><span class="sxs-lookup"><span data-stu-id="45254-208">However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span></span>

<span data-ttu-id="45254-209">当外接程序和模板或文档内容紧密集成以致用户不会选择退出 Autoopen 功能时，将 `visibility` 设置为“1”是一个不错的选择。</span><span class="sxs-lookup"><span data-stu-id="45254-209">Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature.</span></span>

> [!NOTE]
> <span data-ttu-id="45254-210">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1.</span><span class="sxs-lookup"><span data-stu-id="45254-210">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1.</span></span> <span data-ttu-id="45254-211">You can only do this via Open XML.</span><span class="sxs-lookup"><span data-stu-id="45254-211">You can only do this via Open XML.</span></span>

<span data-ttu-id="45254-212">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated.</span><span class="sxs-lookup"><span data-stu-id="45254-212">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated.</span></span> <span data-ttu-id="45254-213">Office will detect and provide the appropriate attribute values.</span><span class="sxs-lookup"><span data-stu-id="45254-213">Office will detect and provide the appropriate attribute values.</span></span> <span data-ttu-id="45254-214">You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span><span class="sxs-lookup"><span data-stu-id="45254-214">You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span></span>

## <a name="test-and-verify-opening-task-panes"></a><span data-ttu-id="45254-215">对打开任务窗格进行测试和验证</span><span class="sxs-lookup"><span data-stu-id="45254-215">Test and verify opening task panes</span></span>

<span data-ttu-id="45254-216">您可以部署外接程序的测试版本，它将通过 Microsoft 365 管理中心使用集中部署自动打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="45254-216">You can deploy a test version of your add-in that will automatically open a task pane using Centralized Deployment via the Microsoft 365 admin center.</span></span> <span data-ttu-id="45254-217">以下示例演示如何使用 EXCatalog 应用商店版本从集中部署目录插入外接程序。</span><span class="sxs-lookup"><span data-stu-id="45254-217">The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="45254-218">您可以使用 Microsoft 365 订阅测试上一个示例，以尝试进行集中部署，并验证您的外接程序是否按预期工作。</span><span class="sxs-lookup"><span data-stu-id="45254-218">You can test the previous example by using your Microsoft 365 subscription to try out Centralized Deployment and verify that your add-in works as expected.</span></span> <span data-ttu-id="45254-219">如果你还没有 Microsoft 365 订阅，则可以通过加入[microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取免费的90天 renewable microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="45254-219">If you don't already have a Microsoft 365 subscription, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="see-also"></a><span data-ttu-id="45254-220">另请参阅</span><span class="sxs-lookup"><span data-stu-id="45254-220">See also</span></span>

<span data-ttu-id="45254-221">有关演示如何使用 AutoOpen 功能的示例，请参阅 [Office 外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)。</span><span class="sxs-lookup"><span data-stu-id="45254-221">For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).</span></span>
<span data-ttu-id="45254-222">[加入 Microsoft 365 开发人员计划](/office/developer-program/office-365-developer-program)。</span><span class="sxs-lookup"><span data-stu-id="45254-222">[Join the Microsoft 365 developer program](/office/developer-program/office-365-developer-program).</span></span>
