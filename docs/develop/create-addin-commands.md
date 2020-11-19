---
title: 在清单中创建 Excel、PowerPoint 和 Word 加载项命令
description: 在清单中使用 VersionOverrides 定义 Excel、PowerPoint 和 Word 的外接程序命令。使用外接命令创建 UI 元素、添加按钮或列表并执行操作。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 9257e7ba840db31149ae606c7f2c072c433140ad
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131917"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a><span data-ttu-id="50791-104">在清单中创建 Excel、PowerPoint 和 Word 加载项命令</span><span class="sxs-lookup"><span data-stu-id="50791-104">Create add-in commands in your manifest for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="50791-p102">在清单中使用 **[VersionOverrides](../reference/manifest/versionoverrides.md)** 定义 Excel、PowerPoint 和 Word 的外接程序命令。外接程序命令提供了使用执行操作的指定 UI 元素) 自定义默认 Office 用户界面 (UI 的简单方法。您可以使用外接命令执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="50791-p102">Use **[VersionOverrides](../reference/manifest/versionoverrides.md)** in your manifest to define add-in commands for Excel, PowerPoint, and Word. Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. You can use add-in commands to:</span></span>

- <span data-ttu-id="50791-108">创建 UI 元素或入口点，以便能够更易于使用你的外接程序功能。</span><span class="sxs-lookup"><span data-stu-id="50791-108">Create UI elements or entry points that make your add-in's functionality easier to use.</span></span>
- <span data-ttu-id="50791-109">向功能区中添加按钮或下拉列表按钮。</span><span class="sxs-lookup"><span data-stu-id="50791-109">Add buttons or a drop-down list of buttons to the ribbon.</span></span>
- <span data-ttu-id="50791-110">将单个菜单项（每一个都包含可选的子菜单）添加到特定上下文（快捷方式）菜单中。</span><span class="sxs-lookup"><span data-stu-id="50791-110">Add individual menu items — each containing optional submenus — to specific context (shortcut) menus.</span></span>
- <span data-ttu-id="50791-p103">在选择你的外接程序命令时执行操作。可以：</span><span class="sxs-lookup"><span data-stu-id="50791-p103">Perform actions when your add-in command is chosen. You can:</span></span>
  - <span data-ttu-id="50791-p104">显示一个或多个任务窗格外接程序，让用户与其进行交互。在任务窗格外接程序内，可以显示使用 Office UI 结构创建自定义 UI 的 HTML。</span><span class="sxs-lookup"><span data-stu-id="50791-p104">Show one or more task pane add-ins for users to interact with. Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span></span>

     <span data-ttu-id="50791-115">*或者*</span><span class="sxs-lookup"><span data-stu-id="50791-115">*or*</span></span>

  - <span data-ttu-id="50791-116">运行 JavaScript 代码，该代码通常在不显示任何 UI 的情况下运行。</span><span class="sxs-lookup"><span data-stu-id="50791-116">Run JavaScript code, which normally runs without displaying any UI.</span></span>

<span data-ttu-id="50791-p105">本文介绍如何编辑您的清单来定义外接程序命令。下图显示了用来定义外接程序命令的元素的层次结构。本文将具体介绍这些元素。</span><span class="sxs-lookup"><span data-stu-id="50791-p105">This article describes how to edit your manifest to define add-in commands. The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="50791-120">Outlook 中也支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="50791-120">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="50791-121">有关详细信息，请参阅 [适用于 Outlook 的外接程序命令](../outlook/add-in-commands-for-outlook.md)</span><span class="sxs-lookup"><span data-stu-id="50791-121">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md)</span></span>

<span data-ttu-id="50791-122">下图是对清单中的加载项命令元素的概述。</span><span class="sxs-lookup"><span data-stu-id="50791-122">The following image is an overview of add-in commands elements in the manifest.</span></span>

![清单中的外接命令元素的概述。](../images/version-overrides.png)

## <a name="step-1-start-from-a-sample"></a><span data-ttu-id="50791-131">第 1 步：从示例入手</span><span class="sxs-lookup"><span data-stu-id="50791-131">Step 1: Start from a sample</span></span>

<span data-ttu-id="50791-p108">强烈建议从 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Command-Sample)中的示例之一入手。也可以按照本指南中的步骤操作，创建自己的清单。可以使用“Office 加载项命令示例”网站中的 XSD 文件来验证清单。使用加载项命令前，请确保已阅读 [Excel、Word 和 PowerPoint 加载项命令](../design/add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="50791-p108">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Optionally, you can create your own manifest by following the steps in this guide. You can validate your manifest using the XSD file in the Office Add-in Commands Samples site. Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span></span>

## <a name="step-2-create-a-task-pane-add-in"></a><span data-ttu-id="50791-136">第 2 步：创建任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="50791-136">Step 2: Create a task pane add-in</span></span>

<span data-ttu-id="50791-137">若要开始使用加载项命令，必须先创建任务窗格加载项，然后按本文所述修改加载项的清单。</span><span class="sxs-lookup"><span data-stu-id="50791-137">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article.</span></span> <span data-ttu-id="50791-138">不能将外接程序命令与内容外接程序一起使用。如果要更新现有清单，则必须添加相应的 **XML 命名空间** ，并将 **VersionOverrides** 元素添加到清单中（如 [步骤3： add VersionOverrides 元素](#step-3-add-versionoverrides-element)中所述）。</span><span class="sxs-lookup"><span data-stu-id="50791-138">You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropriate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span></span>

<span data-ttu-id="50791-p110">以下示例显示了 Office 2013 外接程序的清单。此清单中没有任何外接程序命令，因为没有 **VersionOverrides** 元素。Office 2013 不支持外接程序命令，但是通过将 **VersionOverrides** 添加到此清单，外接程序可同时在 Office 2013 和 Office 2016 中运行。在 Office 2013 中，外接程序不会显示外接程序命令，并且使用 **SourceLocation** 的值运行外接程序作为单一任务窗格外接程序。在 Office 2016 中，如果未包含 **VersionOverrides** 元素，则使用 **SourceLocation** 运行外接程序。但是，如果包含了 **VersionOverrides**，外接程序将只显示外接程序命令，并且不会将外接程序显示为单一任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="50791-p110">The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **VersionOverrides** element. Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in. In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in. If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span></span>
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## <a name="step-3-add-versionoverrides-element"></a><span data-ttu-id="50791-145">步骤 3：添加 VersionOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="50791-145">Step 3: Add VersionOverrides element</span></span>

<span data-ttu-id="50791-p111">**VersionOverrides** 元素是包含外接程序命令定义的根元素。**VersionOverrides** 是清单中 **OfficeApp** 元素的子元素。下表列出了 **VersionOverrides** 元素的属性。</span><span class="sxs-lookup"><span data-stu-id="50791-p111">The **VersionOverrides** element is the root element that contains the definition of your add-in command. **VersionOverrides** is a child element of the **OfficeApp** element in the manifest. The following table lists the attributes of the **VersionOverrides** element.</span></span>

|<span data-ttu-id="50791-149">属性</span><span class="sxs-lookup"><span data-stu-id="50791-149">Attribute</span></span>|<span data-ttu-id="50791-150">说明</span><span class="sxs-lookup"><span data-stu-id="50791-150">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-151">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="50791-151">**xmlns**</span></span> <br/> | <span data-ttu-id="50791-152">必需。</span><span class="sxs-lookup"><span data-stu-id="50791-152">Required.</span></span> <span data-ttu-id="50791-153">架构位置必须是 `http://schemas.microsoft.com/office/taskpaneappversionoverrides`。</span><span class="sxs-lookup"><span data-stu-id="50791-153">The schema location, which must be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span> <br/> |
|<span data-ttu-id="50791-154">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="50791-154">**xsi:type**</span></span> <br/> |<span data-ttu-id="50791-p113">必需。架构版本。本文中所述的版本为"VersionOverridesV1_0"。</span><span class="sxs-lookup"><span data-stu-id="50791-p113">Required. The schema version. The version described in this article is "VersionOverridesV1_0".</span></span>  <br/> |

<span data-ttu-id="50791-158">下表标识了 **VersionOverrides** 的子元素。</span><span class="sxs-lookup"><span data-stu-id="50791-158">The following table identifies the child elements of **VersionOverrides**.</span></span>
  
|<span data-ttu-id="50791-159">元素</span><span class="sxs-lookup"><span data-stu-id="50791-159">Element</span></span>|<span data-ttu-id="50791-160">说明</span><span class="sxs-lookup"><span data-stu-id="50791-160">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-161">**说明**</span><span class="sxs-lookup"><span data-stu-id="50791-161">**Description**</span></span> <br/> |<span data-ttu-id="50791-p114">可选。描述外接程序。此子级 **Description** 元素替代清单中父级部分中的旧 **Description** 元素。此 **Description** 元素的 **resid** 属性将设置为 **String** 元素的 **id**。**String** 元素包含 **Description** 的文本。 </span><span class="sxs-lookup"><span data-stu-id="50791-p114">Optional. Describes the add-in. This child **Description** element overrides a previous **Description** element in the parent portion of the manifest. The **resid** attribute for this **Description** element is set to the **id** of a **String** element. The **String** element contains the text for **Description**. </span></span><br/> |
|<span data-ttu-id="50791-167">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="50791-167">**Requirements**</span></span> <br/> |<span data-ttu-id="50791-168">可选。</span><span class="sxs-lookup"><span data-stu-id="50791-168">Optional.</span></span> <span data-ttu-id="50791-169">指定外接程序要求的最低要求集和 Office.js 的版本。</span><span class="sxs-lookup"><span data-stu-id="50791-169">Specifies the minimum requirement set and version of Office.js that the add-in requires.</span></span> <span data-ttu-id="50791-170">此子级 **Requirements** 元素替代清单中父级部分中的 **Requirements** 元素。</span><span class="sxs-lookup"><span data-stu-id="50791-170">This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest.</span></span> <span data-ttu-id="50791-171">有关详细信息，请参阅 [指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="50791-171">For more information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>  <br/> |
|<span data-ttu-id="50791-172">**Hosts**</span><span class="sxs-lookup"><span data-stu-id="50791-172">**Hosts**</span></span> <br/> |<span data-ttu-id="50791-173">必需。</span><span class="sxs-lookup"><span data-stu-id="50791-173">Required.</span></span> <span data-ttu-id="50791-174">指定 Office 应用程序的集合。</span><span class="sxs-lookup"><span data-stu-id="50791-174">Specifies a collection of Office applications.</span></span> <span data-ttu-id="50791-175">子级 **Hosts** 元素替代清单中父级部分中的 **Hosts** 元素。</span><span class="sxs-lookup"><span data-stu-id="50791-175">The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest.</span></span> <span data-ttu-id="50791-176">必须包含已设置为“Workbook”或“Document”的 **xsi:type** 属性</span><span class="sxs-lookup"><span data-stu-id="50791-176">You must include a **xsi:type** attribute set to "Workbook" or "Document".</span></span> <br/> |
|<span data-ttu-id="50791-177">**Resources**</span><span class="sxs-lookup"><span data-stu-id="50791-177">**Resources**</span></span> <br/> |<span data-ttu-id="50791-p117">定义其他清单元素引用的资源集合（字符串、URL 和图像）。例如，**Description** 元素的值引用了 **Resources** 中的子元素。**Resources** 元素将在本文后续部分中的 [步骤 7：添加 Resources 元素](#step-7-add-the-resources-element)中进行介绍。 </span><span class="sxs-lookup"><span data-stu-id="50791-p117">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **Description** element's value refers to a child element in **Resources**. The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. </span></span><br/> |

<span data-ttu-id="50791-181">下面的示例演示如何使用 **VersionOverrides** 元素及其子元素。</span><span class="sxs-lookup"><span data-stu-id="50791-181">The following example shows how to use the **VersionOverrides** element and its child elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a><span data-ttu-id="50791-182">步骤 4：添加 Hosts、Host 和 DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="50791-182">Step 4: Add Hosts, Host, and DesktopFormFactor elements</span></span>

<span data-ttu-id="50791-183">“Hosts”元素包含一个或多个“Host”元素。</span><span class="sxs-lookup"><span data-stu-id="50791-183">The **Hosts** element contains one or more **Host** elements.</span></span> <span data-ttu-id="50791-184">**Host** 元素指定特定的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="50791-184">A **Host** element specifies a particular Office application.</span></span> <span data-ttu-id="50791-185">**Host** 元素包含子元素，这些子元素指定在将外接程序安装在该 Office 应用程序中后要显示的外接程序命令。</span><span class="sxs-lookup"><span data-stu-id="50791-185">The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office application.</span></span> <span data-ttu-id="50791-186">若要在两个或更多不同的 Office 应用程序中显示相同的外接程序命令，必须复制每个 **主机** 中的子元素。</span><span class="sxs-lookup"><span data-stu-id="50791-186">To show the same add-in commands in two or more different Office applications, you must duplicate the child elements in each **Host**.</span></span>

<span data-ttu-id="50791-187">**DesktopFormFactor** 元素指定在 Office 网页版（浏览器版）和 Windows 版 Office 中运行的加载项的设置。</span><span class="sxs-lookup"><span data-stu-id="50791-187">The **DesktopFormFactor** element specifies the settings for an add-in that runs in Office on the web (in a browser) and Windows.</span></span>

<span data-ttu-id="50791-188">以下是一个包含 **Hosts**、**Host** 和 **DesktopFormFactor** 元素的示例。</span><span class="sxs-lookup"><span data-stu-id="50791-188">The following is an example of **Hosts**, **Host**, and **DesktopFormFactor** elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-5-add-the-functionfile-element"></a><span data-ttu-id="50791-189">步骤 5：添加 FunctionFile 元素</span><span class="sxs-lookup"><span data-stu-id="50791-189">Step 5: Add the FunctionFile element</span></span>

<span data-ttu-id="50791-p119">"FunctionFile"元素指定了一个文件，其中包含当外接程序命令使用"ExecuteFunction"操作时要运行的 JavaScript 代码（请参阅 按钮控件了解相关说明）。将"FunctionFile"元素的"resid"属性设置为包括外接程序命令需要的所有 JavaScript 文件的 HTML 文件。不能只链接到 JavaScript 文件。将文件名称指定为"Resources"元素中的"Url"元素。</span><span class="sxs-lookup"><span data-stu-id="50791-p119">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](../reference/manifest/control.md#button-control) for a description). The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **Url** element in the **Resources** element.</span></span>

<span data-ttu-id="50791-195">下面的示例展示了 **FunctionFile** 元素。</span><span class="sxs-lookup"><span data-stu-id="50791-195">The following is an example of the **FunctionFile** element.</span></span>
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint>

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> <span data-ttu-id="50791-196">请确保 JavaScript 代码调用了 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="50791-196">Make sure your JavaScript code calls  `Office.initialize`.</span></span>

<span data-ttu-id="50791-p120">**FunctionFile** 元素引用的 HTML 文件中的 JavaScript 必须调用 `Office.initialize`。**FunctionName** 元素（请参阅 [按钮控件](../reference/manifest/control.md#button-control)查看相关说明）使用 **FunctionFile** 中的函数。</span><span class="sxs-lookup"><span data-stu-id="50791-p120">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`. The **FunctionName** element (see [Button controls](../reference/manifest/control.md#button-control) for a description) uses the functions in **FunctionFile**.</span></span>

<span data-ttu-id="50791-199">下面的代码展示了如何实现 **FunctionName** 使用的函数。</span><span class="sxs-lookup"><span data-stu-id="50791-199">The following code shows how to implement the function used by **FunctionName**.</span></span>

```js
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Your function must be in the global namespace.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // Show error message.
                }
                else {
                    // Show success message.
                }
            });

        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
    }
</script>
```

> [!IMPORTANT]
> <span data-ttu-id="50791-p121">调用 **event.completed** 表示已成功处理事件。如果函数获得多次调用（如多次单击同一加载项命令），所有事件都会自动排入队列。首个事件会自动运行，而其他事件则继续留在队列中。如果函数调用 **event.completed**，将运行此函数在队列中的下一个调用。必须实现 **event.completed**，否则函数不会运行。</span><span class="sxs-lookup"><span data-stu-id="50791-p121">The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.</span></span>

## <a name="step-6-add-extensionpoint-elements"></a><span data-ttu-id="50791-205">第 6 步：添加 ExtensionPoint 元素</span><span class="sxs-lookup"><span data-stu-id="50791-205">Step 6: Add ExtensionPoint elements</span></span>

<span data-ttu-id="50791-p122">**ExtensionPoint** 元素定义外接程序命令应在 Office UI 中的哪个位置出现。可以使用以下 **xsi:type** 值定义 **ExtensionPoint** 元素：</span><span class="sxs-lookup"><span data-stu-id="50791-p122">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI. You can define **ExtensionPoint** elements with these **xsi:type** values:</span></span>

- <span data-ttu-id="50791-208">**PrimaryCommandSurface**，它是指 Office 中的功能区。</span><span class="sxs-lookup"><span data-stu-id="50791-208">**PrimaryCommandSurface**, which refers to the ribbon in Office.</span></span>

- <span data-ttu-id="50791-209">**ContextMenu**，它是当你在 Office UI 中右键单击时出现的快捷菜单。</span><span class="sxs-lookup"><span data-stu-id="50791-209">**ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="50791-210">下面的示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="50791-210">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="50791-p123">对于包含 ID 属性的元素，请务必提供唯一 ID。建议将公司名称与 ID 结合使用。例如，请使用以下格式：`<CustomTab id="mycompanyname.mygroupname">`。</span><span class="sxs-lookup"><span data-stu-id="50791-p123">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span></span>
  
```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso Tab">
  <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
  <!-- <OfficeTab id="TabData"> -->
    <Label resid="residLabel4" />
    <Group id="Group1Id12">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Button1Id1">

        <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

|<span data-ttu-id="50791-214">元素</span><span class="sxs-lookup"><span data-stu-id="50791-214">Element</span></span>|<span data-ttu-id="50791-215">说明</span><span class="sxs-lookup"><span data-stu-id="50791-215">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-216">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="50791-216">**CustomTab**</span></span> <br/> |<span data-ttu-id="50791-p124">如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。 </span><span class="sxs-lookup"><span data-stu-id="50791-p124">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required. </span></span><br/> |
|<span data-ttu-id="50791-220">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="50791-220">**OfficeTab**</span></span> <br/> |<span data-ttu-id="50791-221">如果要使用 **PrimaryCommandSurface**) 扩展默认的 Office 应用功能区选项卡 (，则为必需。</span><span class="sxs-lookup"><span data-stu-id="50791-221">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="50791-222">如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。</span><span class="sxs-lookup"><span data-stu-id="50791-222">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <br/> <span data-ttu-id="50791-223">有关 **id** 属性使用的更多 tab 值，请参阅 [默认 Office 应用功能区选项卡的 tab 值](../reference/manifest/officetab.md)。</span><span class="sxs-lookup"><span data-stu-id="50791-223">For more tab values to use with the **id** attribute, see [Tab values for default Office app ribbon tabs](../reference/manifest/officetab.md).</span></span>  <br/> |
|<span data-ttu-id="50791-224">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="50791-224">**OfficeMenu**</span></span> <br/> | <span data-ttu-id="50791-p126">如果要（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： </span><span class="sxs-lookup"><span data-stu-id="50791-p126">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="50791-p127">当用户选定文本，然后右键单击所选文本时，适用于 Excel 或 Word 的 **ContextMenuText** 显示上下文菜单上的项。</span><span class="sxs-lookup"><span data-stu-id="50791-p127">**ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="50791-p128">适用于 Excel 的 **ContextMenuCell**。当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。 </span><span class="sxs-lookup"><span data-stu-id="50791-p128">**ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. </span></span><br/> |
|<span data-ttu-id="50791-231">**Group**</span><span class="sxs-lookup"><span data-stu-id="50791-231">**Group**</span></span> <br/> |<span data-ttu-id="50791-p129">选项卡上的一组用户界面扩展点。一组可以有多达六个控件。**id** 属性是必需的。它是一个最多为 125 个字符的字符串。 </span><span class="sxs-lookup"><span data-stu-id="50791-p129">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="50791-235">**Label**</span><span class="sxs-lookup"><span data-stu-id="50791-235">**Label**</span></span> <br/> |<span data-ttu-id="50791-p130">必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p130">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-240">**Icon**</span><span class="sxs-lookup"><span data-stu-id="50791-240">**Icon**</span></span> <br/> |<span data-ttu-id="50791-p131">必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。 </span><span class="sxs-lookup"><span data-stu-id="50791-p131">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="50791-248">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="50791-248">**Tooltip**</span></span> <br/> |<span data-ttu-id="50791-p132">可选。组的工具提示 **resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p132">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-253">**Control**</span><span class="sxs-lookup"><span data-stu-id="50791-253">**Control**</span></span> <br/> |<span data-ttu-id="50791-p133">每个组都要求至少有一个控件。**Control** 元素可以是 **Button**，也可以是 **Menu**。使用 **Menu** 可指定按钮控件的下拉列表。目前仅支持按钮和菜单。请参阅 [按钮控件](../reference/manifest/control.md#button-control)和 [菜单控件](../reference/manifest/control.md#menu-dropdown-button-controls)部分，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="50791-p133">Each group requires at least one control. A **Control** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the  [Button controls](../reference/manifest/control.md#button-control) and [Menu controls](../reference/manifest/control.md#menu-dropdown-button-controls) sections for more information. </span></span><br/><span data-ttu-id="50791-259">**注意：** 建议一次添加一个 **Control** 元素及相关 **Resources** 子元素，以便于进行故障排除。</span><span class="sxs-lookup"><span data-stu-id="50791-259">**Note:** To make troubleshooting easier, we recommend that you add a **Control** element and the related **Resources** child elements one at a time.</span></span>          |

### <a name="button-controls"></a><span data-ttu-id="50791-260">按钮控件</span><span class="sxs-lookup"><span data-stu-id="50791-260">Button controls</span></span>

<span data-ttu-id="50791-p134">当用户选择某个按钮时，将执行一个操作。它可以执行 JavaScript 函数或显示任务窗格。以下示例演示了如何定义两种按钮。第一个按钮在不显示 UI 的情况下运行 JavaScript 函数，第二个按钮显示任务窗格。在 **Control** 元素中：</span><span class="sxs-lookup"><span data-stu-id="50791-p134">A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **Control** element:</span></span>

- <span data-ttu-id="50791-266">**type** 属性是必需的，并且必须设置为 **Button**。</span><span class="sxs-lookup"><span data-stu-id="50791-266">The **type** attribute is required, and must be set to **Button**.</span></span>

- <span data-ttu-id="50791-267">**Control** 元素的 **id** 属性是一个最多为 125 个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="50791-267">The **id** attribute of the **Control** element is a string with a maximum of 125 characters.</span></span>

```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|<span data-ttu-id="50791-268">元素</span><span class="sxs-lookup"><span data-stu-id="50791-268">Elements</span></span>|<span data-ttu-id="50791-269">说明</span><span class="sxs-lookup"><span data-stu-id="50791-269">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-270">**Label**</span><span class="sxs-lookup"><span data-stu-id="50791-270">**Label**</span></span> <br/> |<span data-ttu-id="50791-p135">必需。按钮文本。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p135">Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-275">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="50791-275">**Tooltip**</span></span> <br/> |<span data-ttu-id="50791-p136">可选。按钮的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p136">Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-280">**Supertip**</span><span class="sxs-lookup"><span data-stu-id="50791-280">**Supertip**</span></span> <br/> | <span data-ttu-id="50791-p137">必需。此按钮的 SuperTip，定义如下： </span><span class="sxs-lookup"><span data-stu-id="50791-p137">Required. The supertip for this button, which is defined by the following: </span></span><br/> <span data-ttu-id="50791-283">**标题**</span><span class="sxs-lookup"><span data-stu-id="50791-283">**Title**</span></span> <br/>  <span data-ttu-id="50791-p138">必需。supertip 的文本。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 ShortStrings 元素的子元素，而  元素是“Resources”元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p138">Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="50791-288">**说明**</span><span class="sxs-lookup"><span data-stu-id="50791-288">**Description**</span></span> <br/>  <span data-ttu-id="50791-p139">必需。supertip 的说明。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 LongStrings 元素的子元素，而  元素是“Resources”元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p139">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-293">**Icon**</span><span class="sxs-lookup"><span data-stu-id="50791-293">**Icon**</span></span> <br/> | <span data-ttu-id="50791-p140">必需。包含按钮的 **Image** 元素。图像文件必须为 .png 格式。 </span><span class="sxs-lookup"><span data-stu-id="50791-p140">Required. Contains the **Image** elements for the button. Image files must be .png format. </span></span><br/> <span data-ttu-id="50791-297">**Image**</span><span class="sxs-lookup"><span data-stu-id="50791-297">**Image**</span></span> <br/>  <span data-ttu-id="50791-p141">定义按钮上要显示的图像。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性指示图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。 </span><span class="sxs-lookup"><span data-stu-id="50791-p141">Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="50791-304">**操作**</span><span class="sxs-lookup"><span data-stu-id="50791-304">**Action**</span></span> <br/> | <span data-ttu-id="50791-p142">必需。指定用户选择按钮时将执行的操作。可以为 **xsi:type** 属性指定下列任意值之一： </span><span class="sxs-lookup"><span data-stu-id="50791-p142">Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: </span></span><br/> <span data-ttu-id="50791-p143">**ExecuteFunction**，它运行位于 **FunctionFile** 引用的文件中的 JavaScript 函数。**ExecuteFunction** 不显示 UI。**FunctionName** 子元素指定要执行的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="50791-p143">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**. **ExecuteFunction** does not display a UI. The **FunctionName** child element specifies the name of the function to execute. </span></span><br/> <span data-ttu-id="50791-p144">**ShowTaskPane**，它显示任务窗格外接程序。**SourceLocation** 子元素指定要显示的任务窗格外接程序的源文件位置。**resid** 属性必须设置为 **Resources** 元素的 **Urls** 元素中 **Url** 元素的 **id** 属性的值。 </span><span class="sxs-lookup"><span data-stu-id="50791-p144">**ShowTaskPane**, which shows a task pane add-in. The **SourceLocation** child element specifies the source file location of the task pane add-in to display. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element. </span></span><br/> |

### <a name="menu-controls"></a><span data-ttu-id="50791-314">菜单控件</span><span class="sxs-lookup"><span data-stu-id="50791-314">Menu controls</span></span>

<span data-ttu-id="50791-315">**Menu** 控件可与 **PrimaryCommandSurface** 或 **ContextMenu** 结合使用，并定义：</span><span class="sxs-lookup"><span data-stu-id="50791-315">A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:</span></span>
  
- <span data-ttu-id="50791-316">根级别菜单项。</span><span class="sxs-lookup"><span data-stu-id="50791-316">A root-level menu item.</span></span>
- <span data-ttu-id="50791-317">子菜单项的列表。</span><span class="sxs-lookup"><span data-stu-id="50791-317">A list of submenu items.</span></span>

<span data-ttu-id="50791-p145">与 **PrimaryCommandSurface** 结合使用时，根菜单项显示为功能区上的一个按钮。选择此按钮时，子菜单显示为下拉列表。与 **ContextMenu** 结合使用时，将在上下文菜单上插入包含子菜单的菜单项。在这两种情况中，单个子菜单项均可以执行 JavaScript 函数或显示任务窗格。目前只支持一种子菜单级别。</span><span class="sxs-lookup"><span data-stu-id="50791-p145">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="50791-p146">下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。在 **Control** 元素中：</span><span class="sxs-lookup"><span data-stu-id="50791-p146">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **Control** element:</span></span>

- <span data-ttu-id="50791-326">**xsi:type** 属性是必需的，并且必须设置为 **Menu**。</span><span class="sxs-lookup"><span data-stu-id="50791-326">The **xsi:type** attribute is required, and must be set to **Menu**.</span></span>
- <span data-ttu-id="50791-327">**id** 属性是一个最多为 125 个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="50791-327">The **id** attribute is a string with a maximum of 125 characters.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|<span data-ttu-id="50791-328">元素</span><span class="sxs-lookup"><span data-stu-id="50791-328">Elements</span></span>|<span data-ttu-id="50791-329">说明</span><span class="sxs-lookup"><span data-stu-id="50791-329">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-330">**Label**</span><span class="sxs-lookup"><span data-stu-id="50791-330">**Label**</span></span> <br/> |<span data-ttu-id="50791-p147">必需。根菜单项的文本。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p147">Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-335">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="50791-335">**Tooltip**</span></span> <br/> |<span data-ttu-id="50791-p148">可选。菜单的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p148">Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-340">**SuperTip**</span><span class="sxs-lookup"><span data-stu-id="50791-340">**SuperTip**</span></span> <br/> | <span data-ttu-id="50791-p149">必需。菜单的 SuperTip，定义如下： </span><span class="sxs-lookup"><span data-stu-id="50791-p149">Required. The supertip for the menu, which is defined by the following: </span></span><br/> <span data-ttu-id="50791-343">**标题**</span><span class="sxs-lookup"><span data-stu-id="50791-343">**Title**</span></span> <br/>  <span data-ttu-id="50791-p150">必需。supertip 的文本。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 ShortStrings 元素的子元素，而  元素是“Resources”元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p150">Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="50791-348">**说明**</span><span class="sxs-lookup"><span data-stu-id="50791-348">**Description**</span></span> <br/>  <span data-ttu-id="50791-p151">必需。supertip 的说明。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 LongStrings 元素的子元素，而  元素是“Resources”元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="50791-p151">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="50791-353">**Icon**</span><span class="sxs-lookup"><span data-stu-id="50791-353">**Icon**</span></span> <br/> | <span data-ttu-id="50791-p152">必需。包含菜单的 **Image** 元素。图像文件必须为 .png 格式。 </span><span class="sxs-lookup"><span data-stu-id="50791-p152">Required. Contains the **Image** elements for the menu. Image files must be .png format. </span></span><br/> <span data-ttu-id="50791-357">**Image**</span><span class="sxs-lookup"><span data-stu-id="50791-357">**Image**</span></span> <br/>  <span data-ttu-id="50791-p153">菜单的图像。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性指示图像的大小（以像素为单位）。要求三种图像大小（以像素为单位）：16、32 和 80。也同样支持五种可选大小（以像素为单位）：20、24、40、48 和 64。 </span><span class="sxs-lookup"><span data-stu-id="50791-p153">An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="50791-364">**Items**</span><span class="sxs-lookup"><span data-stu-id="50791-364">**Items**</span></span> <br/> |<span data-ttu-id="50791-p154">必需。包含每个子菜单项的 **Item** 元素。每个 **Item** 元素包含的子元素均与 [按钮控件](../reference/manifest/control.md#button-control)相同。  </span><span class="sxs-lookup"><span data-stu-id="50791-p154">Required. Contains the **Item** elements for each submenu item. Each **Item** element contains the same child elements as [Button controls](../reference/manifest/control.md#button-control).  </span></span><br/> |

## <a name="step-7-add-the-resources-element"></a><span data-ttu-id="50791-368">步骤 7：添加 Resources 元素</span><span class="sxs-lookup"><span data-stu-id="50791-368">Step 7: Add the Resources element</span></span>

<span data-ttu-id="50791-p155">**Resources** 元素包含 **VersionOverrides** 元素的不同子元素所使用的资源。这些资源包括图标、字符串和 URL。清单中的元素可以通过引用资源的 **id** 来使用此资源。使用 **id** 有助于使清单保持有序状态，尤其是当多个区域设置拥有不同的资源版本时。一个 **id** 最多可包含 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="50791-p155">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.</span></span>
  
<span data-ttu-id="50791-p156">以下示例演示了如何使用 **Resources** 元素。每个资源可以具有一个或多个 **Override** 子元素以定义特定区域设置的不同资源。</span><span class="sxs-lookup"><span data-stu-id="50791-p156">The following shows an example of how to use the **Resources** element. Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|<span data-ttu-id="50791-376">资源</span><span class="sxs-lookup"><span data-stu-id="50791-376">Resource</span></span>|<span data-ttu-id="50791-377">说明</span><span class="sxs-lookup"><span data-stu-id="50791-377">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-378">**Images**/ **Image**</span><span class="sxs-lookup"><span data-stu-id="50791-378">**Images**/ **Image**</span></span> <br/> | <span data-ttu-id="50791-p157">提供图像文件的 HTTPS URL。每个图像必须定义三个必需的图像大小：</span><span class="sxs-lookup"><span data-stu-id="50791-p157">Provides the HTTPS URL to an image file. Each image must define the three required image sizes:</span></span> <br/>  <span data-ttu-id="50791-381">16×16</span><span class="sxs-lookup"><span data-stu-id="50791-381">16×16</span></span> <br/>  <span data-ttu-id="50791-382">32×32</span><span class="sxs-lookup"><span data-stu-id="50791-382">32×32</span></span> <br/>  <span data-ttu-id="50791-383">80×80</span><span class="sxs-lookup"><span data-stu-id="50791-383">80×80</span></span> <br/>  <span data-ttu-id="50791-384">也支持下面的图像大小，但不是必需：</span><span class="sxs-lookup"><span data-stu-id="50791-384">The following image sizes are also supported, but not required:</span></span> <br/>  <span data-ttu-id="50791-385">20×20</span><span class="sxs-lookup"><span data-stu-id="50791-385">20×20</span></span> <br/>  <span data-ttu-id="50791-386">24×24</span><span class="sxs-lookup"><span data-stu-id="50791-386">24×24</span></span> <br/>  <span data-ttu-id="50791-387">40×40</span><span class="sxs-lookup"><span data-stu-id="50791-387">40×40</span></span> <br/>  <span data-ttu-id="50791-388">48×48</span><span class="sxs-lookup"><span data-stu-id="50791-388">48×48</span></span> <br/>  <span data-ttu-id="50791-389">64×64</span><span class="sxs-lookup"><span data-stu-id="50791-389">64×64</span></span> <br/> |
|<span data-ttu-id="50791-390">**Urls**/ **Url**</span><span class="sxs-lookup"><span data-stu-id="50791-390">**Urls**/ **Url**</span></span> <br/> |<span data-ttu-id="50791-p158">提供 HTTPS URL 位置。URL 最多可为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="50791-p158">Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.</span></span>  <br/> |
|<span data-ttu-id="50791-393">**ShortStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="50791-393">**ShortStrings**/ **String**</span></span> <br/> |<span data-ttu-id="50791-p159">**Label** 和 **Title** 元素的文本。每个 **String** 最多可包含 125 个字符。 </span><span class="sxs-lookup"><span data-stu-id="50791-p159">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="50791-396">**LongStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="50791-396">**LongStrings**/ **String**</span></span> <br/> |<span data-ttu-id="50791-p160">**Tooltip** 和 **Description** 元素的文本。每个 **String** 最多可包含 250 个字符。</span><span class="sxs-lookup"><span data-stu-id="50791-p160">The text for **Tooltip** and **Description** elements. Each **String** contains a maximum of 250 characters. </span></span><br/> |

> [!NOTE]
> <span data-ttu-id="50791-399">必须对 **Image** 和 **Url** 元素中的所有 URL 使用安全套接字层 (SSL)。</span><span class="sxs-lookup"><span data-stu-id="50791-399">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a><span data-ttu-id="50791-400">默认 Office 应用功能区选项卡的 Tab 值</span><span class="sxs-lookup"><span data-stu-id="50791-400">Tab values for default Office app ribbon tabs</span></span>

<span data-ttu-id="50791-p161">在 Excel 和 Word 中，可以使用默认 Office UI 选项卡，在功能区上添加加载项命令。下表列出了可用于 **OfficeTab** 元素的 **id** 属性的值。这些 Tab 值区分大小写。</span><span class="sxs-lookup"><span data-stu-id="50791-p161">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element. The tab values are case sensitive.</span></span>

|<span data-ttu-id="50791-404">Office 客户端应用程序</span><span class="sxs-lookup"><span data-stu-id="50791-404">Office client application</span></span>|<span data-ttu-id="50791-405">Tab 值</span><span class="sxs-lookup"><span data-stu-id="50791-405">Tab values</span></span>|
|:-----|:-----|
|<span data-ttu-id="50791-406">Excel</span><span class="sxs-lookup"><span data-stu-id="50791-406">Excel</span></span>  <br/> |<span data-ttu-id="50791-407">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span><span class="sxs-lookup"><span data-stu-id="50791-407">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span></span> <br/> |
|<span data-ttu-id="50791-408">Word</span><span class="sxs-lookup"><span data-stu-id="50791-408">Word</span></span>  <br/> |<span data-ttu-id="50791-409">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span><span class="sxs-lookup"><span data-stu-id="50791-409">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span></span> <br/> |
|<span data-ttu-id="50791-410">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="50791-410">PowerPoint</span></span>  <br/> |<span data-ttu-id="50791-411">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span><span class="sxs-lookup"><span data-stu-id="50791-411">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span></span>          <br/> |

## <a name="see-also"></a><span data-ttu-id="50791-412">另请参阅</span><span class="sxs-lookup"><span data-stu-id="50791-412">See also</span></span>

- [<span data-ttu-id="50791-413">Excel、PowerPoint 和 Word 的加载项命令</span><span class="sxs-lookup"><span data-stu-id="50791-413">Add-in commands for Excel, PowerPoint, and Word</span></span>](../design/add-in-commands.md)
