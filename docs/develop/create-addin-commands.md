---
title: ?????? Excel?Word ? PowerPoint ??????
description: ?????? VersionOverrides ?? Excel?Word ? PowerPoint ?????? ?????????? UI ????????????????????????
ms.date: 12/04/2017
ms.openlocfilehash: 95861fe0de6f0f56f6436b98cd7ad8dee510e82d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a><span data-ttu-id="184dd-104">?????? Excel?Word ? PowerPoint ?????</span><span class="sxs-lookup"><span data-stu-id="184dd-104">Create add-in commands in your manifest for Excel, Word, and PowerPoint</span></span>


<span data-ttu-id="184dd-105">?????? **[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)** ?? Excel?Word ? PowerPoint ??????</span><span class="sxs-lookup"><span data-stu-id="184dd-105">Use **[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)** in your manifest to define add-in commands for Excel, Word, and PowerPoint.</span></span> <span data-ttu-id="184dd-106">????????????????? UI ????????? Office ???? (UI) ??????</span><span class="sxs-lookup"><span data-stu-id="184dd-106">Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions.</span></span> <span data-ttu-id="184dd-107">????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-107">You can use add-in commands to:</span></span>
- <span data-ttu-id="184dd-108">?? UI ?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-108">Create UI elements or entry points that make your add-in's functionality easier to use.</span></span>  
  
- <span data-ttu-id="184dd-109">?????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-109">Add buttons or a drop-down list of buttons to the ribbon.</span></span>    
  
- <span data-ttu-id="184dd-110">??????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-110">Add individual menu items ? each containing optional submenus ? to specific context (shortcut) menus.</span></span>    
  
- <span data-ttu-id="184dd-p103">????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p103">Perform actions when your add-in command is chosen. You can:</span></span>
    
  - <span data-ttu-id="184dd-p104">??????????????????????????????????????????? Office UI ??????? UI ? HTML?</span><span class="sxs-lookup"><span data-stu-id="184dd-p104">Show one or more task pane add-ins for users to interact with. Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span></span>
    
     <span data-ttu-id="184dd-115">*??*</span><span class="sxs-lookup"><span data-stu-id="184dd-115">*or*</span></span> 
      
  - <span data-ttu-id="184dd-116">?? JavaScript ?????????????? UI ???????</span><span class="sxs-lookup"><span data-stu-id="184dd-116">Run JavaScript code, which normally runs without displaying any UI.</span></span>
      
<span data-ttu-id="184dd-p105">??????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p105">This article describes how to edit your manifest to define add-in commands. The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.</span></span> 
      
<span data-ttu-id="184dd-120">???????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-120">The following image is an overview of add-in commands elements in the manifest.</span></span> 
<span data-ttu-id="184dd-121">![?????????????](../images/version-overrides.png)</span><span class="sxs-lookup"><span data-stu-id="184dd-121">![Overview of add-in commands elements in the manifest](../images/version-overrides.png)</span></span>
 
## <a name="step-1-start-from-a-sample"></a><span data-ttu-id="184dd-122">? 1 ???????</span><span class="sxs-lookup"><span data-stu-id="184dd-122">Step 1: Start from a sample</span></span>

<span data-ttu-id="184dd-p107">????? [Office ???????](https://github.com/OfficeDev/Office-Add-in-Command-Sample)?????????????????????????????????????Office ???????????? XSD ??????????????????????? [Excel?Word ? PowerPoint ?????](../design/add-in-commands.md)?</span><span class="sxs-lookup"><span data-stu-id="184dd-p107">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Optionally, you can create your own manifest by following the steps in this guide. You can validate your manifest using the XSD file in the Office Add-in Commands Samples site. Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span></span>

## <a name="step-2-create-a-task-pane-add-in"></a><span data-ttu-id="184dd-127">? 2 ???????????</span><span class="sxs-lookup"><span data-stu-id="184dd-127">Step 2: Create a task pane add-in</span></span>

<span data-ttu-id="184dd-p108">??????????????????????????????????????????????????????????????????????? **XML ????**? **VersionOverrides** ??????????[? 3 ???? VersionOverrides ??](#step-3-add-versionoverrides-element)????</span><span class="sxs-lookup"><span data-stu-id="184dd-p108">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article. You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span></span>
   
<span data-ttu-id="184dd-p109">??????? Office 2013 ??????????????????????????? **VersionOverrides** ???Office 2013 ??????????????? **VersionOverrides** ??????????????? Office 2013 ? Office 2016 ????? Office 2013 ????????????????????? **SourceLocation** ?????????????????????? Office 2016 ??????? **VersionOverrides** ?????? **SourceLocation** ??????????????? **VersionOverrides**???????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p109">The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **VersionOverrides** element. Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in. In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in. If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span></span>
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
 
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

## <a name="step-3-add-versionoverrides-element"></a><span data-ttu-id="184dd-136">?? 3??? VersionOverrides ??</span><span class="sxs-lookup"><span data-stu-id="184dd-136">Step 3: Add VersionOverrides element</span></span>
<span data-ttu-id="184dd-p110">**VersionOverrides** ??????????????????**VersionOverrides** ???? **OfficeApp** ???????????? **VersionOverrides** ??????</span><span class="sxs-lookup"><span data-stu-id="184dd-p110">The **VersionOverrides** element is the root element that contains the definition of your add-in command. **VersionOverrides** is a child element of the **OfficeApp** element in the manifest. The following table lists the attributes of the **VersionOverrides** element.</span></span>

|<span data-ttu-id="184dd-140">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-140">**Attribute**</span></span>|<span data-ttu-id="184dd-141">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-141">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-142">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="184dd-142">**xmlns**</span></span> <br/> | <span data-ttu-id="184dd-p111">????????????http://schemas.microsoft.com/office/taskpaneappversionoverrides??</span><span class="sxs-lookup"><span data-stu-id="184dd-p111">Required. The schema location, which must be "http://schemas.microsoft.com/office/taskpaneappversionoverrides".</span></span> <br/> |
|<span data-ttu-id="184dd-145">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="184dd-145">**xsi:type**</span></span> <br/> |<span data-ttu-id="184dd-p112">?????????????????"VersionOverridesV1_0"?</span><span class="sxs-lookup"><span data-stu-id="184dd-p112">Required. The schema version. The version described in this article is "VersionOverridesV1_0".</span></span>  <br/> |
   
<span data-ttu-id="184dd-149">????? **VersionOverrides** ?????</span><span class="sxs-lookup"><span data-stu-id="184dd-149">The following table identifies the child elements of **VersionOverrides**.</span></span>
  
|<span data-ttu-id="184dd-150">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-150">**Element**</span></span>|<span data-ttu-id="184dd-151">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-151">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-152">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-152">**Description**</span></span> <br/> |<span data-ttu-id="184dd-p113">????????????? **Description** ?????????????? **Description** ???? **Description** ??? **resid** ?????? **String** ??? **id**?**String** ???? **Description** ???? </span><span class="sxs-lookup"><span data-stu-id="184dd-p113">Optional. Describes the add-in. This child **Description** element overrides a previous **Description** element in the parent portion of the manifest. The **resid** attribute for this **Description** element is set to the **id** of a **String** element. The **String** element contains the text for **Description**. </span></span><br/> |
|<span data-ttu-id="184dd-158">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="184dd-158">**Requirements**</span></span> <br/> |<span data-ttu-id="184dd-p114">?????????????????? Office.js ??????? **Requirements** ????????????? **Requirements** ?????????????[?? Office ??? API ??](../develop/specify-office-hosts-and-api-requirements.md)?  </span><span class="sxs-lookup"><span data-stu-id="184dd-p114">Optional. Specifies the minimum requirement set and version of Office.js that the add-in requires. This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest. For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).  </span></span><br/> |
|<span data-ttu-id="184dd-163">**Hosts**</span><span class="sxs-lookup"><span data-stu-id="184dd-163">**Hosts**</span></span> <br/> |<span data-ttu-id="184dd-p115">????? Office ???????? **Hosts** ????????????? **Hosts** ????????????Workbook???Document?? **xsi:type** ?? </span><span class="sxs-lookup"><span data-stu-id="184dd-p115">Required. Specifies a collection of Office hosts. The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest. You must include a **xsi:type** attribute set to "Workbook" or "Document". </span></span><br/> |
|<span data-ttu-id="184dd-168">**Resources**</span><span class="sxs-lookup"><span data-stu-id="184dd-168">**Resources**</span></span> <br/> |<span data-ttu-id="184dd-p116">????????????????????URL ????????**Description** ??????? **Resources** ??????**Resources** ????????????[?? 7??? Resources ??](#step-7-add-the-resources-element)?????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p116">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **Description** element's value refers to a child element in **Resources**. The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. </span></span><br/> |
   
<span data-ttu-id="184dd-172">??????????? **VersionOverrides** ????????</span><span class="sxs-lookup"><span data-stu-id="184dd-172">The following example shows how to use the **VersionOverrides** element and its child elements.</span></span>

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

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a><span data-ttu-id="184dd-173">?? 4??? Hosts?Host ? DesktopFormFactor ??</span><span class="sxs-lookup"><span data-stu-id="184dd-173">Step 4: Add Hosts, Host, and DesktopFormFactor elements</span></span>

<span data-ttu-id="184dd-p117">**Hosts** ????????? **Host** ????? **Host** ????????? Office ???**Host** ????????????????????? Office ???????????????????????????????? Office ???????????????????? **Host** ??????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p117">The **Hosts** element contains one or more **Host** elements. A **Host** element specifies a particular Office host. The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host. To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.</span></span>
       
<span data-ttu-id="184dd-178">**DesktopFormFactor** ??????? Windows ??? Office ????? Office Online???????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-178">The **DesktopFormFactor** element specifies the settings for an add-in that runs in Office on Windows desktop, and Office Online (in browser).</span></span>
      
<span data-ttu-id="184dd-179">??????? **Hosts**?**Host** ? **DesktopFormFactor** ??????</span><span class="sxs-lookup"><span data-stu-id="184dd-179">The following is an example of **Hosts**, **Host**, and **DesktopFormFactor** elements.</span></span>

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

## <a name="step-5-add-the-functionfile-element"></a><span data-ttu-id="184dd-180">?? 5??? FunctionFile ??</span><span class="sxs-lookup"><span data-stu-id="184dd-180">Step 5: Add the FunctionFile element</span></span>

<span data-ttu-id="184dd-p118">"FunctionFile"???????????????????????"ExecuteFunction"??????? JavaScript ?????? ?????????????"FunctionFile"???"resid"?????????????????? JavaScript ??? HTML ????????? JavaScript ???????????"Resources"????"Url"???********[](https://dev.office.com/reference/add-ins/manifest/control#Button-control)****************</span><span class="sxs-lookup"><span data-stu-id="184dd-p118">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](https://dev.office.com/reference/add-ins/manifest/control#Button-control) for a description). The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **Url** element in the **Resources** element.</span></span>
        
<span data-ttu-id="184dd-186">???????? **FunctionFile** ???</span><span class="sxs-lookup"><span data-stu-id="184dd-186">The following is an example of the **FunctionFile** element.</span></span>
  
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
> <span data-ttu-id="184dd-187">??? JavaScript ????? `Office.initialize`?</span><span class="sxs-lookup"><span data-stu-id="184dd-187">Make sure your JavaScript code calls  `Office.initialize`.</span></span> 
   
<span data-ttu-id="184dd-p119">**FunctionFile** ????? HTML ???? JavaScript ???? `Office.initialize`?**FunctionName** ??????[????](https://dev.office.com/reference/add-ins/manifest/control#Button-control)????????? **FunctionFile** ?????</span><span class="sxs-lookup"><span data-stu-id="184dd-p119">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`. The **FunctionName** element (see [Button controls](https://dev.office.com/reference/add-ins/manifest/control#Button-control) for a description) uses the functions in **FunctionFile**.</span></span>
     
<span data-ttu-id="184dd-190">???????????? **FunctionName** ??????</span><span class="sxs-lookup"><span data-stu-id="184dd-190">The following code shows how to implement the function used by **FunctionName**.</span></span>

```javascript

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
                if (asyncResult.status === "failed") {
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
> <span data-ttu-id="184dd-p120">?? **event.completed** ?????????????????????????????????????????????????????????????????????????????? **event.completed**?????????????????????? **event.completed**??????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p120">The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.</span></span>
 
## <a name="step-6-add-extensionpoint-elements"></a><span data-ttu-id="184dd-196">? 6 ???? ExtensionPoint ??</span><span class="sxs-lookup"><span data-stu-id="184dd-196">Step 6: Add ExtensionPoint elements</span></span>

<span data-ttu-id="184dd-p121">**ExtensionPoint** ???????????? Office UI ??????????????? **xsi:type** ??? **ExtensionPoint** ???</span><span class="sxs-lookup"><span data-stu-id="184dd-p121">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI. You can define **ExtensionPoint** elements with these **xsi:type** values:</span></span>
   
- <span data-ttu-id="184dd-199">**PrimaryCommandSurface**???? Office ??????</span><span class="sxs-lookup"><span data-stu-id="184dd-199">**PrimaryCommandSurface**, which refers to the ribbon in Office.</span></span>
     
- <span data-ttu-id="184dd-200">**ContextMenu**?????? Office UI ??????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-200">**ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.</span></span>
    
<span data-ttu-id="184dd-201">?????????? **ExtensionPoint** ??? **PrimaryCommandSurface** ? **ContextMenu** ??????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-201">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>
    
> [!IMPORTANT]
> <span data-ttu-id="184dd-p122">???? ID ????????????? ID????????? ID ????????????????`<CustomTab id="mycompanyname.mygroupname">`?</span><span class="sxs-lookup"><span data-stu-id="184dd-p122">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span></span> 
  
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

|<span data-ttu-id="184dd-205">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-205">**Element**</span></span>|<span data-ttu-id="184dd-206">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-206">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-207">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="184dd-207">**CustomTab**</span></span> <br/> |<span data-ttu-id="184dd-p123">??????? **PrimaryCommandSurface**???????????????????????? **CustomTab** ???????? **OfficeTab** ???**id** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p123">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required. </span></span><br/> |
|<span data-ttu-id="184dd-211">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="184dd-211">**OfficeTab**</span></span> <br/> |<span data-ttu-id="184dd-p124">??????? **PrimaryCommandSurface**????? Office ????????????????? **OfficeTab** ???????? **CustomTab** ??? </span><span class="sxs-lookup"><span data-stu-id="184dd-p124">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the **OfficeTab** element, you can't use the **CustomTab** element. </span></span><br/> <span data-ttu-id="184dd-214">??? **id** ????????? tab ?????[?? Office ??????? Tab ?](https://dev.office.com/reference/add-ins/manifest/officetab)?</span><span class="sxs-lookup"><span data-stu-id="184dd-214">For more tab values to use with the **id** attribute, see [Tab values for default Office ribbon tabs](https://dev.office.com/reference/add-ins/manifest/officetab).</span></span>  <br/> |
|<span data-ttu-id="184dd-215">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="184dd-215">**OfficeMenu**</span></span> <br/> | <span data-ttu-id="184dd-p125">?????? **ContextMenu**??????????????????????????**id** ???????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p125">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="184dd-p126">??????????????????????? Excel ? Word ? **ContextMenuText**??????????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p126">**ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="184dd-p127">??? Excel ? **ContextMenuCell**??????????????????????????????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p127">**ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. </span></span><br/> |
|<span data-ttu-id="184dd-222">**Group**</span><span class="sxs-lookup"><span data-stu-id="184dd-222">**Group**</span></span> <br/> |<span data-ttu-id="184dd-p128">???????????????????????????**id** ?????????????? 125 ???????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p128">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="184dd-226">**Label**</span><span class="sxs-lookup"><span data-stu-id="184dd-226">**Label**</span></span> <br/> |<span data-ttu-id="184dd-p129">???????**resid** ??????? **String** ??? **id** ?????**String** ??? **ShortStrings** ???????? ShortStrings ??? **Resources** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p129">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-231">**Icon**</span><span class="sxs-lookup"><span data-stu-id="184dd-231">**Icon**</span></span> <br/> |<span data-ttu-id="184dd-p130">?????????????????????????????????**resid** ??????? **Image** ??? **id** ?????**Image** ??? **Images** ???????? Images ??? **Resources** ???????**size** ???????????????????????????16?32 ? 80?????????????20?24?40?48 ? 64? </span><span class="sxs-lookup"><span data-stu-id="184dd-p130">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="184dd-239">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="184dd-239">**Tooltip**</span></span> <br/> |<span data-ttu-id="184dd-p131">?????????**resid** ??????? **String** ??? **id** ?????**String** ??? **LongStrings** ???????? LongStrings ??? **Resources** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p131">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-244">**Control**</span><span class="sxs-lookup"><span data-stu-id="184dd-244">**Control**</span></span> <br/> |<span data-ttu-id="184dd-p132">??????????????**Control** ????? **Button**????? **Menu**??? **Menu** ???????????????????????????[????](https://dev.office.com/reference/add-ins/manifest/control)?[????](https://dev.office.com/reference/add-ins/manifest/control)?????????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p132">Each group requires at least one control. A **Control** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the  [Button controls](https://dev.office.com/reference/add-ins/manifest/control) and [Menu controls](https://dev.office.com/reference/add-ins/manifest/control) sections for more information. </span></span><br/><span data-ttu-id="184dd-250">**???**???????? **Control** ????? **Resources** ??????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-250">**Note:** To make troubleshooting easier, we recommend that you add a **Control** element and the related **Resources** child elements one at a time.</span></span>          |
   

### <a name="button-controls"></a><span data-ttu-id="184dd-251">????</span><span class="sxs-lookup"><span data-stu-id="184dd-251">Button controls</span></span>
<span data-ttu-id="184dd-p133">???????????????????????? JavaScript ??????????????????????????????????? UI ?????? JavaScript ???????????????? **Control** ????</span><span class="sxs-lookup"><span data-stu-id="184dd-p133">A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **Control** element:</span></span>        

- <span data-ttu-id="184dd-257">**type** ?????????????? **Button**?</span><span class="sxs-lookup"><span data-stu-id="184dd-257">The **type** attribute is required, and must be set to **Button**.</span></span>
    
- <span data-ttu-id="184dd-258">**Control** ??? **id** ???????? 125 ????????</span><span class="sxs-lookup"><span data-stu-id="184dd-258">The **id** attribute of the **Control** element is a string with a maximum of 125 characters.</span></span>
    
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

|<span data-ttu-id="184dd-259">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-259">**Elements**</span></span>|<span data-ttu-id="184dd-260">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-260">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-261">**Label**</span><span class="sxs-lookup"><span data-stu-id="184dd-261">**Label**</span></span> <br/> |<span data-ttu-id="184dd-p134">????????**resid** ??????? **String** ??? **id** ?????**String** ??? **ShortStrings** ???????? ShortStrings ??? **Resources** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p134">Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-266">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="184dd-266">**Tooltip**</span></span> <br/> |<span data-ttu-id="184dd-p135">???????????**resid** ??????? **String** ??? **id** ?????**String** ??? **LongStrings** ???????? LongStrings ??? **Resources** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p135">Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-271">**Supertip**</span><span class="sxs-lookup"><span data-stu-id="184dd-271">**Supertip**</span></span> <br/> | <span data-ttu-id="184dd-p136">??????? SuperTip?????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p136">Required. The supertip for this button, which is defined by the following: </span></span><br/> <span data-ttu-id="184dd-274">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-274">**Title**</span></span> <br/>  <span data-ttu-id="184dd-p137">???supertip ????????resid?????? String ??? id ????String ??? ShortStrings ????????  ????Resources???????? ************************</span><span class="sxs-lookup"><span data-stu-id="184dd-p137">Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="184dd-279">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-279">**Description**</span></span> <br/>  <span data-ttu-id="184dd-p138">???supertip ????????resid?????? String ??? id ????String ??? LongStrings ????????  ????Resources???????? ************************</span><span class="sxs-lookup"><span data-stu-id="184dd-p138">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-284">**Icon**</span><span class="sxs-lookup"><span data-stu-id="184dd-284">**Icon**</span></span> <br/> | <span data-ttu-id="184dd-p139">???????? **Image** ?????????? .png ??? </span><span class="sxs-lookup"><span data-stu-id="184dd-p139">Required. Contains the **Image** elements for the button. Image files must be .png format. </span></span><br/> <span data-ttu-id="184dd-288">**Image**</span><span class="sxs-lookup"><span data-stu-id="184dd-288">**Image**</span></span> <br/>  <span data-ttu-id="184dd-p140">????????????**resid** ??????? **Image** ??? **id** ?????**Image** ??? **Images** ???????? Images ??? **Resources** ???????**size** ???????????????????????????16?32 ? 80?????????????20?24?40?48 ? 64? </span><span class="sxs-lookup"><span data-stu-id="184dd-p140">Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="184dd-295">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-295">**Action**</span></span> <br/> | <span data-ttu-id="184dd-p141">?????????????????????? **xsi:type** ???????????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p141">Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: </span></span><br/> <span data-ttu-id="184dd-p142">**ExecuteFunction**?????? **FunctionFile** ??????? JavaScript ???**ExecuteFunction** ??? UI?**FunctionName** ??????????????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p142">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**. **ExecuteFunction** does not display a UI. The **FunctionName** child element specifies the name of the function to execute. </span></span><br/> <span data-ttu-id="184dd-p143">**ShowTaskPane**?????????????**SourceLocation** ????????????????????????**resid** ??????? **Resources** ??? **Urls** ??? **Url** ??? **id** ????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p143">**ShowTaskPane**, which shows a task pane add-in. The **SourceLocation** child element specifies the source file location of the task pane add-in to display. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element. </span></span><br/> |
   

### <a name="menu-controls"></a><span data-ttu-id="184dd-305">????</span><span class="sxs-lookup"><span data-stu-id="184dd-305">Menu controls</span></span>
<span data-ttu-id="184dd-306">**Menu** ???? **PrimaryCommandSurface** ? **ContextMenu** ?????????</span><span class="sxs-lookup"><span data-stu-id="184dd-306">A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:</span></span>
  
- <span data-ttu-id="184dd-307">???????</span><span class="sxs-lookup"><span data-stu-id="184dd-307">A root-level menu item.</span></span>
   
- <span data-ttu-id="184dd-308">????????</span><span class="sxs-lookup"><span data-stu-id="184dd-308">A list of submenu items.</span></span>
 
<span data-ttu-id="184dd-p144">? **PrimaryCommandSurface** ?????????????????????????????????????????? **ContextMenu** ????????????????????????????????????????????? JavaScript ???????????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p144">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>
       
<span data-ttu-id="184dd-p145">???????????????????????????????????????????????? JavaScript ???? **Control** ????</span><span class="sxs-lookup"><span data-stu-id="184dd-p145">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **Control** element:</span></span>
    
- <span data-ttu-id="184dd-317">**xsi:type** ?????????????? **Menu**?</span><span class="sxs-lookup"><span data-stu-id="184dd-317">The **xsi:type** attribute is required, and must be set to **Menu**.</span></span>
  
- <span data-ttu-id="184dd-318">**id** ???????? 125 ????????</span><span class="sxs-lookup"><span data-stu-id="184dd-318">The **id** attribute is a string with a maximum of 125 characters.</span></span>
    
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

|<span data-ttu-id="184dd-319">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-319">**Elements**</span></span>|<span data-ttu-id="184dd-320">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-320">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-321">**Label**</span><span class="sxs-lookup"><span data-stu-id="184dd-321">**Label**</span></span> <br/> |<span data-ttu-id="184dd-p146">???????????**resid** ??????? **String** ??? **id** ?????**String** ??? **ShortStrings** ???????? ShortStrings ??? **Resources** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p146">Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-326">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="184dd-326">**Tooltip**</span></span> <br/> |<span data-ttu-id="184dd-p147">???????????**resid** ??????? **String** ??? **id** ?????**String** ??? **LongStrings** ???????? LongStrings ??? **Resources** ??????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p147">Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-331">**SuperTip**</span><span class="sxs-lookup"><span data-stu-id="184dd-331">**SuperTip**</span></span> <br/> | <span data-ttu-id="184dd-p148">?????? SuperTip?????? </span><span class="sxs-lookup"><span data-stu-id="184dd-p148">Required. The supertip for the menu, which is defined by the following: </span></span><br/> <span data-ttu-id="184dd-334">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-334">**Title**</span></span> <br/>  <span data-ttu-id="184dd-p149">???supertip ????????resid?????? String ??? id ????String ??? ShortStrings ????????  ????Resources???????? ************************</span><span class="sxs-lookup"><span data-stu-id="184dd-p149">Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="184dd-339">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-339">**Description**</span></span> <br/>  <span data-ttu-id="184dd-p150">???supertip ????????resid?????? String ??? id ????String ??? LongStrings ????????  ????Resources???????? ************************</span><span class="sxs-lookup"><span data-stu-id="184dd-p150">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="184dd-344">**Icon**</span><span class="sxs-lookup"><span data-stu-id="184dd-344">**Icon**</span></span> <br/> | <span data-ttu-id="184dd-p151">???????? **Image** ?????????? .png ??? </span><span class="sxs-lookup"><span data-stu-id="184dd-p151">Required. Contains the **Image** elements for the menu. Image files must be .png format. </span></span><br/> <span data-ttu-id="184dd-348">**Image**</span><span class="sxs-lookup"><span data-stu-id="184dd-348">**Image**</span></span> <br/>  <span data-ttu-id="184dd-p152">??????**resid** ??????? **Image** ??? **id** ?????**Image** ??? **Images** ???????? Images ??? **Resources** ???????**size** ???????????????????????????????????16?32 ? 80?????????????????????20?24?40?48 ? 64? </span><span class="sxs-lookup"><span data-stu-id="184dd-p152">An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="184dd-355">**Items**</span><span class="sxs-lookup"><span data-stu-id="184dd-355">**Items**</span></span> <br/> |<span data-ttu-id="184dd-p153">???????????? **Item** ????? **Item** ??????????[????](https://dev.office.com/reference/add-ins/manifest/control)???  </span><span class="sxs-lookup"><span data-stu-id="184dd-p153">Required. Contains the **Item** elements for each submenu item. Each **Item** element contains the same child elements as [Button controls](https://dev.office.com/reference/add-ins/manifest/control).  </span></span><br/> |
   
## <a name="step-7-add-the-resources-element"></a><span data-ttu-id="184dd-359">?? 7??? Resources ??</span><span class="sxs-lookup"><span data-stu-id="184dd-359">Step 7: Add the Resources element</span></span>

<span data-ttu-id="184dd-p154">**Resources** ???? **VersionOverrides** ???????????????????????????? URL???????????????? **id** ????????? **id** ???????????????????????????????????? **id** ????? 32 ????</span><span class="sxs-lookup"><span data-stu-id="184dd-p154">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.</span></span>
  
    
    
<span data-ttu-id="184dd-p155">??????????? **Resources** ???????????????? **Override** ??????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p155">The following shows an example of how to use the **Resources** element. Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>


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

|<span data-ttu-id="184dd-367">**Resource**</span><span class="sxs-lookup"><span data-stu-id="184dd-367">**Resource**</span></span>|<span data-ttu-id="184dd-368">**??**</span><span class="sxs-lookup"><span data-stu-id="184dd-368">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-369">**Images**/ **Image**</span><span class="sxs-lookup"><span data-stu-id="184dd-369">**Images**/ **Image**</span></span> <br/> | <span data-ttu-id="184dd-p156">??????? HTTPS URL???????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-p156">Provides the HTTPS URL to an image file. Each image must define the three required image sizes:</span></span> <br/>  <span data-ttu-id="184dd-372">16?16</span><span class="sxs-lookup"><span data-stu-id="184dd-372">16?16</span></span> <br/>  <span data-ttu-id="184dd-373">32?32</span><span class="sxs-lookup"><span data-stu-id="184dd-373">32?32</span></span> <br/>  <span data-ttu-id="184dd-374">80?80</span><span class="sxs-lookup"><span data-stu-id="184dd-374">80?80</span></span> <br/>  <span data-ttu-id="184dd-375">?????????????????</span><span class="sxs-lookup"><span data-stu-id="184dd-375">The following image sizes are also supported, but not required:</span></span> <br/>  <span data-ttu-id="184dd-376">20?20</span><span class="sxs-lookup"><span data-stu-id="184dd-376">20?20</span></span> <br/>  <span data-ttu-id="184dd-377">24?24</span><span class="sxs-lookup"><span data-stu-id="184dd-377">24?24</span></span> <br/>  <span data-ttu-id="184dd-378">40?40</span><span class="sxs-lookup"><span data-stu-id="184dd-378">40?40</span></span> <br/>  <span data-ttu-id="184dd-379">48?48</span><span class="sxs-lookup"><span data-stu-id="184dd-379">48?48</span></span> <br/>  <span data-ttu-id="184dd-380">64?64</span><span class="sxs-lookup"><span data-stu-id="184dd-380">64?64</span></span> <br/> |
|<span data-ttu-id="184dd-381">**Urls**/ **Url**</span><span class="sxs-lookup"><span data-stu-id="184dd-381">**Urls**/ **Url**</span></span> <br/> |<span data-ttu-id="184dd-p157">?? HTTPS URL ???URL ???? 2048 ????</span><span class="sxs-lookup"><span data-stu-id="184dd-p157">Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.</span></span>  <br/> |
|<span data-ttu-id="184dd-384">**ShortStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="184dd-384">**ShortStrings**/ **String**</span></span> <br/> |<span data-ttu-id="184dd-p158">**Label** ? **Title** ???????? **String** ????? 125 ???? </span><span class="sxs-lookup"><span data-stu-id="184dd-p158">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="184dd-387">**LongStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="184dd-387">**LongStrings**/ **String**</span></span> <br/> |<span data-ttu-id="184dd-p159">**Tooltip** ? **Description** ???????? **String** ????? 250 ???? </span><span class="sxs-lookup"><span data-stu-id="184dd-p159">The text for **Tooltip** and **Description** elements. Each **String** contains a maximum of 250 characters. </span></span><br/> |
   
> [!NOTE] 
> <span data-ttu-id="184dd-390">??? **Image** ? **Url** ?????? URL ???????? (SSL)?</span><span class="sxs-lookup"><span data-stu-id="184dd-390">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="tab-values-for-default-office-ribbon-tabs"></a><span data-ttu-id="184dd-391">?? Office ??????? tab ?</span><span class="sxs-lookup"><span data-stu-id="184dd-391">Tab values for default Office ribbon tabs</span></span>
<span data-ttu-id="184dd-p160">? Excel ? Word ???????? Office UI ????????????????????????? **OfficeTab** ??? **id** ??????? Tab ???????</span><span class="sxs-lookup"><span data-stu-id="184dd-p160">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element. The tab values are case sensitive.</span></span>

|<span data-ttu-id="184dd-395">**Office ????**</span><span class="sxs-lookup"><span data-stu-id="184dd-395">**Office host application**</span></span>|<span data-ttu-id="184dd-396">**Tab ?**</span><span class="sxs-lookup"><span data-stu-id="184dd-396">**Tab values**</span></span>|
|:-----|:-----|
|<span data-ttu-id="184dd-397">Excel</span><span class="sxs-lookup"><span data-stu-id="184dd-397">Excel</span></span>  <br/> |<span data-ttu-id="184dd-398">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span><span class="sxs-lookup"><span data-stu-id="184dd-398">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span></span> <br/> |
|<span data-ttu-id="184dd-399">Word</span><span class="sxs-lookup"><span data-stu-id="184dd-399">Word</span></span>  <br/> |<span data-ttu-id="184dd-400">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span><span class="sxs-lookup"><span data-stu-id="184dd-400">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span></span> <br/> |
|<span data-ttu-id="184dd-401">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="184dd-401">PowerPoint</span></span>  <br/> |<span data-ttu-id="184dd-402">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span><span class="sxs-lookup"><span data-stu-id="184dd-402">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span></span>          <br/> |
   
## <a name="see-also"></a><span data-ttu-id="184dd-403">????</span><span class="sxs-lookup"><span data-stu-id="184dd-403">See also</span></span>

-  [<span data-ttu-id="184dd-404">Excel?Word ? PowerPoint ?????</span><span class="sxs-lookup"><span data-stu-id="184dd-404">Add-in commands for Excel, Word and PowerPoint</span></span>](../design/add-in-commands.md)      
