---
title: 在清单中创建 Excel、Word 和 PowerPoint 的加载项命令
description: 在清单中使用 VersionOverrides 定义 Excel、Word 和 PowerPoint 加载项命令。 加载项命令可用于创建 UI 元素，也可用于添加按钮或列表，同时还能执行操作。
ms.date: 12/04/2017
ms.openlocfilehash: 4d0bb5eb82ef931c94e6791aaeab598af9f0e298
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005027"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a>在清单中创建 Excel、Word 和 PowerPoint 加载项命令


在清单中使用 **[VersionOverrides](https://docs.microsoft.com/javascript/office/manifest/versionoverrides?view=office-js)** 定义 Excel、Word 和 PowerPoint 加载项命令。 加载项命令提供了使用执行操作的特定 UI 元素来自定义默认的 Office 用户界面 (UI) 的简单方法。 可以使用加载项命令执行以下操作：
- 创建 UI 元素或入口点，以便能够更易于使用你的外接程序功能。  
  
- 向功能区中添加按钮或下拉列表按钮。    
  
- 将单个菜单项（每一个都包含可选的子菜单）添加到特定上下文（快捷方式）菜单中。    
  
- 在选择你的外接程序命令时执行操作。可以：
    
  - 显示一个或多个任务窗格外接程序，让用户与其进行交互。在任务窗格外接程序内，可以显示使用 Office UI 结构创建自定义 UI 的 HTML。
    
     *或者* 
      
  - 运行 JavaScript 代码，该代码通常在不显示任何 UI 的情况下运行。
      
本文介绍如何编辑您的清单来定义外接程序命令。下图显示了用来定义外接程序命令的元素的层次结构。本文将具体介绍这些元素。 
      
下图是对清单中的加载项命令元素的概述。 
![清单中的加载项命令元素概述](../images/version-overrides.png)
 
## <a name="step-1-start-from-a-sample"></a>第 1 步：从示例入手

强烈建议从 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Command-Sample)中的示例之一入手。也可以按照本指南中的步骤操作，创建自己的清单。可以使用“Office 加载项命令示例”网站中的 XSD 文件来验证清单。使用加载项命令前，请确保已阅读 [Excel、Word 和 PowerPoint 加载项命令](../design/add-in-commands.md)。

## <a name="step-2-create-a-task-pane-add-in"></a>第 2 步：创建任务窗格加载项

若要开始使用加载项命令，必须先创建任务窗格加载项，再按照本文所述来修改加载项清单。无法对内容加载项使用加载项命令。若要更新现有清单，必须将相应 **XML 命名空间**和 **VersionOverrides** 元素添加到清单中（如[第 3 步：添加 VersionOverrides 元素](#step-3-add-versionoverrides-element)所述）。
   
以下示例显示了 Office 2013 外接程序的清单。此清单中没有任何外接程序命令，因为没有 **VersionOverrides** 元素。Office 2013 不支持外接程序命令，但是通过将 **VersionOverrides** 添加到此清单，外接程序可同时在 Office 2013 和 Office 2016 中运行。在 Office 2013 中，外接程序不会显示外接程序命令，并且使用 **SourceLocation** 的值运行外接程序作为单一任务窗格外接程序。在 Office 2016 中，如果未包含 **VersionOverrides** 元素，则使用 **SourceLocation** 运行外接程序。但是，如果包含了 **VersionOverrides**，外接程序将只显示外接程序命令，并且不会将外接程序显示为单一任务窗格外接程序。
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
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

## <a name="step-3-add-versionoverrides-element"></a>步骤 3：添加 VersionOverrides 元素
**VersionOverrides** 元素是包含外接程序命令定义的根元素。**VersionOverrides** 是清单中 **OfficeApp** 元素的子元素。下表列出了 **VersionOverrides** 元素的属性。

|**属性**|**说明**|
|:-----|:-----|
|**xmlns** <br/> | 必需。架构位置，必须是“http://schemas.microsoft.com/office/taskpaneappversionoverrides”。 <br/> |
|**xsi:type** <br/> |必需。架构版本。本文中所述的版本为"VersionOverridesV1_0"。  <br/> |
   
下表标识了 **VersionOverrides** 的子元素。
  
|**元素**|**说明**|
|:-----|:-----|
|**说明** <br/> |可选。描述外接程序。此子级 **Description** 元素替代清单中父级部分中的旧 **Description** 元素。此 **Description** 元素的 **resid** 属性将设置为 **String** 元素的 **id**。**String** 元素包含 **Description** 的文本。 <br/> |
|**Requirements** <br/> |可选。指定外接程序要求的最低要求集和 Office.js 的版本。此子级 **Requirements** 元素替代清单中父级部分中的 **Requirements** 元素。有关详细信息，请参阅[指定 Office 主机和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。  <br/> |
|**Hosts** <br/> |必需。指定 Office 主机的集合。子级 **Hosts** 元素替代清单中父级部分中的 **Hosts** 元素。必须包含已设置为“Workbook”或“Document”的 **xsi:type** 属性 <br/> |
|**Resources** <br/> |定义其他清单元素引用的资源集合（字符串、URL 和图像）。例如，**Description** 元素的值引用了 **Resources** 中的子元素。**Resources** 元素将在本文后续部分中的[步骤 7：添加 Resources 元素](#step-7-add-the-resources-element)中进行介绍。 <br/> |
   
下面的示例演示如何使用 **VersionOverrides** 元素及其子元素。

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

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a>步骤 4：添加 Hosts、Host 和 DesktopFormFactor 元素

**Hosts** 元素包含一个或多个 **Host** 元素。一个 **Host** 元素指定一个特定的 Office 主机。**Host** 元素包含子元素，这些子元素用于指定在对应的 Office 主机安装外接程序后要显示的外接程序命令。若要在两个或更多个不同的 Office 主机中显示相同的外接程序命令，必须在每个 **Host** 中使用相同的子元素。
       
**DesktopFormFactor** 元素指定运行在 Windows 桌面的 Office 上和运行在 Office Online（在浏览器中）中的外接程序的相关设置。
      
以下是一个包含 **Hosts**、**Host** 和 **DesktopFormFactor** 元素的示例。

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

## <a name="step-5-add-the-functionfile-element"></a>步骤 5：添加 FunctionFile 元素

"FunctionFile"元素指定了一个文件，其中包含当外接程序命令使用"ExecuteFunction"操作时要运行的 JavaScript 代码（请参阅 按钮控件了解相关说明）。将"FunctionFile"元素的"resid"属性设置为包括外接程序命令需要的所有 JavaScript 文件的 HTML 文件。不能只链接到 JavaScript 文件。将文件名称指定为"Resources"元素中的"Url"元素。********[](https://docs.microsoft.com/javascript/office/manifest/control?view=office-js#Button-control)****************
        
下面的示例展示了 **FunctionFile** 元素。
  
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
> 请确保 JavaScript 代码调用了 `Office.initialize`。 
   
**FunctionFile** 元素引用的 HTML 文件中的 JavaScript 必须调用 `Office.initialize`。**FunctionName** 元素（请参阅[按钮控件](https://docs.microsoft.com/javascript/office/manifest/control?view=office-js#Button-control)查看相关说明）使用 **FunctionFile** 中的函数。
     
下面的代码展示了如何实现 **FunctionName** 使用的函数。

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
> 调用 **event.completed** 表示已成功处理事件。如果函数获得多次调用（如多次单击同一加载项命令），所有事件都会自动排入队列。首个事件会自动运行，而其他事件则继续留在队列中。如果函数调用 **event.completed**，将运行此函数在队列中的下一个调用。必须实现 **event.completed**，否则函数不会运行。
 
## <a name="step-6-add-extensionpoint-elements"></a>第 6 步：添加 ExtensionPoint 元素

**ExtensionPoint** 元素定义外接程序命令应在 Office UI 中的哪个位置出现。可以使用以下 **xsi:type** 值定义 **ExtensionPoint** 元素：
   
- **PrimaryCommandSurface**，它是指 Office 中的功能区。
     
- **ContextMenu**，它是当你在 Office UI 中右键单击时出现的快捷菜单。
    
下面的示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。
    
> [!IMPORTANT]
> 对于包含 ID 属性的元素，请务必提供唯一 ID。建议将公司名称与 ID 结合使用。例如，请使用以下格式：`<CustomTab id="mycompanyname.mygroupname">`。 
  
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

|**元素**|**说明**|
|:-----|:-----|
|**CustomTab** <br/> |如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。 <br/> |
|**OfficeTab** <br/> |如果想要（使用 **PrimaryCommandSurface**）扩展默认 Office 功能区选项卡，则为必需项。如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。 <br/> 对于与 **id** 属性一起使用的多个 tab 值，请参阅[默认 Office 功能区选项卡的 Tab 值](https://docs.microsoft.com/javascript/office/manifest/officetab?view=office-js)。  <br/> |
|**OfficeMenu** <br/> | 如果要（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： <br/> 当用户选定文本，然后右键单击所选文本时，适用于 Excel 或 Word 的 **ContextMenuText**显示上下文菜单上的项。 <br/> 适用于 Excel 的 **ContextMenuCell**。当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。 <br/> |
|**Group** <br/> |选项卡上的一组用户界面扩展点。一组可以有多达六个控件。**id** 属性是必需的。它是一个最多为 125 个字符的字符串。 <br/> |
|**Label** <br/> |必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 <br/> |
|**Icon** <br/> |必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。 <br/> |
|**Tooltip** <br/> |可选。组的工具提示**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 <br/> |
|**Control** <br/> |每个组都要求至少有一个控件。**Control** 元素可以是 **Button**，也可以是 **Menu**。使用 **Menu** 可指定按钮控件的下拉列表。目前仅支持按钮和菜单。请参阅[按钮控件](https://docs.microsoft.com/javascript/office/manifest/control?view=office-js#Button-control)和[菜单控件](https://docs.microsoft.com/javascript/office/manifest/control?view=office-js#menu-dropdown-button-controls)部分，了解详细信息。 <br/>**注意：** 建议一次添加一个 **Control** 元素及相关 **Resources** 子元素，以便于进行故障排除。          |
   

### <a name="button-controls"></a>按钮控件
当用户选择某个按钮时，将执行一个操作。它可以执行 JavaScript 函数或显示任务窗格。以下示例演示了如何定义两种按钮。第一个按钮在不显示 UI 的情况下运行 JavaScript 函数，第二个按钮显示任务窗格。在 **Control** 元素中：        

- **type** 属性是必需的，并且必须设置为 **Button**。
    
- **Control** 元素的 **id** 属性是一个最多为 125 个字符的字符串。
    
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

|**元素**|**说明**|
|:-----|:-----|
|**Label** <br/> |必需。按钮文本。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 <br/> |
|**Tooltip** <br/> |可选。按钮的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 <br/> |
|**Supertip** <br/> | 必需。此按钮的 SuperTip，定义如下： <br/> **标题** <br/>  必需。supertip 的文本。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 ShortStrings 元素的子元素，而  元素是“Resources”元素的子元素。 ************************<br/> **说明** <br/>  必需。supertip 的说明。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 LongStrings 元素的子元素，而  元素是“Resources”元素的子元素。 ************************<br/> |
|**Icon** <br/> | 必需。包含按钮的 **Image** 元素。图像文件必须为 .png 格式。 <br/> **Image** <br/>  定义按钮上要显示的图像。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性指示图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。 <br/> |
|**操作** <br/> | 必需。指定用户选择按钮时将执行的操作。可以为 **xsi:type** 属性指定下列任意值之一： <br/> **ExecuteFunction**，它运行位于 **FunctionFile** 引用的文件中的 JavaScript 函数。**ExecuteFunction** 不显示 UI。**FunctionName** 子元素指定要执行的函数的名称。 <br/> **ShowTaskPane**，它显示任务窗格外接程序。**SourceLocation** 子元素指定要显示的任务窗格外接程序的源文件位置。**resid** 属性必须设置为 **Resources** 元素的 **Urls** 元素中 **Url** 元素的 **id** 属性的值。 <br/> |
   

### <a name="menu-controls"></a>菜单控件
**Menu** 控件可与 **PrimaryCommandSurface** 或 **ContextMenu** 结合使用，并定义：
  
- 根级别菜单项。
   
- 子菜单项的列表。
 
与 **PrimaryCommandSurface** 结合使用时，根菜单项显示为功能区上的一个按钮。选择此按钮时，子菜单显示为下拉列表。与 **ContextMenu** 结合使用时，将在上下文菜单上插入包含子菜单的菜单项。在这两种情况中，单个子菜单项均可以执行 JavaScript 函数或显示任务窗格。目前只支持一种子菜单级别。
       
下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。在 **Control** 元素中：
    
- **xsi:type** 属性是必需的，并且必须设置为 **Menu**。
  
- **id** 属性是一个最多为 125 个字符的字符串。
    
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

|**元素**|**说明**|
|:-----|:-----|
|**Label** <br/> |必需。根菜单项的文本。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 <br/> |
|**Tooltip** <br/> |可选。菜单的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 <br/> |
|**SuperTip** <br/> | 必需。菜单的 SuperTip，定义如下： <br/> **标题** <br/>  必需。supertip 的文本。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 ShortStrings 元素的子元素，而  元素是“Resources”元素的子元素。 ************************<br/> **说明** <br/>  必需。supertip 的说明。必须将“resid”属性设置为 String 元素的 id 属性值。String 元素是 LongStrings 元素的子元素，而  元素是“Resources”元素的子元素。 ************************<br/> |
|**Icon** <br/> | 必需。包含菜单的 **Image** 元素。图像文件必须为 .png 格式。 <br/> **Image** <br/>  菜单的图像。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性指示图像的大小（以像素为单位）。要求三种图像大小（以像素为单位）：16、32 和 80。也同样支持五种可选大小（以像素为单位）：20、24、40、48 和 64。 <br/> |
|**Items** <br/> |必需。包含每个子菜单项的 **Item** 元素。每个 **Item** 元素包含的子元素均与[按钮控件](https://docs.microsoft.com/javascript/office/manifest/control?view=office-js#Button-control)相同。  <br/> |
   
## <a name="step-7-add-the-resources-element"></a>步骤 7：添加 Resources 元素

**Resources** 元素包含 **VersionOverrides** 元素的不同子元素所使用的资源。这些资源包括图标、字符串和 URL。清单中的元素可以通过引用资源的 **id** 来使用此资源。使用 **id** 有助于使清单保持有序状态，尤其是当多个区域设置拥有不同的资源版本时。一个 **id** 最多可包含 32 个字符。
  
    
    
以下示例演示了如何使用 **Resources** 元素。每个资源可以具有一个或多个 **Override** 子元素以定义特定区域设置的不同资源。


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

|**Resource**|**说明**|
|:-----|:-----|
|**Images**/ **Image** <br/> | 提供图像文件的 HTTPS URL。每个图像必须定义三个必需的图像大小： <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  也支持下面的图像大小，但不是必需： <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**Urls**/ **Url** <br/> |提供 HTTPS URL 位置。URL 最多可为 2048 个字符。  <br/> |
|**ShortStrings**/ **String** <br/> |**Label** 和 **Title** 元素的文本。每个 **String** 最多可包含 125 个字符。 <br/> |
|**LongStrings**/ **String** <br/> |**Tooltip** 和 **Description** 元素的文本。每个 **String** 最多可包含 250 个字符。 <br/> |
   
> [!NOTE] 
> 必须对 **Image** 和 **Url** 元素中的所有 URL 使用安全套接字层 (SSL)。

### <a name="tab-values-for-default-office-ribbon-tabs"></a>默认 Office 功能区选项卡的 tab 值
在 Excel 和 Word 中，可以使用默认 Office UI 选项卡，在功能区上添加加载项命令。下表列出了可用于 **OfficeTab** 元素的 **id** 属性的值。这些 Tab 值区分大小写。

|**Office 主机应用**|**Tab 值**|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |
   
## <a name="see-also"></a>另请参阅

-  [Excel、Word 和 PowerPoint 加载项命令](../design/add-in-commands.md)      
