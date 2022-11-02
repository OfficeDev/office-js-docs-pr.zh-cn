---
title: 在清单中创建 Excel、PowerPoint 和 Word 加载项命令
description: 在清单中使用 VersionOverrides 定义 Excel、PowerPoint 和 Word 的外接程序命令。 加载项命令可用于创建 UI 元素，也可用于添加按钮或列表，同时还能执行操作。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 82e921fef7ba37deaa2b20f9f2aa684304cd44ba
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810181"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a>在清单中创建 Excel、PowerPoint 和 Word 加载项命令

> [!NOTE]
> Outlook 中也支持加载项命令。 有关详细信息，请参阅 [Outlook 的外接程序命令](../outlook/add-in-commands-for-outlook.md)

在清单中使用 **[VersionOverrides](/javascript/api/manifest/versionoverrides)** 定义 Excel、PowerPoint 和 Word 的外接程序命令。 加载项命令提供了使用执行操作的特定 UI 元素来自定义默认的 Office 用户界面 (UI) 的简单方法。 有关外接程序命令的简介，请参阅 [Excel、PowerPoint 和 Word 的外接程序命令](../design/add-in-commands.md)。

本文介绍如何编辑清单以定义外接程序命令，以及如何为 [函数命令](../design/add-in-commands.md#types-of-add-in-commands)创建代码。 下图显示了用来定义外接程序命令的元素的层次结构。 本文将具体介绍这些元素。

![清单中的外接程序命令元素概述。 此处的顶级节点是具有子级主机和资源的 VersionOverrides。 在“主机”下，依次为“主机”和“DesktopFormFactor”。 DesktopFormFactor 下是 FunctionFile 和 ExtensionPoint。 在 ExtensionPoint 下是 CustomTab 或 OfficeTab 和 Office 菜单。 在 CustomTab 或 Office 选项卡下，依次为“组”、“控制”、“操作”。 在“Office 菜单”下，依次为“控制”、“操作”。 在“资源” (VersionOverrides 的子级) 为“图像”、“URL”、“ShortString”和“LongString”。](../images/version-overrides.png)

## <a name="step-1-create-the-project"></a>步骤 1：创建项目

建议按照某个快速入门（例如 [生成 Excel 任务窗格加载项](../quickstarts/excel-quickstart-jquery.md)）创建项目。 Excel、PowerPoint 和 Word 的每个快速入门都会生成一个项目，其中包含加载项命令 (按钮) 以显示任务窗格。 在使用外接程序命令之前，请确保已阅读 [Excel、PowerPoint 和 Word](../design/add-in-commands.md) 的外接程序命令。

## <a name="step-2-create-a-task-pane-add-in"></a>第 2 步：创建任务窗格加载项

若要开始使用外接程序命令，必须首先创建任务窗格加载项，然后按照本文所述修改加载项的清单。 不能将外接程序命令与内容加载项一起使用。如果要更新现有清单，则必须添加相应的 **XML 命名空间** ，并将 元素添加到 **\<VersionOverrides\>** 清单，如 [步骤 3：添加 VersionOverrides 元素](#step-3-add-versionoverrides-element)中所述。

以下示例显示了 Office 2013 外接程序的清单。 此清单中没有加载项命令，因为没有 **\<VersionOverrides\>** 元素。 Office 2013 不支持外接程序命令，但通过添加到 **\<VersionOverrides\>** 此清单，外接程序将在 Office 2013 和 Office 2016 中运行。 在 Office 2013 中，外接程序不会显示外接程序命令，并使用 的值 **\<SourceLocation\>** 将外接程序作为单个任务窗格加载项运行。 在 Office 2016 中，如果未包含任何 **\<VersionOverrides\>** 元素，外接程序的任务窗格会自动打开到 中指定的 **\<SourceLocation\>** URL。 但是，如果包含 **\<VersionOverrides\>**，则外接程序仅显示外接程序命令，并且最初不会显示外接程序的任务窗格。
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="https://www.contoso.com/Images/Icon_32.png" />
  <SupportUrl DefaultValue="https://www.contoso.com/contact" />
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

元素 **\<VersionOverrides\>** 是包含外接程序命令定义的根元素。 **\<VersionOverrides\>** 是清单中 元素的 **\<OfficeApp\>** 子元素。 下表列出了 元素的属性 **\<VersionOverrides\>** 。

|属性|说明|
|:-----|:-----|
|**xmlns** <br/> | 必需。 架构位置必须是 `http://schemas.microsoft.com/office/taskpaneappversionoverrides`。 <br/> |
|**xsi:type** <br/> |必需。 架构版本。 本文中所述的版本为"VersionOverridesV1_0"。  <br/> |

下表标识 的子元素 **\<VersionOverrides\>**。
  
|元素|说明|
|:-----|:-----|
|**\<Description\>** <br/> |可选。 Describes the add-in. 此子 **\<Description\>** 元素替代清单父部分中的上 **\<Description\>** 一个元素。 此 **\<Description\>** 元素的 **resid** 属性设置为元素的 **\<String\>** **ID**。 元素 **\<String\>** 包含 的文本 **\<Description\>**。 <br/> |
|**\<Requirements\>** <br/> |可选。 指定外接程序要求的最低要求集和 Office.js 的版本。 此子 **\<Requirements\>** 元素替代 **\<Requirements\>** 清单的父部分中的 元素。 有关详细信息，请参阅 [指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。  <br/> |
|**\<Hosts\>** <br/> |必填。 指定 Office 应用程序的集合。 子 **\<Hosts\>** 元素替代 **\<Hosts\>** 清单的父部分中的 元素。 必须包含已设置为“Workbook”或“Document”的 **xsi:type** 属性 <br/> |
|**\<Resources\>** <br/> |定义其他清单元素引用的资源集合（字符串、URL 和图像）。 例如， **\<Description\>** 元素的值引用 中的 **\<Resources\>** 子元素。 **\<Resources\>** 本文后面的 [步骤 7：添加 Resources 元素](#step-7-add-the-resources-element)中介绍了该元素。 <br/> |

以下示例演示如何使用 **\<VersionOverrides\>** 元素及其子元素。

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

元素 **\<Hosts\>** 包含一个或多个 **\<Host\>** 元素。 元素 **\<Host\>** 指定特定的 Office 应用程序。 元素 **\<Host\>** 包含子元素，这些元素指定要在该 Office 应用程序中安装外接程序后显示的外接程序命令。 若要在两个或多个不同的 Office 应用程序中显示相同的外接程序命令，必须在每个 **\<Host\>** 中复制子元素。

元素 **\<DesktopFormFactor\>** 指定在浏览器) 和 Windows Office web 版 (中运行的加载项的设置。

下面是 、 **\<Host\>** 和 **\<DesktopFormFactor\>** 元素的示例 **\<Hosts\>**。

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

元素 **\<FunctionFile\>** 指定一个文件，该文件包含当外接程序命令使用 **ExecuteFunction** 操作时要运行的 JavaScript 代码 (有关说明) ，请参阅 [Button 控件](/javascript/api/manifest/control-button) 。 元素 **\<FunctionFile\>** 的 **resid** 属性设置为 HTML 文件，其中包含外接程序命令所需的所有 JavaScript 文件。 You can't link directly to a JavaScript file. You can only link to an HTML file. 文件名指定为 元素中的 **\<Resources\>** 元素 **\<Url\>**。

下面是 元素的示例 **\<FunctionFile\>** 。
  
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

元素引用的 HTML 文件中的 **\<FunctionFile\>** JavaScript 必须调用 `Office.initialize`。 元素 **\<FunctionName\>** (有关使用 中的 **\<FunctionFile\>** 函数的说明) ，请参阅 [Button 控件](/javascript/api/manifest/control-button)。

以下代码演示如何实现 所使用的 **\<FunctionName\>** 函数。

```html
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Define the function.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("Function command works. Button ID=" + event.source.id,
            function (asyncResult) {
                const error = asyncResult.error;
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
    
    // You must register the function with the following line.
    Office.actions.associate("writeText", writeText);
</script>
```

> [!IMPORTANT]
> The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.

## <a name="step-6-add-extensionpoint-elements"></a>第 6 步：添加 ExtensionPoint 元素

元素 **\<ExtensionPoint\>** 定义加载项命令在 Office UI 中的显示位置。 可以使用这些 **xsi：type** 值定义 **\<ExtensionPoint\>** 元素。

- **PrimaryCommandSurface**，它是指 Office 中的功能区。

- **ContextMenu**，它是当你在 Office UI 中右键单击时出现的快捷菜单。

以下示例演示如何将 **\<ExtensionPoint\>** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值一起使用，以及应与每个元素一起使用的子元素。

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.
  
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

|元素|说明|
|:-----|:-----|
|**\<CustomTab\>** <br/> |如果要使用 **PrimaryCommandSurface**) 将自定义选项卡添加到功能区 (，则是必需的。 如果使用 **\<CustomTab\>** 元素，则不能使用 **\<OfficeTab\>** 元素。 **id** 属性是必需的。 <br/> |
|**\<OfficeTab\>** <br/> |如果要使用 **PrimaryCommandSurface**)  (扩展默认的 Office 应用功能区选项卡，则是必需的。 如果使用 **\<OfficeTab\>** 元素，则不能使用 **\<CustomTab\>** 元素。 <br/> 有关用于 **id** 属性的更多选项卡值，请参阅 [默认 Office 应用功能区选项卡的选项卡值](/javascript/api/manifest/officetab)。  <br/> |
|**\<OfficeMenu\>** <br/> | 如果要使用 **ContextMenu**) 将加载项命令添加到默认上下文菜单 (，则是必需的。 **id** 属性必须设置为： <br/> **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> 适用于 Excel 的 **ContextMenuCell**。 当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。 <br/> |
|**\<Group\>** <br/> |A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. <br/> |
|**\<Label\>** <br/> |必填。 组的标签。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<ShortStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Icon\>** <br/> |必填。 指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。 **resid** 属性必须设置为元素的 **id** 属性 **\<Image\>** 的值。 元素 **\<Image\>** 是 元素的 **\<Images\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 **size** 属性给出图像的大小（以像素为单位）。 要求三种图像大小：16、32 和 80。 也同样支持五种可选大小：20、24、40、48 和 64。 <br/> |
|**\<Tooltip\>** <br/> |可选。 The tooltip of the group. **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<LongStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Control\>** <br/> |每个组需要至少一个控件。 **\<Control\>** 元素可以是 **Button** 或 **Menu**。 使用 **Menu** 指定按钮控件的下拉列表。 目前，仅支持“按钮”和“菜单”。 有关详细信息 [，请参阅按钮控件](/javascript/api/manifest/control-button) 和 [菜单控件](/javascript/api/manifest/control-menu) 。 <br/>**注意：** 为了更轻松地进行故障排除，建议一次添加一个 **\<Control\>** 元素和相关 **\<Resources\>** 子元素。          |

### <a name="button-controls"></a>按钮控件

当用户选择某个按钮时，将执行一个操作。 它可以执行 JavaScript 函数或显示任务窗格。 以下示例演示了如何定义两种按钮。 第一个按钮在不显示 UI 的情况下运行 JavaScript 函数，第二个按钮显示任务窗格。 在 元素中 **\<Control\>** ：

- **type** 属性是必需的，并且必须设置为 **Button**。

- 元素的 **\<Control\>** **id** 属性是最多 125 个字符的字符串。

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

|元素|说明|
|:-----|:-----|
|**\<Label\>** <br/> |必需。 按钮的文本。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<ShortStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Tooltip\>** <br/> |可选。 按钮的工具提示。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<LongStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Supertip\>** <br/> | 必填。 此按钮的 SuperTip，定义如下： <br/> **标题** <br/>  必需。 supertip 的文本。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<ShortStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> **\<Description\>** <br/>  必填。 supertip 的说明。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<LongStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Icon\>** <br/> | 必填。 包含 **\<Image\>** 按钮的元素。 图像文件必须为 .png 格式。 <br/> **\<Image\>** <br/>  定义要显示在按钮上的图像。 **resid** 属性必须设置为元素的 **id** 属性 **\<Image\>** 的值。 元素 **\<Image\>** 是 元素的 **\<Images\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. 也同样支持五种可选大小：20、24、40、48 和 64。 <br/> |
|**\<Action\>** <br/> | Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: <br/> **ExecuteFunction**，它运行位于 引用的文件中的 **\<FunctionFile\>** JavaScript 函数。 子 **\<FunctionName\>** 元素指定要执行的函数的名称。 <br/> **ShowTaskPane**，显示加载项的任务窗格。 子 **\<SourceLocation\>** 元素指定要显示的页面的源文件位置。 **resid** 属性必须设置为 元素中 **\<Urls\>** **\<Resources\>** 元素的 **\<Url\>** **id** 属性的值。 <br/> |

### <a name="menu-controls"></a>菜单控件

**Menu** 控件可与 **PrimaryCommandSurface** 或 **ContextMenu** 结合使用，并定义：
  
- 根级别菜单项。
- 子菜单项的列表。

When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.

下面的示例演示如何定义具有两个子菜单项的菜单项。 第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。 在 元素中 **\<Control\>** ：

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

|元素|说明|
|:-----|:-----|
|**\<Label\>** <br/> |必填。 根菜单项的文本。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<ShortStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Tooltip\>** <br/> |可选。 菜单的工具提示。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<LongStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<SuperTip\>** <br/> | 必填。 菜单的 SuperTip，定义如下： <br/> **\<Title\>** <br/>  必填。 supertip 的文本。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<ShortStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> **\<Description\>** <br/>  必填。 supertip 的说明。 **resid** 属性必须设置为元素的 **id** 属性 **\<String\>** 的值。 元素 **\<String\>** 是 元素的 **\<LongStrings\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 <br/> |
|**\<Icon\>** <br/> | 必填。 包含 **\<Image\>** 菜单的元素。 图像文件必须为 .png 格式。 <br/> **\<Image\>** <br/>  菜单的图像。 **resid** 属性必须设置为元素的 **id** 属性 **\<Image\>** 的值。 元素 **\<Image\>** 是 元素的 **\<Images\>** 子元素，它是 元素的 **\<Resources\>** 子元素。 The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. 还支持五个可选大小（以像素为单位）：20、24、40、48 和 64。 <br/> |
|**\<Items\>** <br/> |必填。 包含 **\<Item\>** 每个子菜单项的元素。 每个 **\<Item\>** 元素包含与 [Button 控件](/javascript/api/manifest/control-button)相同的子元素。  <br/> |

## <a name="step-7-add-the-resources-element"></a>步骤 7：添加 Resources 元素

元素 **\<Resources\>** 包含元素的不同子元素 **\<VersionOverrides\>** 使用的资源。 这些资源包括图标、字符串和 URL。 清单中的元素可以通过引用资源的 **id** 来使用此资源。 使用 **id** 有助于使清单保持有序状态，尤其是当多个区域设置拥有不同的资源版本时。 一个 **id** 最多可包含 32 个字符。
  
下面演示了如何使用 **\<Resources\>** 元素的示例。 每个资源可以具有一个或多个 **\<Override\>** 子元素，用于为特定区域设置定义不同的资源。

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

|资源|说明|
|:-----|:-----|
|**\<Images\>**/ **\<Image\>** <br/> | 提供图像文件的 HTTPS URL。 每个图像必须定义三个必需的图像大小： <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  也支持下面的图像大小，但不是必需： <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**\<Urls\>**/ **\<Url\>** <br/> |提供 HTTPS URL 位置。 URL 最多可为 2048 个字符。  <br/> |
|**\<ShortStrings\>**/ **\<String\>** <br/> |和 **\<Title\>** 元素的文本 **\<Label\>**。 每个 **\<String\>** 字符最多包含 125 个字符。 <br/> |
|**\<LongStrings\>**/ **\<String\>** <br/> |和 **\<Description\>** 元素的文本 **\<Tooltip\>**。 每个 **\<String\>** 字符最多包含 250 个字符。 <br/> |

> [!NOTE]
> 必须为 和 **\<Url\>** 元素中的所有 URL **\<Image\>** 使用安全套接字层 (SSL) 。

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a>默认 Office 应用功能区选项卡的选项卡值

在 Excel 和 Word 中，可以使用默认 Office UI 选项卡，在功能区上添加加载项命令。 下表列出了可用于 元素的 **\<OfficeTab\>** **id** 属性的值。 这些 Tab 值区分大小写。

|Office 客户端应用程序|Tab 值|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>另请参阅

- [Excel、PowerPoint 和 Word 的加载项命令](../design/add-in-commands.md)
- [示例：使用命令按钮创建 Excel 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [示例：使用命令按钮创建 Word 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [示例：使用命令按钮创建 PowerPoint 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
