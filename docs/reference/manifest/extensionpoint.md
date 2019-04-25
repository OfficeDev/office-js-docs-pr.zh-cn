---
title: 清单文件中的 ExtensionPoint 元件
description: ''
ms.date: 03/11/2018
localization_priority: Priority
ms.openlocfilehash: 4473790a0dd0daeae8042f8ba15421b8e3f9dc64
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450483"
---
# <a name="extensionpoint-element"></a>ExtensionPoint 元素

 定义 Office UI 中加载项公开功能的位置。 **ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。 

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  是  | 定义的扩展点类型。|

## <a name="extension-points-for-excel-only"></a>仅适用于 Excel 的扩展点

- **CustomFunctions** - 针对 Excel 使用 JavaScript 编写的自定义函数。

[此 XML 示例代码](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)演示如何将 **ExtensionPoint** 元素与 **CustomFunctions** 属性值配合使用，以及如何使用子元素。

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>适用于 Word、Excel、PowerPoint 和 OneNote 加载项命令的扩展点

- **PrimaryCommandSurface** - Office 中的功能区。
- **ContextMenu** - Office UI 中右键单击时出现的快捷菜单。

下面的示例演示如何将  **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。

> [!IMPORTANT] 
> 对于包含 ID 属性的元素，请务必提供唯一 ID。建议将公司名称与 ID 结合使用。例如，请使用以下格式：<CustomTab id="mycompanyname.mygroupname">

```XML
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

#### <a name="child-elements"></a>子元素
 
|**元素**|**说明**|
|:-----|:-----|
|**CustomTab**|如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。|
|**OfficeTab**|如果想要（使用 **PrimaryCommandSurface**）扩展默认 Office 功能区选项卡，则为必需项。如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。有关详细信息，请参阅 [OfficeTab](officetab.md)。|
|**OfficeMenu**|如果正（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： <br/> 适用于 Excel 或 Word 的 - **ContextMenuText**当用户选定文本，然后右键单击所选定的文本时显示上下文菜单上的项。 <br/> 适用于 Excel 的 - **ContextMenuCell**当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。|
|**Group**|选项卡上的一组用户界面扩展点。一个组可以有最多六个控件。 **id** 属性是必需项。它是最多使用 125 个字符的字符串。|
|**Label**|必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。|
|**Icon**|必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。|
|**Tooltip**|可选。组的工具提示**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。|
|**Control**|每个组需要至少一个控件。 **Control** 元素可以是一个**按钮**，也可以是一个**菜单**。 使用**菜单**指定按钮控件的下拉列表。 目前，仅支持“按钮”和“菜单”。 请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。<br/>**注意**  为了使故障排除变得更简单，我们建议一次性添加 **Control** 元素和相关的 **Resources** 子元素。|
|**Script**|使用自定义函数定义和注册代码链接到 JavaScript 文件。 在开发者预览版中不使用此元素。 实际上，HTML 页负责加载所有 JavaScript 文件。|
|**Page**|链接到自定义函数的 HTML 页。|

## <a name="extension-points-for-outlook"></a>仅适用于 Outlook 的扩展点

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface
此扩展点将按钮放置在邮件阅读窗体的命令界面。在 Outlook 桌面，它显示在功能区中。

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface
此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。 

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。 

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

此扩展点将按钮置于向会议与会者显示的窗体的功能区上。 

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

此扩展点将按钮置于模块扩展的功能区上。 

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](customtab.md) |  将命令添加到自定义功能区选项卡。  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface
此扩展点将按钮置于移动外形规格中的邮件阅读视图的命令界面中。

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [Group](group.md) |  将按钮组添加到命令界面。  |

此种类型的 **ExtensionPoint** 元素仅能具有一个子元素，即 **Group** 元素。

此扩展点中包含的 **Control** 元素必须将 **xsi:type** 属性设置为 `MobileButton`。

#### <a name="example"></a>示例
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a>事件

此扩展点添加了指定事件的事件处理程序。

> [!NOTE]
> 仅 Office 365 中的 Outlook 网页版支持此元素类型。

| 元素 | 说明  |
|:-----|:-----|
|  [Event](event.md) |  指定事件和事件处理程序函数。  |

#### <a name="itemsend-event-example"></a>ItemSend 事件示例

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a>DetectedEntity

此扩展点在指定实体类型上添加上下文外接程序激活。

包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。

> [!NOTE]
> 仅 Office 365 中的 Outlook 网页版支持此元素类型。

|  元素 |  说明  |
|:-----|:-----|
|  [Label](#label) |  在上下文窗口中指定外接程序的标签。  |
|  [SourceLocation](sourcelocation.md) |  指定上下文窗口的 URL。  |
|  [Rule](rule.md) |  指定确定外接程序激活时间的一个或多个规则。  |

#### <a name="label"></a>标签

必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。

#### <a name="highlight-requirements"></a>突出显示要求

用户可以激活上下文外接程序的唯一方法是与突出显示实体进行交互。开发人员可以使用 `ItemHasKnownEntity` 和`ItemHasRegularExpressionMatch` 规则类型的 `Rule` 元素的 `Highlight` 属性来控制突出显示哪些实体。

但是，存在一些需要注意的限制。存在这些限制是为了确保在适用的邮件或约会中始终存在一个突出显示实体，以便为用户提供一种激活外接程序的方法。

- 无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。
- 如果使用单个规则，`Highlight` 必须设置为 `all`。
- 如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。
- 如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。

#### <a name="detectedentity-event-example"></a>DetectedEntity 事件示例

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```
