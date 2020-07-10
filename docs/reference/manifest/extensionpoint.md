---
title: 清单文件中的 ExtensionPoint 元件
description: 定义 Office UI 中加载项公开功能的位置。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 44824e0c74b35105833f1f05cdda87bc873a4427
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094454"
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

下面的示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname">

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
|**CustomTab**|Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|**OfficeTab**|如果要使用**PrimaryCommandSurface**) 扩展默认的 Office 应用功能区选项卡 (，则为必需。 如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。 有关详细信息，请参阅 [OfficeTab](officetab.md)。|
|**OfficeMenu**|Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: <br/> - **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> - **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.|
|**Group**|A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.|
|**Icon**|Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.|
|**Tooltip**|Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.|
|**Control**|每个组需要至少一个控件。 **Control**元素可以是**按钮**，也可以是**菜单**。 使用**菜单**指定按钮控件的下拉列表。 目前，仅支持“按钮”和“菜单”。 请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。<br/>**注意：** 为了使故障排除变得更简单，建议一次添加一个**Control**元素和相关的**Resources**子元素。|
|**Script**|使用自定义函数定义和注册代码链接到 JavaScript 文件。 在开发者预览版中不使用此元素。 实际上，HTML 页负责加载所有 JavaScript 文件。|
|**Page**|链接到自定义函数的 HTML 页。|

## <a name="extension-points-for-outlook"></a>仅适用于 Outlook 的扩展点

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface-preview)
- [LaunchEvent](#launchevent-preview)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a>MobileOnlineMeetingCommandSurface (预览) 

> [!NOTE]
> 仅在使用 Microsoft 365 订阅的 Android 上的[预览](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)中支持此扩展点。

此扩展点在命令界面中为移动外观的约会放置一个适合模式的切换。 会议组织者可以创建联机会议。 与会者随后可以加入联机会议。 若要了解有关此方案的详细信息，请参阅为[联机会议提供商文章创建 Outlook 移动外](../../outlook/online-meeting.md)接程序一文。

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [Control](control.md) |  将按钮添加到命令界面。  |

`ExtensionPoint`此类型的元素只能有一个子元素：一个 `Control` 元素。

`Control`此扩展点中包含的元素的属性必须 `xsi:type` 设置为 `MobileButton` 。

`Icon`图像应使用十六进制代码 `#919191` 或以[其他颜色格式](https://convertingcolors.com/hex-color-919191.html)的等效项进行灰度。

#### <a name="example"></a>示例

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="launchevent-preview"></a>LaunchEvent (预览) 

> [!NOTE]
> 仅在使用 Microsoft 365 订阅的 Outlook 网页[预览版](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)中支持此扩展点。

此扩展点使外接程序能够根据桌面外形规格中受支持的事件进行激活。 目前，唯一受支持的事件是 `OnNewMessageCompose` 和 `OnNewAppointmentOrganizer` 。 若要了解有关此方案的详细信息，请参阅[Configure The Outlook 外接程序以获取基于事件的激活一](../../outlook/autolaunch.md)文。

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  基于事件的激活的[LaunchEvent](launchevent.md)列表。  |
| [SourceLocation](sourcelocation.md) |  源 JavaScript 文件的位置。  |

#### <a name="example"></a>示例

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### <a name="events"></a>事件

此扩展点添加了指定事件的事件处理程序。 有关使用此扩展点的详细信息，请参阅[On a send feature For Outlook 外接程序](../../outlook/outlook-on-send-addins.md)。

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
> 此元素类型适用于[支持要求集 1.6 和更高版本的 Outlook 客户端](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

|  元素 |  说明  |
|:-----|:-----|
|  [Label](#label) |  在上下文窗口中指定外接程序的标签。  |
|  [SourceLocation](sourcelocation.md) |  指定上下文窗口的 URL。  |
|  [Rule](rule.md) |  指定确定外接程序激活时间的一个或多个规则。  |

#### <a name="label"></a>标签

必需。 组的标签。 **Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。

#### <a name="highlight-requirements"></a>突出显示要求

The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.

However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.

- 无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。
- 如果使用单个规则，`Highlight` 必须设置为 `all`。
- 如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。
- 如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。

#### <a name="detectedentity-event-example"></a>DetectedEntity 事件示例

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
