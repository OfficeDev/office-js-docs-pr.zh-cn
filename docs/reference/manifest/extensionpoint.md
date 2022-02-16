---
title: 清单文件中的 ExtensionPoint 元件
description: 定义 Office UI 中加载项公开功能的位置。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8ccc08a9c0d42edf89c904b8809a530239be4c
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855630"
---
# <a name="extensionpoint-element"></a>ExtensionPoint 元素

 定义 Office UI 中加载项公开功能的位置。 **ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  是  | 定义的扩展点类型。 可能的值取决于Office Host 元素值中定义的 **主机应用程序。**|

## <a name="extension-points-for-excel-onenote-powerpoint-and-word-add-in-commands"></a>Excel、OneNote、PowerPoint 和 Word 外接程序命令的扩展点

在这些主机的一个或多个主机中，有三种类型的扩展点可用。

- [PrimaryCommandSurface](#primarycommandsurface) (Word、Excel、PowerPoint 和 OneNote) - Office 中的功能区。
- [ContextMenu](#contextmenu) (对 Word、Excel、PowerPoint 和 OneNote) 有效 - 选择并按住 (或在 Office UI 中右键单击) 时出现的快捷菜单。
- [CustomFunctions](#customfunctions) (仅对 Excel) 有效 - 使用 JavaScript 为 Excel 编写的自定义函数。

有关子元素和这些扩展点类型的示例，请参阅以下子部分。

### <a name="primarycommandsurface"></a>PrimaryCommandSurface

Word、Excel、PowerPoint 和 OneNote 中的主要命令图面是功能区。

#### <a name="child-elements"></a>子元素

|元素|说明|
|:-----|:-----|
|[CustomTab] (customtab.md|如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。 |
|[OfficeTab](officetab.md)|如果要使用 **PrimaryCommandSurface** Office 应用扩展默认功能区选项卡 (，) 。 如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。|

#### <a name="example"></a>示例

以下示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface 一同使用**。 它将自定义选项卡添加到功能区。

> [!IMPORTANT]
> 对于包含 ID 属性的元素，请务必提供唯一的 ID。

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### <a name="contextmenu"></a>ContextMenu

上下文菜单是当你在用户界面中右键单击时出现的Office菜单。

#### <a name="child-elements"></a>子元素
 
|元素|说明|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|如果要使用 **ContextMenu** 命令将外接程序命令添加到默认上下文菜单 (必需) 。 **id** 属性必须设置为以下字符串之一： <br/> - **ContextMenuText** 当用户右键单击所选文本时，上下文菜单应打开。 <br/> - **ContextMenuCell** 当用户右键单击电子表格中的单元格时，上下文菜单Excel菜单。|

#### <a name="example"></a>示例

下面向电子表格中的单元格添加一个自定义Excel菜单。

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

### <a name="customfunctions"></a>CustomFunctions

用 JavaScript 或 TypeScript 编写的自定义函数Excel。

#### <a name="child-elements"></a>子元素

|元素|说明|
|:-----|:-----|
|[Script](script.md)|必需项。 指向包含自定义函数的定义和注册代码的 JavaScript 文件的链接。|
|[页面](page.md)|必需项。 链接到自定义函数的 HTML 页。|
|[MetaData](metadata.md)|必需项。 定义 Excel 中的自定义函数所使用的元数据设置。|
|[命名空间](namespace.md)|可选。 定义 Excel 中的自定义函数使用的命名空间。|

#### <a name="example"></a>示例

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

## <a name="extension-points-for-outlook"></a>仅适用于 Outlook 的扩展点

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent)
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
  <CustomTab id="Contoso.TabCustom2">
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
  <CustomTab id="Contoso.TabCustom3">
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
  <CustomTab id="Contoso.TabCustom4">
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
  <CustomTab id="Contoso.TabCustom5">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

此扩展点将按钮置于模块扩展的功能区上。

> [!IMPORTANT]
> 注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。

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
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
      <Control  xsi:type="MobileButton id="Contoso.mobileButton1"">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

此扩展点将适合模式的切换置于移动外形外形中约会的命令图面中。 会议组织者可以创建联机会议。 与会者随后可以加入联机会议。 若要了解有关此方案的信息，请参阅为联机[Outlook创建移动](../../outlook/online-meeting.md)外接程序一文。

> [!NOTE]
> 此扩展点仅在 Android 和 iOS 上受支持，Microsoft 365订阅。
>
> 注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [Control](control.md) |  将按钮添加到命令图面。  |

`ExtensionPoint` 此类型的元素只能有一个子元素：一个元素 `Control` 。

此 `Control` 扩展点中包含的 元素必须将 属性 `xsi:type` 设置为 `MobileButton`。

图像 `Icon` 应该使用十六进制代码或其他 `#919191` 颜色格式的等效值以 [灰度显示](https://convertingcolors.com/hex-color-919191.html)。

#### <a name="example"></a>示例

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
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

### <a name="launchevent"></a>LaunchEvent

通过此扩展点，加载项可以基于桌面设备类型中支持的事件进行激活。 若要了解有关此方案以及受支持事件的完整列表，请参阅为基于事件的激活配置 Outlook 外接程序[一](../../outlook/autolaunch.md)文。

> [!IMPORTANT]
> 注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。

#### <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  用于 [基于事件的激活的 LaunchEvent](launchevent.md) 列表。  |
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

此扩展点添加了指定事件的事件处理程序。 有关使用此扩展点的信息，请参阅加载项的Outlook[功能](../../outlook/outlook-on-send-addins.md)。

> [!IMPORTANT]
> 注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。

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

> [!IMPORTANT]
> 注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。

包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。

> [!NOTE]
> 此元素类型适用于[支持要求集 1.6 和更高版本的 Outlook 客户端](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

|  元素 |  说明  |
|:-----|:-----|
|  [Label](#label) |  在上下文窗口中指定外接程序的标签。  |
|  [SourceLocation](sourcelocation.md) |  指定上下文窗口的 URL。  |
|  [Rule](rule.md) |  指定确定外接程序激活时间的一个或多个规则。  |

#### <a name="label"></a>标签

必需。 组的标签。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。

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
