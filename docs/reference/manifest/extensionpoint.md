---
title: 清单文件中的 ExtensionPoint 元件
description: 定义 Office UI 中加载项公开功能的位置。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 40b5ab8c2dcae01238854f5a3bfcc599f6b01a9b
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241033"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="fe522-103">ExtensionPoint 元素</span><span class="sxs-lookup"><span data-stu-id="fe522-103">ExtensionPoint element</span></span>

 <span data-ttu-id="fe522-104">定义 Office UI 中加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="fe522-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="fe522-105">**ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="fe522-106">属性</span><span class="sxs-lookup"><span data-stu-id="fe522-106">Attributes</span></span>

|  <span data-ttu-id="fe522-107">属性</span><span class="sxs-lookup"><span data-stu-id="fe522-107">Attribute</span></span>  |  <span data-ttu-id="fe522-108">必需</span><span class="sxs-lookup"><span data-stu-id="fe522-108">Required</span></span>  |  <span data-ttu-id="fe522-109">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="fe522-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="fe522-110">**xsi:type**</span></span>  |  <span data-ttu-id="fe522-111">是</span><span class="sxs-lookup"><span data-stu-id="fe522-111">Yes</span></span>  | <span data-ttu-id="fe522-112">定义的扩展点类型。</span><span class="sxs-lookup"><span data-stu-id="fe522-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="fe522-113">仅适用于 Excel 的扩展点</span><span class="sxs-lookup"><span data-stu-id="fe522-113">Extension points for Excel only</span></span>

- <span data-ttu-id="fe522-114">**CustomFunctions** - 针对 Excel 使用 JavaScript 编写的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="fe522-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="fe522-115">[此 XML 示例代码](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)演示如何将 **ExtensionPoint** 元素与 **CustomFunctions** 属性值配合使用，以及如何使用子元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="fe522-116">适用于 Word、Excel、PowerPoint 和 OneNote 加载项命令的扩展点</span><span class="sxs-lookup"><span data-stu-id="fe522-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="fe522-117">**PrimaryCommandSurface** - Office 中的功能区。</span><span class="sxs-lookup"><span data-stu-id="fe522-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="fe522-118">**ContextMenu** - Office UI 中右键单击时出现的快捷菜单。</span><span class="sxs-lookup"><span data-stu-id="fe522-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="fe522-119">下面的示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe522-p102">对于包含 ID 属性的元素，请务必提供唯一 ID。建议将公司名称与 ID 结合使用。例如，请使用以下格式：<CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="fe522-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="fe522-123">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-123">Child elements</span></span>
 
|<span data-ttu-id="fe522-124">**元素**</span><span class="sxs-lookup"><span data-stu-id="fe522-124">**Element**</span></span>|<span data-ttu-id="fe522-125">**说明**</span><span class="sxs-lookup"><span data-stu-id="fe522-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="fe522-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="fe522-126">**CustomTab**</span></span>|<span data-ttu-id="fe522-p103">如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。 </span><span class="sxs-lookup"><span data-stu-id="fe522-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="fe522-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="fe522-130">**OfficeTab**</span></span>|<span data-ttu-id="fe522-131">如果想要（使用 **PrimaryCommandSurface**）扩展默认 Office 功能区选项卡，则为必需项。</span><span class="sxs-lookup"><span data-stu-id="fe522-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="fe522-132">如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="fe522-133">有关详细信息，请参阅 [OfficeTab](officetab.md)。</span><span class="sxs-lookup"><span data-stu-id="fe522-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="fe522-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="fe522-134">**OfficeMenu**</span></span>|<span data-ttu-id="fe522-p105">如果要（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： </span><span class="sxs-lookup"><span data-stu-id="fe522-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="fe522-p106">适用于 Excel 或 Word 的 - **ContextMenuText**当用户选定文本，然后右键单击所选定的文本时显示上下文菜单上的项。 </span><span class="sxs-lookup"><span data-stu-id="fe522-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="fe522-p107">适用于 Excel 的 - **ContextMenuCell**当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。</span><span class="sxs-lookup"><span data-stu-id="fe522-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="fe522-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="fe522-141">**Group**</span></span>|<span data-ttu-id="fe522-p108">选项卡上的一组用户界面扩展点。一组可以有多达六个控件。**id** 属性是必需的。它是一个最多为 125 个字符的字符串。 </span><span class="sxs-lookup"><span data-stu-id="fe522-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="fe522-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="fe522-145">**Label**</span></span>|<span data-ttu-id="fe522-p109">必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="fe522-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="fe522-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="fe522-150">**Icon**</span></span>|<span data-ttu-id="fe522-p110">必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。 </span><span class="sxs-lookup"><span data-stu-id="fe522-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="fe522-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="fe522-158">**Tooltip**</span></span>|<span data-ttu-id="fe522-p111">可选。组的工具提示**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。 </span><span class="sxs-lookup"><span data-stu-id="fe522-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="fe522-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="fe522-163">**Control**</span></span>|<span data-ttu-id="fe522-164">每个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="fe522-164">Each group requires at least one control.</span></span> <span data-ttu-id="fe522-165">**Control**元素可以是**按钮**，也可以是**菜单**。</span><span class="sxs-lookup"><span data-stu-id="fe522-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="fe522-166">使用**菜单**指定按钮控件的下拉列表。</span><span class="sxs-lookup"><span data-stu-id="fe522-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="fe522-167">目前，仅支持“按钮”和“菜单”。</span><span class="sxs-lookup"><span data-stu-id="fe522-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="fe522-168">请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="fe522-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="fe522-169">**注意：** 为了使故障排除变得更简单，建议一次添加一个**Control**元素和相关的**Resources**子元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="fe522-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="fe522-170">**Script**</span></span>|<span data-ttu-id="fe522-171">使用自定义函数定义和注册代码链接到 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="fe522-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="fe522-172">在开发者预览版中不使用此元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="fe522-173">实际上，HTML 页负责加载所有 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="fe522-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="fe522-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="fe522-174">**Page**</span></span>|<span data-ttu-id="fe522-175">链接到自定义函数的 HTML 页。</span><span class="sxs-lookup"><span data-stu-id="fe522-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="fe522-176">仅适用于 Outlook 的扩展点</span><span class="sxs-lookup"><span data-stu-id="fe522-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="fe522-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="fe522-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="fe522-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="fe522-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="fe522-181">[Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）</span><span class="sxs-lookup"><span data-stu-id="fe522-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="fe522-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="fe522-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="fe522-184">Events</span><span class="sxs-lookup"><span data-stu-id="fe522-184">Events</span></span>](#events)
- [<span data-ttu-id="fe522-185">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="fe522-185">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="fe522-186">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-186">MessageReadCommandSurface</span></span>

<span data-ttu-id="fe522-p114">此扩展点将按钮放置在邮件阅读窗体的命令界面。在 Outlook 桌面，它显示在功能区中。</span><span class="sxs-lookup"><span data-stu-id="fe522-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="fe522-189">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-189">Child elements</span></span>

|  <span data-ttu-id="fe522-190">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-190">Element</span></span> |  <span data-ttu-id="fe522-191">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-191">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-192">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="fe522-192">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="fe522-193">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-193">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="fe522-194">CustomTab</span><span class="sxs-lookup"><span data-stu-id="fe522-194">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="fe522-195">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-195">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="fe522-196">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-196">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="fe522-197">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-197">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="fe522-198">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-198">MessageComposeCommandSurface</span></span>

<span data-ttu-id="fe522-199">此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。</span><span class="sxs-lookup"><span data-stu-id="fe522-199">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="fe522-200">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-200">Child elements</span></span>

|  <span data-ttu-id="fe522-201">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-201">Element</span></span> |  <span data-ttu-id="fe522-202">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-202">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-203">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="fe522-203">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="fe522-204">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-204">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="fe522-205">CustomTab</span><span class="sxs-lookup"><span data-stu-id="fe522-205">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="fe522-206">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-206">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="fe522-207">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-207">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="fe522-208">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-208">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="fe522-209">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-209">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="fe522-210">此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="fe522-210">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="fe522-211">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-211">Child elements</span></span>

|  <span data-ttu-id="fe522-212">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-212">Element</span></span> |  <span data-ttu-id="fe522-213">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-213">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-214">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="fe522-214">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="fe522-215">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-215">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="fe522-216">CustomTab</span><span class="sxs-lookup"><span data-stu-id="fe522-216">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="fe522-217">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-217">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="fe522-218">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-218">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="fe522-219">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-219">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="fe522-220">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-220">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="fe522-221">此扩展点将按钮置于向会议与会者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="fe522-221">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="fe522-222">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-222">Child elements</span></span>

|  <span data-ttu-id="fe522-223">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-223">Element</span></span> |  <span data-ttu-id="fe522-224">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-224">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-225">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="fe522-225">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="fe522-226">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-226">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="fe522-227">CustomTab</span><span class="sxs-lookup"><span data-stu-id="fe522-227">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="fe522-228">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-228">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="fe522-229">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-229">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="fe522-230">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="fe522-230">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="fe522-231">Module</span><span class="sxs-lookup"><span data-stu-id="fe522-231">Module</span></span>

<span data-ttu-id="fe522-232">此扩展点将按钮置于模块扩展的功能区上。</span><span class="sxs-lookup"><span data-stu-id="fe522-232">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="fe522-233">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-233">Child elements</span></span>

|  <span data-ttu-id="fe522-234">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-234">Element</span></span> |  <span data-ttu-id="fe522-235">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-235">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-236">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="fe522-236">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="fe522-237">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-237">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="fe522-238">CustomTab</span><span class="sxs-lookup"><span data-stu-id="fe522-238">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="fe522-239">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe522-239">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="fe522-240">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="fe522-240">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="fe522-241">此扩展点将按钮置于移动外形规格中的邮件阅读视图的命令界面中。</span><span class="sxs-lookup"><span data-stu-id="fe522-241">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="fe522-242">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-242">Child elements</span></span>

|  <span data-ttu-id="fe522-243">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-243">Element</span></span> |  <span data-ttu-id="fe522-244">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-244">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-245">Group</span><span class="sxs-lookup"><span data-stu-id="fe522-245">Group</span></span>](group.md) |  <span data-ttu-id="fe522-246">将按钮组添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="fe522-246">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="fe522-247">此种类型的 **ExtensionPoint** 元素仅能具有一个子元素，即 **Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-247">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="fe522-248">此扩展点中包含的 **Control** 元素必须将 **xsi:type** 属性设置为 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="fe522-248">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="fe522-249">示例</span><span class="sxs-lookup"><span data-stu-id="fe522-249">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="fe522-250">MobileOnlineMeetingCommandSurface （预览）</span><span class="sxs-lookup"><span data-stu-id="fe522-250">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="fe522-251">仅在使用 Office 365 订阅的 Android 上的[预览](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)中支持此扩展点。</span><span class="sxs-lookup"><span data-stu-id="fe522-251">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="fe522-252">此扩展点在命令界面中为移动外观的约会放置一个适合模式的切换。</span><span class="sxs-lookup"><span data-stu-id="fe522-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="fe522-253">会议组织者可以创建联机会议。</span><span class="sxs-lookup"><span data-stu-id="fe522-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="fe522-254">与会者随后可以加入联机会议。</span><span class="sxs-lookup"><span data-stu-id="fe522-254">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="fe522-255">若要了解有关此方案的详细信息，请参阅为[联机会议提供商文章创建 Outlook 移动外](../../outlook/online-meeting.md)接程序一文。</span><span class="sxs-lookup"><span data-stu-id="fe522-255">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="fe522-256">子元素</span><span class="sxs-lookup"><span data-stu-id="fe522-256">Child elements</span></span>

|  <span data-ttu-id="fe522-257">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-257">Element</span></span> |  <span data-ttu-id="fe522-258">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-258">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-259">Control</span><span class="sxs-lookup"><span data-stu-id="fe522-259">Control</span></span>](control.md) |  <span data-ttu-id="fe522-260">将按钮添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="fe522-260">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="fe522-261">`ExtensionPoint`此类型的元素只能有一个子元素：一个`Control`元素。</span><span class="sxs-lookup"><span data-stu-id="fe522-261">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="fe522-262">此`Control`扩展点中包含的元素的`xsi:type`属性必须设置为`MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="fe522-262">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="fe522-263">`Icon`图像应使用十六进制代码`#919191`或以[其他颜色格式](https://convertingcolors.com/hex-color-919191.html)的等效项进行灰度。</span><span class="sxs-lookup"><span data-stu-id="fe522-263">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="fe522-264">示例</span><span class="sxs-lookup"><span data-stu-id="fe522-264">Example</span></span>

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
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
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

### <a name="events"></a><span data-ttu-id="fe522-265">事件</span><span class="sxs-lookup"><span data-stu-id="fe522-265">Events</span></span>

<span data-ttu-id="fe522-266">此扩展点添加了指定事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fe522-266">This extension point adds an event handler for a specified event.</span></span>

| <span data-ttu-id="fe522-267">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-267">Element</span></span> | <span data-ttu-id="fe522-268">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-268">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-269">Event</span><span class="sxs-lookup"><span data-stu-id="fe522-269">Event</span></span>](event.md) |  <span data-ttu-id="fe522-270">指定事件和事件处理程序函数。</span><span class="sxs-lookup"><span data-stu-id="fe522-270">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="fe522-271">ItemSend 事件示例</span><span class="sxs-lookup"><span data-stu-id="fe522-271">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="fe522-272">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="fe522-272">DetectedEntity</span></span>

<span data-ttu-id="fe522-273">此扩展点在指定实体类型上添加上下文外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="fe522-273">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="fe522-274">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="fe522-274">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="fe522-275">此元素类型适用于[支持要求集 1.6 和更高版本的 Outlook 客户端](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="fe522-275">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="fe522-276">元素</span><span class="sxs-lookup"><span data-stu-id="fe522-276">Element</span></span> |  <span data-ttu-id="fe522-277">说明</span><span class="sxs-lookup"><span data-stu-id="fe522-277">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fe522-278">Label</span><span class="sxs-lookup"><span data-stu-id="fe522-278">Label</span></span>](#label) |  <span data-ttu-id="fe522-279">在上下文窗口中指定外接程序的标签。</span><span class="sxs-lookup"><span data-stu-id="fe522-279">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="fe522-280">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fe522-280">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="fe522-281">指定上下文窗口的 URL。</span><span class="sxs-lookup"><span data-stu-id="fe522-281">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="fe522-282">Rule</span><span class="sxs-lookup"><span data-stu-id="fe522-282">Rule</span></span>](rule.md) |  <span data-ttu-id="fe522-283">指定确定外接程序激活时间的一个或多个规则。</span><span class="sxs-lookup"><span data-stu-id="fe522-283">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="fe522-284">标签</span><span class="sxs-lookup"><span data-stu-id="fe522-284">Label</span></span>

<span data-ttu-id="fe522-285">必需。</span><span class="sxs-lookup"><span data-stu-id="fe522-285">Required.</span></span> <span data-ttu-id="fe522-286">组的标签。</span><span class="sxs-lookup"><span data-stu-id="fe522-286">The label of the group.</span></span> <span data-ttu-id="fe522-287">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="fe522-287">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="fe522-288">突出显示要求</span><span class="sxs-lookup"><span data-stu-id="fe522-288">Highlight requirements</span></span>

<span data-ttu-id="fe522-p117">用户可以激活上下文外接程序的唯一方法是与突出显示实体进行交互。开发人员可以使用 `ItemHasKnownEntity` 和`ItemHasRegularExpressionMatch` 规则类型的 `Rule` 元素的 `Highlight` 属性来控制突出显示哪些实体。</span><span class="sxs-lookup"><span data-stu-id="fe522-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="fe522-p118">但是，存在一些需要注意的限制。存在这些限制是为了确保在适用的邮件或约会中始终存在一个突出显示实体，以便为用户提供一种激活外接程序的方法。</span><span class="sxs-lookup"><span data-stu-id="fe522-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="fe522-293">无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="fe522-293">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="fe522-294">如果使用单个规则，`Highlight` 必须设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="fe522-294">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="fe522-295">如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="fe522-295">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="fe522-296">如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="fe522-296">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="fe522-297">DetectedEntity 事件示例</span><span class="sxs-lookup"><span data-stu-id="fe522-297">DetectedEntity event example</span></span>

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
