---
title: 清单文件中的 ExtensionPoint 元件
description: ''
ms.date: 03/11/2018
localization_priority: Priority
ms.openlocfilehash: 4473790a0dd0daeae8042f8ba15421b8e3f9dc64
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477563"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="76b8c-102">ExtensionPoint 元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-102">ExtensionPoint element</span></span>

 <span data-ttu-id="76b8c-103">定义 Office UI 中加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="76b8c-103">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="76b8c-104">**ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-104">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="76b8c-105">属性</span><span class="sxs-lookup"><span data-stu-id="76b8c-105">Attributes</span></span>

|  <span data-ttu-id="76b8c-106">属性</span><span class="sxs-lookup"><span data-stu-id="76b8c-106">Attribute</span></span>  |  <span data-ttu-id="76b8c-107">必需</span><span class="sxs-lookup"><span data-stu-id="76b8c-107">Required</span></span>  |  <span data-ttu-id="76b8c-108">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-108">Description</span></span>  |
|:-----|:-----|:-----|
|  **<span data-ttu-id="76b8c-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="76b8c-109">xsi:type</span></span>**  |  <span data-ttu-id="76b8c-110">是</span><span class="sxs-lookup"><span data-stu-id="76b8c-110">Yes</span></span>  | <span data-ttu-id="76b8c-111">定义的扩展点类型。</span><span class="sxs-lookup"><span data-stu-id="76b8c-111">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="76b8c-112">仅适用于 Excel 的扩展点</span><span class="sxs-lookup"><span data-stu-id="76b8c-112">Extension points for Excel only</span></span>

- <span data-ttu-id="76b8c-113">**CustomFunctions** - 针对 Excel 使用 JavaScript 编写的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="76b8c-113">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="76b8c-114">[此 XML 示例代码](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)演示如何将 **ExtensionPoint** 元素与 **CustomFunctions** 属性值配合使用，以及如何使用子元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-114">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="76b8c-115">适用于 Word、Excel、PowerPoint 和 OneNote 加载项命令的扩展点</span><span class="sxs-lookup"><span data-stu-id="76b8c-115">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="76b8c-116">**PrimaryCommandSurface** - Office 中的功能区。</span><span class="sxs-lookup"><span data-stu-id="76b8c-116">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="76b8c-117">**ContextMenu** - Office UI 中右键单击时出现的快捷菜单。</span><span class="sxs-lookup"><span data-stu-id="76b8c-117">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="76b8c-118">下面的示例演示如何将  **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-118">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="76b8c-p102">对于包含 ID 属性的元素，请确保提供唯一 ID。我们建议您将您的公司名称与您的 ID 配合使用。例如，使用以下格式。 </span><span class="sxs-lookup"><span data-stu-id="76b8c-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. </span></span><CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="76b8c-122">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-122">Child elements</span></span>
 
|**<span data-ttu-id="76b8c-123">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-123">Element</span></span>**|**<span data-ttu-id="76b8c-124">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-124">Description</span></span>**|
|:-----|:-----|
|**<span data-ttu-id="76b8c-125">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-125">CustomTab</span></span>**|<span data-ttu-id="76b8c-p103">如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|**<span data-ttu-id="76b8c-129">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-129">OfficeTab</span></span>**|<span data-ttu-id="76b8c-p104">如果想要（使用 **PrimaryCommandSurface**）扩展默认 Office 功能区选项卡，则为必需项。如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。有关详细信息，请参阅 [OfficeTab](officetab.md)。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|**<span data-ttu-id="76b8c-133">OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="76b8c-133">OfficeMenu</span></span>**|<span data-ttu-id="76b8c-p105">如果正（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： </span><span class="sxs-lookup"><span data-stu-id="76b8c-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="76b8c-p106">适用于 Excel 或 Word 的 - **ContextMenuText**当用户选定文本，然后右键单击所选定的文本时显示上下文菜单上的项。 </span><span class="sxs-lookup"><span data-stu-id="76b8c-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="76b8c-p107">适用于 Excel 的 - **ContextMenuCell**当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|**<span data-ttu-id="76b8c-140">Group</span><span class="sxs-lookup"><span data-stu-id="76b8c-140">Group</span></span>**|<span data-ttu-id="76b8c-p108">选项卡上的一组用户界面扩展点。一个组可以有最多六个控件。 **id** 属性是必需项。它是最多使用 125 个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|**<span data-ttu-id="76b8c-144">Label</span><span class="sxs-lookup"><span data-stu-id="76b8c-144">Label</span></span>**|<span data-ttu-id="76b8c-p109">必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|**<span data-ttu-id="76b8c-149">Icon</span><span class="sxs-lookup"><span data-stu-id="76b8c-149">Icon</span></span>**|<span data-ttu-id="76b8c-p110">必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|**<span data-ttu-id="76b8c-157">Tooltip</span><span class="sxs-lookup"><span data-stu-id="76b8c-157">Tooltip</span></span>**|<span data-ttu-id="76b8c-p111">可选。组的工具提示**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|**<span data-ttu-id="76b8c-162">Control</span><span class="sxs-lookup"><span data-stu-id="76b8c-162">Control</span></span>**|<span data-ttu-id="76b8c-163">每个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="76b8c-163">Each group requires at least one control.</span></span> <span data-ttu-id="76b8c-164">**Control** 元素可以是一个**按钮**，也可以是一个**菜单**。</span><span class="sxs-lookup"><span data-stu-id="76b8c-164">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="76b8c-165">使用**菜单**指定按钮控件的下拉列表。</span><span class="sxs-lookup"><span data-stu-id="76b8c-165">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="76b8c-166">目前，仅支持“按钮”和“菜单”。</span><span class="sxs-lookup"><span data-stu-id="76b8c-166">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="76b8c-167">请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="76b8c-167">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="76b8c-168">**注意**  为了使故障排除变得更简单，我们建议一次性添加 **Control** 元素和相关的 **Resources** 子元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-168">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|**<span data-ttu-id="76b8c-169">Script</span><span class="sxs-lookup"><span data-stu-id="76b8c-169">Script</span></span>**|<span data-ttu-id="76b8c-170">使用自定义函数定义和注册代码链接到 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="76b8c-170">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="76b8c-171">在开发者预览版中不使用此元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-171">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="76b8c-172">实际上，HTML 页负责加载所有 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="76b8c-172">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|**<span data-ttu-id="76b8c-173">Page</span><span class="sxs-lookup"><span data-stu-id="76b8c-173">Page</span></span>**|<span data-ttu-id="76b8c-174">链接到自定义函数的 HTML 页。</span><span class="sxs-lookup"><span data-stu-id="76b8c-174">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="76b8c-175">仅适用于 Outlook 的扩展点</span><span class="sxs-lookup"><span data-stu-id="76b8c-175">Extension points for Outlook</span></span>

- [<span data-ttu-id="76b8c-176">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-176">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="76b8c-177">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-177">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="76b8c-178">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-178">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="76b8c-179">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-179">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="76b8c-180">[Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）</span><span class="sxs-lookup"><span data-stu-id="76b8c-180">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="76b8c-181">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-181">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="76b8c-182">Events</span><span class="sxs-lookup"><span data-stu-id="76b8c-182">Events</span></span>](#events)
- [<span data-ttu-id="76b8c-183">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="76b8c-183">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="76b8c-184">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-184">MessageReadCommandSurface</span></span>
<span data-ttu-id="76b8c-p114">此扩展点将按钮放置在邮件阅读窗体的命令界面。在 Outlook 桌面，它显示在功能区中。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="76b8c-187">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-187">Child elements</span></span>

|  <span data-ttu-id="76b8c-188">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-188">Element</span></span> |  <span data-ttu-id="76b8c-189">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-189">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-190">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-190">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76b8c-191">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-191">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76b8c-192">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-192">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76b8c-193">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-193">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76b8c-194">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-194">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76b8c-195">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-195">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="76b8c-196">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-196">MessageComposeCommandSurface</span></span>
<span data-ttu-id="76b8c-197">此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。</span><span class="sxs-lookup"><span data-stu-id="76b8c-197">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76b8c-198">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-198">Child elements</span></span>

|  <span data-ttu-id="76b8c-199">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-199">Element</span></span> |  <span data-ttu-id="76b8c-200">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-200">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-201">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-201">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76b8c-202">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-202">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76b8c-203">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-203">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76b8c-204">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-204">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76b8c-205">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-205">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76b8c-206">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-206">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="76b8c-207">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-207">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="76b8c-208">此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="76b8c-208">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76b8c-209">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-209">Child elements</span></span>

|  <span data-ttu-id="76b8c-210">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-210">Element</span></span> |  <span data-ttu-id="76b8c-211">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-211">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-212">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-212">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76b8c-213">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-213">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76b8c-214">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-214">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76b8c-215">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-215">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76b8c-216">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-216">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76b8c-217">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-217">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="76b8c-218">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-218">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="76b8c-219">此扩展点将按钮置于向会议与会者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="76b8c-219">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76b8c-220">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-220">Child elements</span></span>

|  <span data-ttu-id="76b8c-221">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-221">Element</span></span> |  <span data-ttu-id="76b8c-222">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-222">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-223">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-223">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76b8c-224">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-224">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76b8c-225">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-225">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76b8c-226">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-226">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76b8c-227">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-227">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76b8c-228">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-228">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="76b8c-229">Module</span><span class="sxs-lookup"><span data-stu-id="76b8c-229">Module</span></span>

<span data-ttu-id="76b8c-230">此扩展点将按钮置于模块扩展的功能区上。</span><span class="sxs-lookup"><span data-stu-id="76b8c-230">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76b8c-231">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-231">Child elements</span></span>

|  <span data-ttu-id="76b8c-232">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-232">Element</span></span> |  <span data-ttu-id="76b8c-233">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-233">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-234">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-234">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76b8c-235">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-235">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76b8c-236">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76b8c-236">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76b8c-237">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="76b8c-237">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="76b8c-238">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76b8c-238">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="76b8c-239">此扩展点将按钮置于移动外形规格中的邮件阅读视图的命令界面中。</span><span class="sxs-lookup"><span data-stu-id="76b8c-239">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="76b8c-240">子元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-240">Child elements</span></span>

|  <span data-ttu-id="76b8c-241">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-241">Element</span></span> |  <span data-ttu-id="76b8c-242">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-242">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-243">组</span><span class="sxs-lookup"><span data-stu-id="76b8c-243">Group</span></span>](group.md) |  <span data-ttu-id="76b8c-244">将按钮组添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="76b8c-244">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="76b8c-245">此种类型的 **ExtensionPoint** 元素仅能具有一个子元素，即 **Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="76b8c-245">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="76b8c-246">此扩展点中包含的 **Control** 元素必须将 **xsi:type** 属性设置为 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="76b8c-246">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="76b8c-247">示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-247">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="76b8c-248">事件</span><span class="sxs-lookup"><span data-stu-id="76b8c-248">Events</span></span>

<span data-ttu-id="76b8c-249">此扩展点添加了指定事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="76b8c-249">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="76b8c-250">仅 Office 365 中的 Outlook 网页版支持此元素类型。</span><span class="sxs-lookup"><span data-stu-id="76b8c-250">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="76b8c-251">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-251">Element</span></span> | <span data-ttu-id="76b8c-252">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-252">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-253">Event</span><span class="sxs-lookup"><span data-stu-id="76b8c-253">Event</span></span>](event.md) |  <span data-ttu-id="76b8c-254">指定事件和事件处理程序函数。</span><span class="sxs-lookup"><span data-stu-id="76b8c-254">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="76b8c-255">ItemSend 事件示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-255">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="76b8c-256">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="76b8c-256">DetectedEntity</span></span>

<span data-ttu-id="76b8c-257">此扩展点在指定实体类型上添加上下文外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="76b8c-257">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="76b8c-258">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="76b8c-258">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="76b8c-259">仅 Office 365 中的 Outlook 网页版支持此元素类型。</span><span class="sxs-lookup"><span data-stu-id="76b8c-259">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="76b8c-260">元素</span><span class="sxs-lookup"><span data-stu-id="76b8c-260">Element</span></span> |  <span data-ttu-id="76b8c-261">说明</span><span class="sxs-lookup"><span data-stu-id="76b8c-261">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76b8c-262">标签</span><span class="sxs-lookup"><span data-stu-id="76b8c-262">Label</span></span>](#label) |  <span data-ttu-id="76b8c-263">在上下文窗口中指定外接程序的标签。</span><span class="sxs-lookup"><span data-stu-id="76b8c-263">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="76b8c-264">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="76b8c-264">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="76b8c-265">指定上下文窗口的 URL。</span><span class="sxs-lookup"><span data-stu-id="76b8c-265">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="76b8c-266">Rule</span><span class="sxs-lookup"><span data-stu-id="76b8c-266">Rule</span></span>](rule.md) |  <span data-ttu-id="76b8c-267">指定确定外接程序激活时间的一个或多个规则。</span><span class="sxs-lookup"><span data-stu-id="76b8c-267">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="76b8c-268">标签</span><span class="sxs-lookup"><span data-stu-id="76b8c-268">Label</span></span>

<span data-ttu-id="76b8c-p115">必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="76b8c-272">突出显示要求</span><span class="sxs-lookup"><span data-stu-id="76b8c-272">Highlight requirements</span></span>

<span data-ttu-id="76b8c-p116">用户可以激活上下文外接程序的唯一方法是与突出显示实体进行交互。开发人员可以使用 `ItemHasKnownEntity` 和`ItemHasRegularExpressionMatch` 规则类型的 `Rule` 元素的 `Highlight` 属性来控制突出显示哪些实体。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="76b8c-p117">但是，存在一些需要注意的限制。存在这些限制是为了确保在适用的邮件或约会中始终存在一个突出显示实体，以便为用户提供一种激活外接程序的方法。</span><span class="sxs-lookup"><span data-stu-id="76b8c-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="76b8c-277">无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="76b8c-277">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="76b8c-278">如果使用单个规则，`Highlight` 必须设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="76b8c-278">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="76b8c-279">如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="76b8c-279">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="76b8c-280">如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="76b8c-280">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="76b8c-281">DetectedEntity 事件示例</span><span class="sxs-lookup"><span data-stu-id="76b8c-281">DetectedEntity event example</span></span>

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
