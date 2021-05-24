---
title: 清单文件中的 ExtensionPoint 元件
description: 定义 Office UI 中加载项公开功能的位置。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 8f84be1f2dcc43d795026fcd28dc3860c5e07a1e
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590923"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="3292e-103">ExtensionPoint 元素</span><span class="sxs-lookup"><span data-stu-id="3292e-103">ExtensionPoint element</span></span>

 <span data-ttu-id="3292e-104">定义 Office UI 中加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="3292e-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="3292e-105">**ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="3292e-106">属性</span><span class="sxs-lookup"><span data-stu-id="3292e-106">Attributes</span></span>

|  <span data-ttu-id="3292e-107">属性</span><span class="sxs-lookup"><span data-stu-id="3292e-107">Attribute</span></span>  |  <span data-ttu-id="3292e-108">必需</span><span class="sxs-lookup"><span data-stu-id="3292e-108">Required</span></span>  |  <span data-ttu-id="3292e-109">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3292e-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="3292e-110">**xsi:type**</span></span>  |  <span data-ttu-id="3292e-111">是</span><span class="sxs-lookup"><span data-stu-id="3292e-111">Yes</span></span>  | <span data-ttu-id="3292e-112">定义的扩展点类型。</span><span class="sxs-lookup"><span data-stu-id="3292e-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="3292e-113">仅适用于 Excel 的扩展点</span><span class="sxs-lookup"><span data-stu-id="3292e-113">Extension points for Excel only</span></span>

- <span data-ttu-id="3292e-114">**CustomFunctions** - 针对 Excel 使用 JavaScript 编写的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3292e-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="3292e-115">[此 XML 示例代码](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)演示如何将 **ExtensionPoint** 元素与 **CustomFunctions** 属性值配合使用，以及如何使用子元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="3292e-116">适用于 Word、Excel、PowerPoint 和 OneNote 加载项命令的扩展点</span><span class="sxs-lookup"><span data-stu-id="3292e-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="3292e-117">**PrimaryCommandSurface** - Office 中的功能区。</span><span class="sxs-lookup"><span data-stu-id="3292e-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="3292e-118">**ContextMenu** - Office UI 中右键单击时出现的快捷菜单。</span><span class="sxs-lookup"><span data-stu-id="3292e-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="3292e-119">下面的示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3292e-p102">对于包含 ID 属性的元素，请务必提供唯一 ID。建议将公司名称与 ID 结合使用。例如，请使用以下格式：<CustomTab id="mycompanyname.mygroupname"></span><span class="sxs-lookup"><span data-stu-id="3292e-p102">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format. <CustomTab id="mycompanyname.mygroupname"></span></span>

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

#### <a name="child-elements"></a><span data-ttu-id="3292e-123">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-123">Child elements</span></span>
 
|<span data-ttu-id="3292e-124">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-124">Element</span></span>|<span data-ttu-id="3292e-125">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-125">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="3292e-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="3292e-126">**CustomTab**</span></span>|<span data-ttu-id="3292e-p103">如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。 </span><span class="sxs-lookup"><span data-stu-id="3292e-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="3292e-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="3292e-130">**OfficeTab**</span></span>|<span data-ttu-id="3292e-131">如果要使用 **PrimaryCommandSurface** Office 应用扩展默认功能区选项卡 (，) 。</span><span class="sxs-lookup"><span data-stu-id="3292e-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="3292e-132">如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="3292e-133">有关详细信息，请参阅 [OfficeTab](officetab.md)。</span><span class="sxs-lookup"><span data-stu-id="3292e-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="3292e-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="3292e-134">**OfficeMenu**</span></span>|<span data-ttu-id="3292e-p105">如果要（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： </span><span class="sxs-lookup"><span data-stu-id="3292e-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="3292e-p106">适用于 Excel 或 Word 的 - **ContextMenuText** 当用户选定文本，然后右键单击所选定的文本时显示上下文菜单上的项。 </span><span class="sxs-lookup"><span data-stu-id="3292e-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="3292e-p107">适用于 Excel 的 - **ContextMenuCell** 当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。</span><span class="sxs-lookup"><span data-stu-id="3292e-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="3292e-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="3292e-141">**Group**</span></span>|<span data-ttu-id="3292e-p108">选项卡上的一组用户界面扩展点。一组可以有多达六个控件。**id** 属性是必需的。它是一个最多为 125 个字符的字符串。 </span><span class="sxs-lookup"><span data-stu-id="3292e-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="3292e-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="3292e-145">**Label**</span></span>|<span data-ttu-id="3292e-146">必需。</span><span class="sxs-lookup"><span data-stu-id="3292e-146">Required.</span></span> <span data-ttu-id="3292e-147">组的标签。</span><span class="sxs-lookup"><span data-stu-id="3292e-147">The label of the group.</span></span> <span data-ttu-id="3292e-148">**resid** 属性的长度不能超过 32 个字符，必须设置为 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="3292e-148">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="3292e-149">**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="3292e-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="3292e-150">**Icon**</span></span>|<span data-ttu-id="3292e-151">必需。</span><span class="sxs-lookup"><span data-stu-id="3292e-151">Required.</span></span> <span data-ttu-id="3292e-152">指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。</span><span class="sxs-lookup"><span data-stu-id="3292e-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="3292e-153">**resid** 属性不能超过 32 个字符，必须设置为 **Image** 元素 **的 id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="3292e-153">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="3292e-154">**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="3292e-155">**size** 属性给出图像的大小（以像素为单位）。</span><span class="sxs-lookup"><span data-stu-id="3292e-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="3292e-156">要求三种图像大小：16、32 和 80。</span><span class="sxs-lookup"><span data-stu-id="3292e-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="3292e-157">也同样支持五种可选大小：20、24、40、48 和 64。</span><span class="sxs-lookup"><span data-stu-id="3292e-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="3292e-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="3292e-158">**Tooltip**</span></span>|<span data-ttu-id="3292e-159">Optional.</span><span class="sxs-lookup"><span data-stu-id="3292e-159">Optional.</span></span> <span data-ttu-id="3292e-160">The tooltip of the group.</span><span class="sxs-lookup"><span data-stu-id="3292e-160">The tooltip of the group.</span></span> <span data-ttu-id="3292e-161">**resid** 属性的长度不能超过 32 个字符，必须设置为 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="3292e-161">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="3292e-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="3292e-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="3292e-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="3292e-163">**Control**</span></span>|<span data-ttu-id="3292e-164">每个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="3292e-164">Each group requires at least one control.</span></span> <span data-ttu-id="3292e-165">Control 元素可以是 Button **或** **Menu。**</span><span class="sxs-lookup"><span data-stu-id="3292e-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="3292e-166">使用 **Menu** 指定按钮控件的下拉列表。</span><span class="sxs-lookup"><span data-stu-id="3292e-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="3292e-167">目前，仅支持“按钮”和“菜单”。</span><span class="sxs-lookup"><span data-stu-id="3292e-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="3292e-168">请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="3292e-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="3292e-169">**注意：**  为了简化疑难解答，我们建议一次添加 **一** 个 Control 元素和相关 **Resources** 子元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="3292e-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="3292e-170">**Script**</span></span>|<span data-ttu-id="3292e-171">使用自定义函数定义和注册代码链接到 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="3292e-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="3292e-172">在开发者预览版中不使用此元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="3292e-173">实际上，HTML 页负责加载所有 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="3292e-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="3292e-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="3292e-174">**Page**</span></span>|<span data-ttu-id="3292e-175">链接到自定义函数的 HTML 页。</span><span class="sxs-lookup"><span data-stu-id="3292e-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="3292e-176">仅适用于 Outlook 的扩展点</span><span class="sxs-lookup"><span data-stu-id="3292e-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="3292e-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="3292e-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="3292e-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="3292e-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="3292e-181">[Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）</span><span class="sxs-lookup"><span data-stu-id="3292e-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="3292e-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="3292e-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface)
- [<span data-ttu-id="3292e-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="3292e-184">LaunchEvent</span></span>](#launchevent)
- [<span data-ttu-id="3292e-185">Events</span><span class="sxs-lookup"><span data-stu-id="3292e-185">Events</span></span>](#events)
- [<span data-ttu-id="3292e-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="3292e-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="3292e-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="3292e-p114">此扩展点将按钮放置在邮件阅读窗体的命令界面。在 Outlook 桌面，它显示在功能区中。</span><span class="sxs-lookup"><span data-stu-id="3292e-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3292e-190">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-190">Child elements</span></span>

|  <span data-ttu-id="3292e-191">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-191">Element</span></span> |  <span data-ttu-id="3292e-192">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3292e-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3292e-194">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3292e-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3292e-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3292e-196">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3292e-197">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3292e-198">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="3292e-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="3292e-200">此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。</span><span class="sxs-lookup"><span data-stu-id="3292e-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3292e-201">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-201">Child elements</span></span>

|  <span data-ttu-id="3292e-202">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-202">Element</span></span> |  <span data-ttu-id="3292e-203">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3292e-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3292e-205">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3292e-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3292e-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3292e-207">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3292e-208">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3292e-209">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="3292e-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="3292e-211">此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="3292e-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3292e-212">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-212">Child elements</span></span>

|  <span data-ttu-id="3292e-213">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-213">Element</span></span> |  <span data-ttu-id="3292e-214">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3292e-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3292e-216">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3292e-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3292e-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3292e-218">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3292e-219">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3292e-220">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="3292e-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="3292e-222">此扩展点将按钮置于向会议与会者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="3292e-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3292e-223">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-223">Child elements</span></span>

|  <span data-ttu-id="3292e-224">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-224">Element</span></span> |  <span data-ttu-id="3292e-225">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3292e-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3292e-227">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3292e-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3292e-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3292e-229">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3292e-230">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3292e-231">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="3292e-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="3292e-232">Module</span><span class="sxs-lookup"><span data-stu-id="3292e-232">Module</span></span>

<span data-ttu-id="3292e-233">此扩展点将按钮置于模块扩展的功能区上。</span><span class="sxs-lookup"><span data-stu-id="3292e-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3292e-234">注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。</span><span class="sxs-lookup"><span data-stu-id="3292e-234">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3292e-235">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-235">Child elements</span></span>

|  <span data-ttu-id="3292e-236">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-236">Element</span></span> |  <span data-ttu-id="3292e-237">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-237">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-238">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3292e-238">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3292e-239">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-239">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3292e-240">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3292e-240">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3292e-241">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="3292e-241">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="3292e-242">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-242">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="3292e-243">此扩展点将按钮置于移动外形规格中的邮件阅读视图的命令界面中。</span><span class="sxs-lookup"><span data-stu-id="3292e-243">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3292e-244">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-244">Child elements</span></span>

|  <span data-ttu-id="3292e-245">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-245">Element</span></span> |  <span data-ttu-id="3292e-246">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-246">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-247">Group</span><span class="sxs-lookup"><span data-stu-id="3292e-247">Group</span></span>](group.md) |  <span data-ttu-id="3292e-248">将按钮组添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="3292e-248">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="3292e-249">此种类型的 **ExtensionPoint** 元素仅能具有一个子元素，即 **Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-249">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="3292e-250">此扩展点中包含的 **Control** 元素必须将 **xsi:type** 属性设置为 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="3292e-250">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="3292e-251">示例</span><span class="sxs-lookup"><span data-stu-id="3292e-251">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface"></a><span data-ttu-id="3292e-252">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3292e-252">MobileOnlineMeetingCommandSurface</span></span>

<span data-ttu-id="3292e-253">此扩展点将适合模式的切换置于移动外形外形中约会的命令图面中。</span><span class="sxs-lookup"><span data-stu-id="3292e-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="3292e-254">会议组织者可以创建联机会议。</span><span class="sxs-lookup"><span data-stu-id="3292e-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="3292e-255">与会者随后可以加入联机会议。</span><span class="sxs-lookup"><span data-stu-id="3292e-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="3292e-256">若要了解有关此方案的信息，请参阅为联机[会议Outlook创建移动外接程序一](../../outlook/online-meeting.md)文。</span><span class="sxs-lookup"><span data-stu-id="3292e-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

> [!NOTE]
> <span data-ttu-id="3292e-257">此扩展点仅在 Android 和 iOS 上受支持，Microsoft 365订阅。</span><span class="sxs-lookup"><span data-stu-id="3292e-257">This extension point is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>
>
> <span data-ttu-id="3292e-258">注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。</span><span class="sxs-lookup"><span data-stu-id="3292e-258">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3292e-259">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-259">Child elements</span></span>

|  <span data-ttu-id="3292e-260">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-260">Element</span></span> |  <span data-ttu-id="3292e-261">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-261">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-262">Control</span><span class="sxs-lookup"><span data-stu-id="3292e-262">Control</span></span>](control.md) |  <span data-ttu-id="3292e-263">将按钮添加到命令图面。</span><span class="sxs-lookup"><span data-stu-id="3292e-263">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="3292e-264">`ExtensionPoint` 此类型的元素只能有一个子元素：一 `Control` 个元素。</span><span class="sxs-lookup"><span data-stu-id="3292e-264">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="3292e-265">此 `Control` 扩展点中包含的 元素必须将 `xsi:type` 属性设置为 `MobileButton` 。</span><span class="sxs-lookup"><span data-stu-id="3292e-265">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="3292e-266">图像 `Icon` 应该使用十六进制代码或其他颜色格式的等效 `#919191` 项以 [灰度显示](https://convertingcolors.com/hex-color-919191.html)。</span><span class="sxs-lookup"><span data-stu-id="3292e-266">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="3292e-267">示例</span><span class="sxs-lookup"><span data-stu-id="3292e-267">Example</span></span>

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

### <a name="launchevent"></a><span data-ttu-id="3292e-268">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="3292e-268">LaunchEvent</span></span>

<span data-ttu-id="3292e-269">通过此扩展点，加载项可以基于桌面设备类型中支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="3292e-269">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="3292e-270">若要了解有关此方案以及受支持事件的完整列表的信息，请参阅 Configure your [Outlook add-in for event-based activation](../../outlook/autolaunch.md)一文。</span><span class="sxs-lookup"><span data-stu-id="3292e-270">To learn more about this scenario and for the full list of supported events, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3292e-271">注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。</span><span class="sxs-lookup"><span data-stu-id="3292e-271">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3292e-272">子元素</span><span class="sxs-lookup"><span data-stu-id="3292e-272">Child elements</span></span>

|  <span data-ttu-id="3292e-273">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-273">Element</span></span> |  <span data-ttu-id="3292e-274">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-274">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="3292e-275">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="3292e-275">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="3292e-276">用于 [基于事件的激活的 LaunchEvent](launchevent.md) 列表。</span><span class="sxs-lookup"><span data-stu-id="3292e-276">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="3292e-277">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3292e-277">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="3292e-278">源 JavaScript 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="3292e-278">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="3292e-279">示例</span><span class="sxs-lookup"><span data-stu-id="3292e-279">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="3292e-280">事件</span><span class="sxs-lookup"><span data-stu-id="3292e-280">Events</span></span>

<span data-ttu-id="3292e-281">此扩展点添加了指定事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3292e-281">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="3292e-282">有关使用此扩展点的信息，请参阅[On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="3292e-282">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3292e-283">注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。</span><span class="sxs-lookup"><span data-stu-id="3292e-283">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

| <span data-ttu-id="3292e-284">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-284">Element</span></span> | <span data-ttu-id="3292e-285">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-285">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-286">Event</span><span class="sxs-lookup"><span data-stu-id="3292e-286">Event</span></span>](event.md) |  <span data-ttu-id="3292e-287">指定事件和事件处理程序函数。</span><span class="sxs-lookup"><span data-stu-id="3292e-287">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="3292e-288">ItemSend 事件示例</span><span class="sxs-lookup"><span data-stu-id="3292e-288">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="3292e-289">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="3292e-289">DetectedEntity</span></span>

<span data-ttu-id="3292e-290">此扩展点在指定实体类型上添加上下文外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="3292e-290">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3292e-291">注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件不适用于此扩展点。</span><span class="sxs-lookup"><span data-stu-id="3292e-291">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

<span data-ttu-id="3292e-292">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="3292e-292">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="3292e-293">此元素类型适用于[支持要求集 1.6 和更高版本的 Outlook 客户端](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="3292e-293">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="3292e-294">元素</span><span class="sxs-lookup"><span data-stu-id="3292e-294">Element</span></span> |  <span data-ttu-id="3292e-295">说明</span><span class="sxs-lookup"><span data-stu-id="3292e-295">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3292e-296">Label</span><span class="sxs-lookup"><span data-stu-id="3292e-296">Label</span></span>](#label) |  <span data-ttu-id="3292e-297">在上下文窗口中指定外接程序的标签。</span><span class="sxs-lookup"><span data-stu-id="3292e-297">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="3292e-298">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3292e-298">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="3292e-299">指定上下文窗口的 URL。</span><span class="sxs-lookup"><span data-stu-id="3292e-299">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="3292e-300">Rule</span><span class="sxs-lookup"><span data-stu-id="3292e-300">Rule</span></span>](rule.md) |  <span data-ttu-id="3292e-301">指定确定外接程序激活时间的一个或多个规则。</span><span class="sxs-lookup"><span data-stu-id="3292e-301">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="3292e-302">标签</span><span class="sxs-lookup"><span data-stu-id="3292e-302">Label</span></span>

<span data-ttu-id="3292e-303">必需。</span><span class="sxs-lookup"><span data-stu-id="3292e-303">Required.</span></span> <span data-ttu-id="3292e-304">组的标签。</span><span class="sxs-lookup"><span data-stu-id="3292e-304">The label of the group.</span></span> <span data-ttu-id="3292e-305">**resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md)元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="3292e-305">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="3292e-306">突出显示要求</span><span class="sxs-lookup"><span data-stu-id="3292e-306">Highlight requirements</span></span>

<span data-ttu-id="3292e-p119">用户可以激活上下文外接程序的唯一方法是与突出显示实体进行交互。开发人员可以使用 `ItemHasKnownEntity` 和`ItemHasRegularExpressionMatch` 规则类型的 `Rule` 元素的 `Highlight` 属性来控制突出显示哪些实体。</span><span class="sxs-lookup"><span data-stu-id="3292e-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="3292e-p120">但是，存在一些需要注意的限制。存在这些限制是为了确保在适用的邮件或约会中始终存在一个突出显示实体，以便为用户提供一种激活外接程序的方法。</span><span class="sxs-lookup"><span data-stu-id="3292e-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="3292e-311">无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="3292e-311">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="3292e-312">如果使用单个规则，`Highlight` 必须设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="3292e-312">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="3292e-313">如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="3292e-313">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="3292e-314">如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="3292e-314">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="3292e-315">DetectedEntity 事件示例</span><span class="sxs-lookup"><span data-stu-id="3292e-315">DetectedEntity event example</span></span>

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
