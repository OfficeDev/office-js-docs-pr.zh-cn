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
# <a name="extensionpoint-element"></a><span data-ttu-id="306cc-103">ExtensionPoint 元素</span><span class="sxs-lookup"><span data-stu-id="306cc-103">ExtensionPoint element</span></span>

 <span data-ttu-id="306cc-104">定义 Office UI 中加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="306cc-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="306cc-105">**ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="306cc-106">属性</span><span class="sxs-lookup"><span data-stu-id="306cc-106">Attributes</span></span>

|  <span data-ttu-id="306cc-107">属性</span><span class="sxs-lookup"><span data-stu-id="306cc-107">Attribute</span></span>  |  <span data-ttu-id="306cc-108">必需</span><span class="sxs-lookup"><span data-stu-id="306cc-108">Required</span></span>  |  <span data-ttu-id="306cc-109">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="306cc-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="306cc-110">**xsi:type**</span></span>  |  <span data-ttu-id="306cc-111">是</span><span class="sxs-lookup"><span data-stu-id="306cc-111">Yes</span></span>  | <span data-ttu-id="306cc-112">定义的扩展点类型。</span><span class="sxs-lookup"><span data-stu-id="306cc-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="306cc-113">仅适用于 Excel 的扩展点</span><span class="sxs-lookup"><span data-stu-id="306cc-113">Extension points for Excel only</span></span>

- <span data-ttu-id="306cc-114">**CustomFunctions** - 针对 Excel 使用 JavaScript 编写的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="306cc-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="306cc-115">[此 XML 示例代码](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)演示如何将 **ExtensionPoint** 元素与 **CustomFunctions** 属性值配合使用，以及如何使用子元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="306cc-116">适用于 Word、Excel、PowerPoint 和 OneNote 加载项命令的扩展点</span><span class="sxs-lookup"><span data-stu-id="306cc-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="306cc-117">**PrimaryCommandSurface** - Office 中的功能区。</span><span class="sxs-lookup"><span data-stu-id="306cc-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="306cc-118">**ContextMenu** - Office UI 中右键单击时出现的快捷菜单。</span><span class="sxs-lookup"><span data-stu-id="306cc-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="306cc-119">下面的示例演示如何将 **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="306cc-120">For elements that contain an ID attribute, make sure you provide a unique ID.</span><span class="sxs-lookup"><span data-stu-id="306cc-120">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="306cc-121">We recommend that you use your company's name along with your ID.</span><span class="sxs-lookup"><span data-stu-id="306cc-121">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="306cc-122">For example, use the following format.</span><span class="sxs-lookup"><span data-stu-id="306cc-122">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="306cc-123">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-123">Child elements</span></span>
 
|<span data-ttu-id="306cc-124">**元素**</span><span class="sxs-lookup"><span data-stu-id="306cc-124">**Element**</span></span>|<span data-ttu-id="306cc-125">**说明**</span><span class="sxs-lookup"><span data-stu-id="306cc-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="306cc-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="306cc-126">**CustomTab**</span></span>|<span data-ttu-id="306cc-127">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="306cc-127">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="306cc-128">If you use the **CustomTab** element, you can't use the **OfficeTab** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-128">If you use the **CustomTab** element, you can't use the **OfficeTab** element.</span></span> <span data-ttu-id="306cc-129">The **id** attribute is required.</span><span class="sxs-lookup"><span data-stu-id="306cc-129">The **id** attribute is required.</span></span>|
|<span data-ttu-id="306cc-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="306cc-130">**OfficeTab**</span></span>|<span data-ttu-id="306cc-131">如果要使用**PrimaryCommandSurface**) 扩展默认的 Office 应用功能区选项卡 (，则为必需。</span><span class="sxs-lookup"><span data-stu-id="306cc-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="306cc-132">如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="306cc-133">有关详细信息，请参阅 [OfficeTab](officetab.md)。</span><span class="sxs-lookup"><span data-stu-id="306cc-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="306cc-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="306cc-134">**OfficeMenu**</span></span>|<span data-ttu-id="306cc-135">Required if you're adding add-in commands to a default context menu (using **ContextMenu**).</span><span class="sxs-lookup"><span data-stu-id="306cc-135">Required if you're adding add-in commands to a default context menu (using **ContextMenu**).</span></span> <span data-ttu-id="306cc-136">The **id** attribute must be set to:</span><span class="sxs-lookup"><span data-stu-id="306cc-136">The **id** attribute must be set to:</span></span> <br/> <span data-ttu-id="306cc-137">- **ContextMenuText** for Excel or Word.</span><span class="sxs-lookup"><span data-stu-id="306cc-137">- **ContextMenuText** for Excel or Word.</span></span> <span data-ttu-id="306cc-138">Displays the item on the context menu when text is selected and then the user right-clicks on the selected text.</span><span class="sxs-lookup"><span data-stu-id="306cc-138">Displays the item on the context menu when text is selected and then the user right-clicks on the selected text.</span></span> <br/> <span data-ttu-id="306cc-139">- **ContextMenuCell** for Excel.</span><span class="sxs-lookup"><span data-stu-id="306cc-139">- **ContextMenuCell** for Excel.</span></span> <span data-ttu-id="306cc-140">Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span><span class="sxs-lookup"><span data-stu-id="306cc-140">Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="306cc-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="306cc-141">**Group**</span></span>|<span data-ttu-id="306cc-142">A group of user interface extension points on a tab. A group can have up to six controls.</span><span class="sxs-lookup"><span data-stu-id="306cc-142">A group of user interface extension points on a tab. A group can have up to six controls.</span></span> <span data-ttu-id="306cc-143">The **id** attribute is required.</span><span class="sxs-lookup"><span data-stu-id="306cc-143">The **id** attribute is required.</span></span> <span data-ttu-id="306cc-144">It's a string with a maximum of 125 characters.</span><span class="sxs-lookup"><span data-stu-id="306cc-144">It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="306cc-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="306cc-145">**Label**</span></span>|<span data-ttu-id="306cc-146">Required.</span><span class="sxs-lookup"><span data-stu-id="306cc-146">Required.</span></span> <span data-ttu-id="306cc-147">The label of the group.</span><span class="sxs-lookup"><span data-stu-id="306cc-147">The label of the group.</span></span> <span data-ttu-id="306cc-148">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-148">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="306cc-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="306cc-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="306cc-150">**Icon**</span></span>|<span data-ttu-id="306cc-151">Required.</span><span class="sxs-lookup"><span data-stu-id="306cc-151">Required.</span></span> <span data-ttu-id="306cc-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span><span class="sxs-lookup"><span data-stu-id="306cc-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="306cc-153">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-153">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="306cc-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="306cc-155">The **size** attribute gives the size, in pixels, of the image.</span><span class="sxs-lookup"><span data-stu-id="306cc-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="306cc-156">Three image sizes are required: 16, 32, and 80.</span><span class="sxs-lookup"><span data-stu-id="306cc-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="306cc-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span><span class="sxs-lookup"><span data-stu-id="306cc-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="306cc-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="306cc-158">**Tooltip**</span></span>|<span data-ttu-id="306cc-159">Optional.</span><span class="sxs-lookup"><span data-stu-id="306cc-159">Optional.</span></span> <span data-ttu-id="306cc-160">The tooltip of the group.</span><span class="sxs-lookup"><span data-stu-id="306cc-160">The tooltip of the group.</span></span> <span data-ttu-id="306cc-161">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-161">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="306cc-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="306cc-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="306cc-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="306cc-163">**Control**</span></span>|<span data-ttu-id="306cc-164">每个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="306cc-164">Each group requires at least one control.</span></span> <span data-ttu-id="306cc-165">**Control**元素可以是**按钮**，也可以是**菜单**。</span><span class="sxs-lookup"><span data-stu-id="306cc-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="306cc-166">使用**菜单**指定按钮控件的下拉列表。</span><span class="sxs-lookup"><span data-stu-id="306cc-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="306cc-167">目前，仅支持“按钮”和“菜单”。</span><span class="sxs-lookup"><span data-stu-id="306cc-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="306cc-168">请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="306cc-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="306cc-169">**注意：** 为了使故障排除变得更简单，建议一次添加一个**Control**元素和相关的**Resources**子元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="306cc-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="306cc-170">**Script**</span></span>|<span data-ttu-id="306cc-171">使用自定义函数定义和注册代码链接到 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="306cc-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="306cc-172">在开发者预览版中不使用此元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="306cc-173">实际上，HTML 页负责加载所有 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="306cc-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="306cc-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="306cc-174">**Page**</span></span>|<span data-ttu-id="306cc-175">链接到自定义函数的 HTML 页。</span><span class="sxs-lookup"><span data-stu-id="306cc-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="306cc-176">仅适用于 Outlook 的扩展点</span><span class="sxs-lookup"><span data-stu-id="306cc-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="306cc-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="306cc-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="306cc-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="306cc-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="306cc-181">[Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）</span><span class="sxs-lookup"><span data-stu-id="306cc-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="306cc-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="306cc-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="306cc-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="306cc-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="306cc-185">Events</span><span class="sxs-lookup"><span data-stu-id="306cc-185">Events</span></span>](#events)
- [<span data-ttu-id="306cc-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="306cc-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="306cc-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="306cc-188">This extension point puts buttons in the command surface for the mail read view.</span><span class="sxs-lookup"><span data-stu-id="306cc-188">This extension point puts buttons in the command surface for the mail read view.</span></span> <span data-ttu-id="306cc-189">In Outlook desktop, this appears in the ribbon.</span><span class="sxs-lookup"><span data-stu-id="306cc-189">In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="306cc-190">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-190">Child elements</span></span>

|  <span data-ttu-id="306cc-191">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-191">Element</span></span> |  <span data-ttu-id="306cc-192">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="306cc-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="306cc-194">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="306cc-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="306cc-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="306cc-196">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="306cc-197">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="306cc-198">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="306cc-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="306cc-200">此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。</span><span class="sxs-lookup"><span data-stu-id="306cc-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="306cc-201">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-201">Child elements</span></span>

|  <span data-ttu-id="306cc-202">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-202">Element</span></span> |  <span data-ttu-id="306cc-203">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="306cc-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="306cc-205">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="306cc-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="306cc-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="306cc-207">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="306cc-208">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="306cc-209">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="306cc-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="306cc-211">此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="306cc-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="306cc-212">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-212">Child elements</span></span>

|  <span data-ttu-id="306cc-213">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-213">Element</span></span> |  <span data-ttu-id="306cc-214">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="306cc-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="306cc-216">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="306cc-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="306cc-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="306cc-218">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="306cc-219">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="306cc-220">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="306cc-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="306cc-222">此扩展点将按钮置于向会议与会者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="306cc-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="306cc-223">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-223">Child elements</span></span>

|  <span data-ttu-id="306cc-224">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-224">Element</span></span> |  <span data-ttu-id="306cc-225">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="306cc-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="306cc-227">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="306cc-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="306cc-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="306cc-229">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="306cc-230">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="306cc-231">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="306cc-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="306cc-232">Module</span><span class="sxs-lookup"><span data-stu-id="306cc-232">Module</span></span>

<span data-ttu-id="306cc-233">此扩展点将按钮置于模块扩展的功能区上。</span><span class="sxs-lookup"><span data-stu-id="306cc-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="306cc-234">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-234">Child elements</span></span>

|  <span data-ttu-id="306cc-235">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-235">Element</span></span> |  <span data-ttu-id="306cc-236">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="306cc-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="306cc-238">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="306cc-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="306cc-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="306cc-240">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="306cc-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="306cc-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="306cc-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="306cc-242">此扩展点将按钮置于移动外形规格中的邮件阅读视图的命令界面中。</span><span class="sxs-lookup"><span data-stu-id="306cc-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="306cc-243">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-243">Child elements</span></span>

|  <span data-ttu-id="306cc-244">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-244">Element</span></span> |  <span data-ttu-id="306cc-245">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-246">Group</span><span class="sxs-lookup"><span data-stu-id="306cc-246">Group</span></span>](group.md) |  <span data-ttu-id="306cc-247">将按钮组添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="306cc-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="306cc-248">此种类型的 **ExtensionPoint** 元素仅能具有一个子元素，即 **Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="306cc-249">此扩展点中包含的 **Control** 元素必须将 **xsi:type** 属性设置为 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="306cc-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="306cc-250">示例</span><span class="sxs-lookup"><span data-stu-id="306cc-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="306cc-251">MobileOnlineMeetingCommandSurface (预览) </span><span class="sxs-lookup"><span data-stu-id="306cc-251">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="306cc-252">仅在使用 Microsoft 365 订阅的 Android 上的[预览](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)中支持此扩展点。</span><span class="sxs-lookup"><span data-stu-id="306cc-252">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="306cc-253">此扩展点在命令界面中为移动外观的约会放置一个适合模式的切换。</span><span class="sxs-lookup"><span data-stu-id="306cc-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="306cc-254">会议组织者可以创建联机会议。</span><span class="sxs-lookup"><span data-stu-id="306cc-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="306cc-255">与会者随后可以加入联机会议。</span><span class="sxs-lookup"><span data-stu-id="306cc-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="306cc-256">若要了解有关此方案的详细信息，请参阅为[联机会议提供商文章创建 Outlook 移动外](../../outlook/online-meeting.md)接程序一文。</span><span class="sxs-lookup"><span data-stu-id="306cc-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="306cc-257">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-257">Child elements</span></span>

|  <span data-ttu-id="306cc-258">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-258">Element</span></span> |  <span data-ttu-id="306cc-259">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-260">Control</span><span class="sxs-lookup"><span data-stu-id="306cc-260">Control</span></span>](control.md) |  <span data-ttu-id="306cc-261">将按钮添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="306cc-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="306cc-262">`ExtensionPoint`此类型的元素只能有一个子元素：一个 `Control` 元素。</span><span class="sxs-lookup"><span data-stu-id="306cc-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="306cc-263">`Control`此扩展点中包含的元素的属性必须 `xsi:type` 设置为 `MobileButton` 。</span><span class="sxs-lookup"><span data-stu-id="306cc-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="306cc-264">`Icon`图像应使用十六进制代码 `#919191` 或以[其他颜色格式](https://convertingcolors.com/hex-color-919191.html)的等效项进行灰度。</span><span class="sxs-lookup"><span data-stu-id="306cc-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="306cc-265">示例</span><span class="sxs-lookup"><span data-stu-id="306cc-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="306cc-266">LaunchEvent (预览) </span><span class="sxs-lookup"><span data-stu-id="306cc-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="306cc-267">仅在使用 Microsoft 365 订阅的 Outlook 网页[预览版](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)中支持此扩展点。</span><span class="sxs-lookup"><span data-stu-id="306cc-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="306cc-268">此扩展点使外接程序能够根据桌面外形规格中受支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="306cc-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="306cc-269">目前，唯一受支持的事件是 `OnNewMessageCompose` 和 `OnNewAppointmentOrganizer` 。</span><span class="sxs-lookup"><span data-stu-id="306cc-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="306cc-270">若要了解有关此方案的详细信息，请参阅[Configure The Outlook 外接程序以获取基于事件的激活一](../../outlook/autolaunch.md)文。</span><span class="sxs-lookup"><span data-stu-id="306cc-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="306cc-271">子元素</span><span class="sxs-lookup"><span data-stu-id="306cc-271">Child elements</span></span>

|  <span data-ttu-id="306cc-272">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-272">Element</span></span> |  <span data-ttu-id="306cc-273">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="306cc-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="306cc-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="306cc-275">基于事件的激活的[LaunchEvent](launchevent.md)列表。</span><span class="sxs-lookup"><span data-stu-id="306cc-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="306cc-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="306cc-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="306cc-277">源 JavaScript 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="306cc-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="306cc-278">示例</span><span class="sxs-lookup"><span data-stu-id="306cc-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="306cc-279">事件</span><span class="sxs-lookup"><span data-stu-id="306cc-279">Events</span></span>

<span data-ttu-id="306cc-280">此扩展点添加了指定事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="306cc-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="306cc-281">有关使用此扩展点的详细信息，请参阅[On a send feature For Outlook 外接程序](../../outlook/outlook-on-send-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="306cc-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="306cc-282">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-282">Element</span></span> | <span data-ttu-id="306cc-283">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-284">Event</span><span class="sxs-lookup"><span data-stu-id="306cc-284">Event</span></span>](event.md) |  <span data-ttu-id="306cc-285">指定事件和事件处理程序函数。</span><span class="sxs-lookup"><span data-stu-id="306cc-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="306cc-286">ItemSend 事件示例</span><span class="sxs-lookup"><span data-stu-id="306cc-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="306cc-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="306cc-287">DetectedEntity</span></span>

<span data-ttu-id="306cc-288">此扩展点在指定实体类型上添加上下文外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="306cc-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="306cc-289">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="306cc-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="306cc-290">此元素类型适用于[支持要求集 1.6 和更高版本的 Outlook 客户端](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="306cc-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="306cc-291">元素</span><span class="sxs-lookup"><span data-stu-id="306cc-291">Element</span></span> |  <span data-ttu-id="306cc-292">说明</span><span class="sxs-lookup"><span data-stu-id="306cc-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="306cc-293">Label</span><span class="sxs-lookup"><span data-stu-id="306cc-293">Label</span></span>](#label) |  <span data-ttu-id="306cc-294">在上下文窗口中指定外接程序的标签。</span><span class="sxs-lookup"><span data-stu-id="306cc-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="306cc-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="306cc-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="306cc-296">指定上下文窗口的 URL。</span><span class="sxs-lookup"><span data-stu-id="306cc-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="306cc-297">Rule</span><span class="sxs-lookup"><span data-stu-id="306cc-297">Rule</span></span>](rule.md) |  <span data-ttu-id="306cc-298">指定确定外接程序激活时间的一个或多个规则。</span><span class="sxs-lookup"><span data-stu-id="306cc-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="306cc-299">标签</span><span class="sxs-lookup"><span data-stu-id="306cc-299">Label</span></span>

<span data-ttu-id="306cc-300">必需。</span><span class="sxs-lookup"><span data-stu-id="306cc-300">Required.</span></span> <span data-ttu-id="306cc-301">组的标签。</span><span class="sxs-lookup"><span data-stu-id="306cc-301">The label of the group.</span></span> <span data-ttu-id="306cc-302">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="306cc-302">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="306cc-303">突出显示要求</span><span class="sxs-lookup"><span data-stu-id="306cc-303">Highlight requirements</span></span>

<span data-ttu-id="306cc-304">The only way a user can activate a contextual add-in is to interact with a highlighted entity.</span><span class="sxs-lookup"><span data-stu-id="306cc-304">The only way a user can activate a contextual add-in is to interact with a highlighted entity.</span></span> <span data-ttu-id="306cc-305">Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span><span class="sxs-lookup"><span data-stu-id="306cc-305">Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="306cc-306">However, there are some limitations to be aware of.</span><span class="sxs-lookup"><span data-stu-id="306cc-306">However, there are some limitations to be aware of.</span></span> <span data-ttu-id="306cc-307">These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span><span class="sxs-lookup"><span data-stu-id="306cc-307">These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="306cc-308">无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="306cc-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="306cc-309">如果使用单个规则，`Highlight` 必须设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="306cc-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="306cc-310">如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="306cc-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="306cc-311">如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="306cc-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="306cc-312">DetectedEntity 事件示例</span><span class="sxs-lookup"><span data-stu-id="306cc-312">DetectedEntity event example</span></span>

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
