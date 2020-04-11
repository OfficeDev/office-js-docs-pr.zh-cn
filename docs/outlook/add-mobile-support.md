---
title: 添加对 Outlook 外接程序的移动支持
description: 添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: de5f1209527c853afb2d0bf2061bd3e3cfa8d3e0
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/10/2020
ms.locfileid: "43225664"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a><span data-ttu-id="11c9a-103">添加对适用于 Outlook Mobile 的外接程序命令的支持</span><span class="sxs-lookup"><span data-stu-id="11c9a-103">Add support for add-in commands for Outlook Mobile</span></span>

<span data-ttu-id="11c9a-104">使用 Outlook Mobile 中的外接程序命令，用户可以访问在 web、Windows 和 Mac 上的 Outlook 中已有的相同功能（有一些[限制](#code-considerations)）。</span><span class="sxs-lookup"><span data-stu-id="11c9a-104">Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="11c9a-105">添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。</span><span class="sxs-lookup"><span data-stu-id="11c9a-105">Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.</span></span>

## <a name="updating-the-manifest"></a><span data-ttu-id="11c9a-106">更新清单</span><span class="sxs-lookup"><span data-stu-id="11c9a-106">Updating the manifest</span></span>

<span data-ttu-id="11c9a-p102">启用 Outlook Mobile 中的外接程序命令的第一步是在外接程序清单中对其进行定义。[VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 架构为移动电话定义新的外形规格，[MobileFormFactor](../reference/manifest/mobileformfactor.md)。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p102">The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span></span>

<span data-ttu-id="11c9a-p103">此元素包含在移动客户端中加载外接程序所需的所有信息。这使你可以为移动体验定义完全不同的 UI 元素和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p103">This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.</span></span>

<span data-ttu-id="11c9a-111">下面的示例显示`MobileFormFactor`元素中的单个任务窗格按钮。</span><span class="sxs-lookup"><span data-stu-id="11c9a-111">The following example shows a single task pane button in a `MobileFormFactor` element.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

<span data-ttu-id="11c9a-112">这与 [DesktopFormFactor](../reference/manifest/desktopformfactor.md) 元素中出现的元素非常相似，但具有一些明显的区别。</span><span class="sxs-lookup"><span data-stu-id="11c9a-112">This is very similar to the elements that appear in a [DesktopFormFactor](../reference/manifest/desktopformfactor.md) element, with some notable differences.</span></span>

- <span data-ttu-id="11c9a-113">不使用 [OfficeTab](../reference/manifest/officetab.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="11c9a-113">The [OfficeTab](../reference/manifest/officetab.md) element is not used.</span></span>
- <span data-ttu-id="11c9a-p104">[ExtensionPoint](../reference/manifest/extensionpoint.md) 元素必须仅具有一个子元素。如果外接程序仅添加一个按钮，则子元素应为 [Control](../reference/manifest/control.md) 元素。如果外接程序添加多个按钮，则子元素应为包含多个 `Control` 元素的 [Group](../reference/manifest/group.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p104">The [ExtensionPoint](../reference/manifest/extensionpoint.md) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](../reference/manifest/control.md) element. If the add-in adds more than one button, the child element should be a [Group](../reference/manifest/group.md) element that contains multiple `Control` elements.</span></span>
- <span data-ttu-id="11c9a-117">没有与 `Control` 元素等效的 `Menu` 类型。</span><span class="sxs-lookup"><span data-stu-id="11c9a-117">There is no `Menu` type equivalent for the `Control` element.</span></span>
- <span data-ttu-id="11c9a-118">不使用 [Supertip](../reference/manifest/supertip.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="11c9a-118">The [Supertip](../reference/manifest/supertip.md) element is not used.</span></span>
- <span data-ttu-id="11c9a-p105">要求的图标大小不同。移动外接程序最少必须支持 25x25、32x32 和 48x48 像素的图标。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p105">The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.</span></span>

## <a name="code-considerations"></a><span data-ttu-id="11c9a-121">代码注意事项</span><span class="sxs-lookup"><span data-stu-id="11c9a-121">Code considerations</span></span>

<span data-ttu-id="11c9a-122">设计适用于移动电话的外接程序引入了一些额外注意事项。</span><span class="sxs-lookup"><span data-stu-id="11c9a-122">Designing an add-in for mobile introduces some additional considerations.</span></span>

### <a name="use-rest-instead-of-exchange-web-services"></a><span data-ttu-id="11c9a-123">使用 REST 代替 Exchange Web 服务</span><span class="sxs-lookup"><span data-stu-id="11c9a-123">Use REST instead of Exchange Web Services</span></span>

<span data-ttu-id="11c9a-p106">Outlook Mobile 中不支持 [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法。外接程序应在可能的情况下首选从 Office.js API 获取信息。如果外接程序需要 Office.js API 未公开的信息，则应使用 [Outlook REST API](/outlook/rest/) 来访问用户邮箱。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p106">The [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.</span></span>

<span data-ttu-id="11c9a-127">邮箱要求集1.5 引入了新版本的[mailbox.getcallbacktokenasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) ，该版本可以请求与 REST api 兼容的访问令牌，以及可用于查找用户的 rest api 终结点的新的[office.context.mailbox.resturl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)属性。</span><span class="sxs-lookup"><span data-stu-id="11c9a-127">Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property that can be used to find the REST API endpoint for the user.</span></span>

### <a name="pinch-zoom"></a><span data-ttu-id="11c9a-128">收缩缩放</span><span class="sxs-lookup"><span data-stu-id="11c9a-128">Pinch zoom</span></span>

<span data-ttu-id="11c9a-p107">在默认情况下，用户可以使用“收缩缩放”手势在任务窗格上进行缩放。如果方案不需要该功能，请确保在 HTML 中禁用收缩缩放。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p107">By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.</span></span>

### <a name="close-task-panes"></a><span data-ttu-id="11c9a-131">关闭任务窗格</span><span class="sxs-lookup"><span data-stu-id="11c9a-131">Close task panes</span></span>

<span data-ttu-id="11c9a-p108">在 Outlook Mobile 中，任务窗格占据整个屏幕，并且在默认情况下需要用户将其关闭以返回到邮件。请考虑使用 [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) 方法在方案完成时关闭任务窗格。</span><span class="sxs-lookup"><span data-stu-id="11c9a-p108">In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) method to close the task pane when your scenario is complete.</span></span>

### <a name="compose-mode-and-appointments"></a><span data-ttu-id="11c9a-134">撰写模式和约会</span><span class="sxs-lookup"><span data-stu-id="11c9a-134">Compose mode and appointments</span></span>

<span data-ttu-id="11c9a-135">目前，Outlook Mobile 中的外接程序仅在读取邮件时支持激活。</span><span class="sxs-lookup"><span data-stu-id="11c9a-135">Currently add-ins in Outlook Mobile only support activation when reading messages.</span></span> <span data-ttu-id="11c9a-136">在撰写邮件时或查看或撰写约会时，不会激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="11c9a-136">Add-ins are not activated when composing messages or when viewing or composing appointments.</span></span> <span data-ttu-id="11c9a-137">但是，可以在约会组织者模式下激活联机会议提供程序集成的外接程序。</span><span class="sxs-lookup"><span data-stu-id="11c9a-137">However, online meeting provider integrated add-ins can be activated in Appointment Organizer mode.</span></span> <span data-ttu-id="11c9a-138">有关此异常的详细信息，请参阅[创建适用于联机会议提供商文章的 Outlook mobile 外](online-meeting.md)接程序。</span><span class="sxs-lookup"><span data-stu-id="11c9a-138">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this exception.</span></span>

### <a name="unsupported-apis"></a><span data-ttu-id="11c9a-139">不支持的 API</span><span class="sxs-lookup"><span data-stu-id="11c9a-139">Unsupported APIs</span></span>

<span data-ttu-id="11c9a-140">Outlook Mobile 不支持在要求集1.6 或更高版本中引入的 Api。</span><span class="sxs-lookup"><span data-stu-id="11c9a-140">APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile.</span></span> <span data-ttu-id="11c9a-141">此外，还不支持来自早期要求集的以下 Api。</span><span class="sxs-lookup"><span data-stu-id="11c9a-141">The following APIs from earlier requirement sets are also not supported.</span></span>

  - [<span data-ttu-id="11c9a-142">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="11c9a-142">Office.context.officeTheme</span></span>](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [<span data-ttu-id="11c9a-143">Office.context.mailbox.ewsUrl</span><span class="sxs-lookup"><span data-stu-id="11c9a-143">Office.context.mailbox.ewsUrl</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [<span data-ttu-id="11c9a-144">Office.context.mailbox.convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="11c9a-144">Office.context.mailbox.convertToEwsId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="11c9a-145">Office.context.mailbox.convertToRestId</span><span class="sxs-lookup"><span data-stu-id="11c9a-145">Office.context.mailbox.convertToRestId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="11c9a-146">Office.context.mailbox.displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="11c9a-146">Office.context.mailbox.displayAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="11c9a-147">Office.context.mailbox.displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="11c9a-147">Office.context.mailbox.displayMessageForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="11c9a-148">Office.context.mailbox.displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="11c9a-148">Office.context.mailbox.displayNewAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="11c9a-149">Office.context.mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="11c9a-149">Office.context.mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [<span data-ttu-id="11c9a-150">Office.context.mailbox.item.dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="11c9a-150">Office.context.mailbox.item.dateTimeModified</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [<span data-ttu-id="11c9a-151">Office.context.mailbox.item.displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="11c9a-151">Office.context.mailbox.item.displayReplyAllForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="11c9a-152">Office.context.mailbox.item.displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="11c9a-152">Office.context.mailbox.item.displayReplyForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="11c9a-153">Office.context.mailbox.item.getEntities</span><span class="sxs-lookup"><span data-stu-id="11c9a-153">Office.context.mailbox.item.getEntities</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="11c9a-154">Office.context.mailbox.item.getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="11c9a-154">Office.context.mailbox.item.getEntitiesByType</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="11c9a-155">Office.context.mailbox.item.getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="11c9a-155">Office.context.mailbox.item.getFilteredEntitiesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="11c9a-156">Office.context.mailbox.item.getRegexMatches</span><span class="sxs-lookup"><span data-stu-id="11c9a-156">Office.context.mailbox.item.getRegexMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [<span data-ttu-id="11c9a-157">Office.context.mailbox.item.getRegexMatchesByName</span><span class="sxs-lookup"><span data-stu-id="11c9a-157">Office.context.mailbox.item.getRegexMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a><span data-ttu-id="11c9a-158">另请参阅</span><span class="sxs-lookup"><span data-stu-id="11c9a-158">See also</span></span>

[<span data-ttu-id="11c9a-159">要求集支持</span><span class="sxs-lookup"><span data-stu-id="11c9a-159">Requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)