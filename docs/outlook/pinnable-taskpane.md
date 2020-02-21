---
title: 在 Outlook 外接程序中实现可固定的任务窗格
description: 用于加载项命令的任务窗格用户体验形状会在打开的邮件或会议请求的右侧打开一个垂直任务窗格，以便用户可以在加载项 UI 中进行更详细的交互。
ms.date: 11/18/2019
localization_priority: Normal
ms.openlocfilehash: 94c136a74dfddac1af663aea06c3c6ca27f22dcd
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165895"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="8b423-103">在 Outlook 中实现可固定的任务窗格</span><span class="sxs-lookup"><span data-stu-id="8b423-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="8b423-p101">用于加载项命令的[任务窗格](add-in-commands-for-outlook.md#launching-a-task-pane)用户体验形状会在打开的邮件或会议请求的右侧打开一个垂直任务窗格，以便用户可以在加载项 UI 中进行更详细的交互（填充多个字段等）。查看邮件列表时，可以在阅读窗格中看到此任务窗格，从而能够快速处理邮件。</span><span class="sxs-lookup"><span data-stu-id="8b423-p101">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="8b423-p102">不过，默认情况下，如果用户在阅读窗格中为某封邮件打开了外接程序任务窗格，然后选择新邮件，此任务窗格会自动关闭。如果频繁使用外接程序，用户可能更倾向于让此任务窗格一直处于打开状态，这样就无需在每封邮件中都重新激活外接程序了。使用可固定的任务窗格，外接程序就可以让用户如愿以偿。</span><span class="sxs-lookup"><span data-stu-id="8b423-p102">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="8b423-109">目前，使用 Windows 版 Outlook 2016 或更高版本（适用于当前频道或 Office 预览体验计划频道用户的内部版本 7668.2000 或更高版本，适用于延期频道用户的内部版本 7900.xxxx 或更高版本）、Mac 版 Outlook 2016 或更高版本（版本 16.13.503 或更高版本）以及 Outlook 网页版的 Office 365 订阅者可使用可固定的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="8b423-109">Pinnable task panes are currently available to Office 365 subscribers using Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels), Outlook 2016 or later on Mac (version 16.13.503 or later), and Outlook on the web.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8b423-110">可固定任务窗格不能用于以下对象。</span><span class="sxs-lookup"><span data-stu-id="8b423-110">Pinnable task panes are not available for the following.</span></span>
> - <span data-ttu-id="8b423-111">约会/会议</span><span class="sxs-lookup"><span data-stu-id="8b423-111">Appointments/Meetings</span></span>
> - <span data-ttu-id="8b423-112">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="8b423-112">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="8b423-113">支持固定任务窗格</span><span class="sxs-lookup"><span data-stu-id="8b423-113">Support task pane pinning</span></span>

<span data-ttu-id="8b423-p103">第一步是添加固定支持，此步操作是在外接程序[清单](manifests.md)中完成。为此，请向描述任务窗格按钮的 `Action` 元素添加 [ SupportsPinning](../reference/manifest/action.md#supportspinning) 元素。</span><span class="sxs-lookup"><span data-stu-id="8b423-p103">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="8b423-116">由于 `SupportsPinning` 元素是在 VersionOverrides v1.1 架构中进行定义，因此需为 v1.0 和 v1.1 架构包含 [VersionOverrides](../reference/manifest/versionoverrides.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="8b423-116">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="8b423-117">如果计划将 Outlook 加载项[发布](../publish/publish.md)到 [AppSource](https://appsource.microsoft.com)，那么在使用 **SupportsPinning** 元素时，加载项内容不得为静态，且必须清晰显示与邮箱中打开或选择的邮件相关的数据，才能通过 [AppSource 验证](/office/dev/store/validation-policies)。</span><span class="sxs-lookup"><span data-stu-id="8b423-117">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/office/dev/store/validation-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

<span data-ttu-id="8b423-118">有关完整示例，请参阅[command-demo 示例清单](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml)中的 `msgReadOpenPaneButton` 控件。</span><span class="sxs-lookup"><span data-stu-id="8b423-118">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="8b423-119">根据当前选择的邮件处理 UI 更新</span><span class="sxs-lookup"><span data-stu-id="8b423-119">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="8b423-120">若要根据当前项更新任务窗格的 UI 或内部变量，必须注册事件处理程序，才能收到变化通知。</span><span class="sxs-lookup"><span data-stu-id="8b423-120">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="8b423-121">实现事件处理程序</span><span class="sxs-lookup"><span data-stu-id="8b423-121">Implement the event handler</span></span>

<span data-ttu-id="8b423-p104">事件处理程序应接受一个参数，即对象文本。该对象的 `type` 属性将设为 `Office.EventType.ItemChanged`。事件调用后，`Office.context.mailbox.item` 对象已更新，以反映当前选定项。</span><span class="sxs-lookup"><span data-stu-id="8b423-p104">The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="8b423-125">ItemChanged 事件的事件处理程序的实现应检查 Office.content.mailbox.item 是否为 NULL。</span><span class="sxs-lookup"><span data-stu-id="8b423-125">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="8b423-126">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="8b423-126">Register the event handler</span></span>

<span data-ttu-id="8b423-p105">使用 [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法注册 `Office.EventType.ItemChanged` 事件的事件处理程序。这步操作应在任务窗格的 `Office.initialize` 函数内完成。</span><span class="sxs-lookup"><span data-stu-id="8b423-p105">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="8b423-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8b423-129">See also</span></span>

<span data-ttu-id="8b423-130">有关实现可固定的任务窗格的示例外接程序，请参阅 GitHub 上的 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。</span><span class="sxs-lookup"><span data-stu-id="8b423-130">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
