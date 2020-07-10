---
title: 在 Outlook 外接程序中实现可固定的任务窗格
description: 用于加载项命令的任务窗格用户体验形状会在打开的邮件或会议请求的右侧打开一个垂直任务窗格，以便用户可以在加载项 UI 中进行更详细的交互。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39af3a532d553835b02709301c998a78dc9958bb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093866"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="cdb65-103">在 Outlook 中实现可固定的任务窗格</span><span class="sxs-lookup"><span data-stu-id="cdb65-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="cdb65-104">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.).</span><span class="sxs-lookup"><span data-stu-id="cdb65-104">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.).</span></span> <span data-ttu-id="cdb65-105">This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span><span class="sxs-lookup"><span data-stu-id="cdb65-105">This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="cdb65-106">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed.</span><span class="sxs-lookup"><span data-stu-id="cdb65-106">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed.</span></span> <span data-ttu-id="cdb65-107">For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message.</span><span class="sxs-lookup"><span data-stu-id="cdb65-107">For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message.</span></span> <span data-ttu-id="cdb65-108">With pinnable task panes, your add-in can give the user that option.</span><span class="sxs-lookup"><span data-stu-id="cdb65-108">With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="cdb65-109">尽管可固定任务窗格功能是在[要求集 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入的，但它目前仅可供 Microsoft 365 订阅者使用，如下所示。</span><span class="sxs-lookup"><span data-stu-id="cdb65-109">Although the pinnable task panes feature was introduced in [requirement set 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only available to Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="cdb65-110">Outlook 2016 或更高版本在 Windows (内部版本或 Office 预览体验部门的用户的内部7668.2000 版本中，为延迟频道中的用户构建7900或更高版本) </span><span class="sxs-lookup"><span data-stu-id="cdb65-110">Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels)</span></span>
> - <span data-ttu-id="cdb65-111">Outlook 2016 或更高版本 Mac 版 (16.13.503 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="cdb65-111">Outlook 2016 or later on Mac (version 16.13.503 or later)</span></span>
> - <span data-ttu-id="cdb65-112">新式 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="cdb65-112">Modern Outlook on the web</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cdb65-113">可固定任务窗格不能用于以下对象。</span><span class="sxs-lookup"><span data-stu-id="cdb65-113">Pinnable task panes are not available for the following.</span></span>
> - <span data-ttu-id="cdb65-114">约会/会议</span><span class="sxs-lookup"><span data-stu-id="cdb65-114">Appointments/Meetings</span></span>
> - <span data-ttu-id="cdb65-115">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="cdb65-115">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="cdb65-116">支持固定任务窗格</span><span class="sxs-lookup"><span data-stu-id="cdb65-116">Support task pane pinning</span></span>

<span data-ttu-id="cdb65-117">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="cdb65-117">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md).</span></span> <span data-ttu-id="cdb65-118">This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span><span class="sxs-lookup"><span data-stu-id="cdb65-118">This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="cdb65-119">由于 `SupportsPinning` 元素是在 VersionOverrides v1.1 架构中进行定义，因此需为 v1.0 和 v1.1 架构包含 [VersionOverrides](../reference/manifest/versionoverrides.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="cdb65-119">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="cdb65-120">如果计划将 Outlook 加载项[发布](../publish/publish.md)到 [AppSource](https://appsource.microsoft.com)，那么在使用 **SupportsPinning** 元素时，加载项内容不得为静态，且必须清晰显示与邮箱中打开或选择的邮件相关的数据，才能通过 [AppSource 验证](/legal/marketplace/certification-policies)。</span><span class="sxs-lookup"><span data-stu-id="cdb65-120">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/legal/marketplace/certification-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

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

<span data-ttu-id="cdb65-121">有关完整示例，请参阅[command-demo 示例清单](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml)中的 `msgReadOpenPaneButton` 控件。</span><span class="sxs-lookup"><span data-stu-id="cdb65-121">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="cdb65-122">根据当前选择的邮件处理 UI 更新</span><span class="sxs-lookup"><span data-stu-id="cdb65-122">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="cdb65-123">若要根据当前项更新任务窗格的 UI 或内部变量，必须注册事件处理程序，才能收到变化通知。</span><span class="sxs-lookup"><span data-stu-id="cdb65-123">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="cdb65-124">实现事件处理程序</span><span class="sxs-lookup"><span data-stu-id="cdb65-124">Implement the event handler</span></span>

<span data-ttu-id="cdb65-125">The event handler should accept a single parameter, which is an object literal.</span><span class="sxs-lookup"><span data-stu-id="cdb65-125">The event handler should accept a single parameter, which is an object literal.</span></span> <span data-ttu-id="cdb65-126">The `type` property of this object will be set to `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="cdb65-126">The `type` property of this object will be set to `Office.EventType.ItemChanged`.</span></span> <span data-ttu-id="cdb65-127">When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span><span class="sxs-lookup"><span data-stu-id="cdb65-127">When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="cdb65-128">ItemChanged 事件的事件处理程序的实现应检查 Office.content.mailbox.item 是否为 NULL。</span><span class="sxs-lookup"><span data-stu-id="cdb65-128">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="cdb65-129">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="cdb65-129">Register the event handler</span></span>

<span data-ttu-id="cdb65-130">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event.</span><span class="sxs-lookup"><span data-stu-id="cdb65-130">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event.</span></span> <span data-ttu-id="cdb65-131">This should be done in the `Office.initialize` function for your task pane.</span><span class="sxs-lookup"><span data-stu-id="cdb65-131">This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="cdb65-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cdb65-132">See also</span></span>

<span data-ttu-id="cdb65-133">有关实现可固定的任务窗格的示例外接程序，请参阅 GitHub 上的 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。</span><span class="sxs-lookup"><span data-stu-id="cdb65-133">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
