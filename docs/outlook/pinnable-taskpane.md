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
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>在 Outlook 中实现可固定的任务窗格

The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> 尽管可固定任务窗格功能是在[要求集 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入的，但它目前仅可供 Microsoft 365 订阅者使用，如下所示。
> - Outlook 2016 或更高版本在 Windows (内部版本或 Office 预览体验部门的用户的内部7668.2000 版本中，为延迟频道中的用户构建7900或更高版本) 
> - Outlook 2016 或更高版本 Mac 版 (16.13.503 或更高版本) 
> - 新式 Outlook 网页版

> [!IMPORTANT]
> 可固定任务窗格不能用于以下对象。
> - 约会/会议
> - Outlook.com

## <a name="support-task-pane-pinning"></a>支持固定任务窗格

The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.

由于 `SupportsPinning` 元素是在 VersionOverrides v1.1 架构中进行定义，因此需为 v1.0 和 v1.1 架构包含 [VersionOverrides](../reference/manifest/versionoverrides.md) 元素。

> [!NOTE]
> 如果计划将 Outlook 加载项[发布](../publish/publish.md)到 [AppSource](https://appsource.microsoft.com)，那么在使用 **SupportsPinning** 元素时，加载项内容不得为静态，且必须清晰显示与邮箱中打开或选择的邮件相关的数据，才能通过 [AppSource 验证](/legal/marketplace/certification-policies)。

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

有关完整示例，请参阅[command-demo 示例清单](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml)中的 `msgReadOpenPaneButton` 控件。

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>根据当前选择的邮件处理 UI 更新

若要根据当前项更新任务窗格的 UI 或内部变量，必须注册事件处理程序，才能收到变化通知。

### <a name="implement-the-event-handler"></a>实现事件处理程序

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> ItemChanged 事件的事件处理程序的实现应检查 Office.content.mailbox.item 是否为 NULL。
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a>注册事件处理程序

Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a>另请参阅

有关实现可固定的任务窗格的示例外接程序，请参阅 GitHub 上的 [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo)。
