---
title: 在 Outlook 外接程序中实现可固定的任务窗格
description: 用于加载项命令的任务窗格用户体验形状会在打开的邮件或会议请求的右侧打开一个垂直任务窗格，以便用户可以在加载项 UI 中进行更详细的交互。
ms.date: 02/28/2020
localization_priority: Normal
ms.openlocfilehash: 892cee7b5ff89e210c68308f03710ee92b6f0f72
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890989"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>在 Outlook 中实现可固定的任务窗格

用于加载项命令的[任务窗格](add-in-commands-for-outlook.md#launching-a-task-pane)用户体验形状会在打开的邮件或会议请求的右侧打开一个垂直任务窗格，以便用户可以在加载项 UI 中进行更详细的交互（填充多个字段等）。查看邮件列表时，可以在阅读窗格中看到此任务窗格，从而能够快速处理邮件。

不过，默认情况下，如果用户在阅读窗格中为某封邮件打开了外接程序任务窗格，然后选择新邮件，此任务窗格会自动关闭。如果频繁使用外接程序，用户可能更倾向于让此任务窗格一直处于打开状态，这样就无需在每封邮件中都重新激活外接程序了。使用可固定的任务窗格，外接程序就可以让用户如愿以偿。

> [!NOTE]
> 尽管可固定任务窗格功能是在[要求集 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入的，但它目前仅适用于 Office 365 订阅者（使用以下程序）。
> - Outlook 2016 或更高版本的 Windows （内部版本7668.2000 或更高版本）对于当前或 Office 预览体验成员的用户，为延迟频道中的用户构建7900或更高版本
> - Mac 上的 Outlook 2016 或更高版本（版本16.13.503 或更高版本）
> - 新式 Outlook 网页版

> [!IMPORTANT]
> 可固定任务窗格不能用于以下对象。
> - 约会/会议
> - Outlook.com

## <a name="support-task-pane-pinning"></a>支持固定任务窗格

第一步是添加固定支持，此步操作是在外接程序[清单](manifests.md)中完成。为此，请向描述任务窗格按钮的 `Action` 元素添加 [ SupportsPinning](../reference/manifest/action.md#supportspinning) 元素。

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

事件处理程序应接受一个参数，即对象文本。该对象的 `type` 属性将设为 `Office.EventType.ItemChanged`。事件调用后，`Office.context.mailbox.item` 对象已更新，以反映当前选定项。

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

使用 [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法注册 `Office.EventType.ItemChanged` 事件的事件处理程序。这步操作应在任务窗格的 `Office.initialize` 函数内完成。

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
