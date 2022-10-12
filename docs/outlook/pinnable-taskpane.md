---
title: 在 Outlook 外接程序中实现可固定的任务窗格
description: 用于加载项命令的任务窗格用户体验形状会在打开的邮件或会议请求的右侧打开一个垂直任务窗格，以便用户可以在加载项 UI 中进行更详细的交互。
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 834d43a6046ddaa63a7c8899cfd5b07d0ea80ef6
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541120"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>在 Outlook 中实现可固定的任务窗格

The [task pane](add-in-commands-for-outlook.md#launch-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> 尽管 [在要求集 1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5) 中引入了可固定任务窗格功能，但目前仅适用于 Microsoft 365 订阅者，使用以下命令：
>
> - Windows 版 Outlook 2016 或更高版本（适用于当前或 Office 预览体验计划频道中的用户的内部版本 7668.2000 或更高版本，适用于延期频道中的用户的内部版本 7900.xxxx 或更高版本）
> - Mac 版 Outlook 2016 或更高版本（版本 16.13.503 或更高版本）
> - 新式 Outlook 网页版

> [!IMPORTANT]
> 可固定任务窗格不适用于以下内容：
>
> - 约会/会议
> - Outlook.com

> [!TIP]
> 如果计划将 Outlook 外接程序 [发布](../publish/publish.md) 到 [AppSource](https://appsource.microsoft.com)，并且它已配置为可固定任务窗格，为了传递 [AppSource 验证](/legal/marketplace/certification-policies)，外接程序内容不得为静态内容，并且必须清楚地显示与邮箱中打开或选择的消息相关的数据。

## <a name="support-task-pane-pinning"></a>支持固定任务窗格

第一步是添加固定支持，此步操作是在外接程序清单中完成。 标记因清单类型而异。

# <a name="xml-manifest"></a>[XML 清单](#tab/xmlmanifest)

将 [SupportsPinning](/javascript/api/manifest/action#supportspinning) 元素添加到 **\<Action\>** 描述任务窗格按钮的元素。 示例如下。

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

该 **\<SupportsPinning\>** 元素在 VersionOverrides v1.1 架构中定义，因此需要包含适用于 v1.0 和 v1.1 的 [VersionOverrides](/javascript/api/manifest/versionoverrides) 元素。

# <a name="teams-manifest-developer-preview"></a>[Teams 清单 (开发人员预览) ](#tab/jsonmanifest)

将设置为的“可固定”属性添加到 `true`定义打开任务窗格的按钮或菜单项的“操作”数组中的对象。 示例如下。

```json
"actions": [
    {
        "id": "OpenTaskPane",
        "type": "openPage",
        "view": "TaskPaneView",
        "displayName": "OpenTaskPane",
        "pinnable": true
    }
]
```

---

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

Use the [Office.context.mailbox.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

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
