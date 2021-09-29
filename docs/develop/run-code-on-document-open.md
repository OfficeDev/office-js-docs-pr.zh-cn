---
title: 文档打开时在 Office 加载项中运行代码
description: 了解如何在文档打开时Office外接程序外接程序中运行代码。
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 884409fb161970c57b32921192544592ca39bb2c
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990549"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>文档打开时在 Office 加载项中运行代码

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

您可以将您的Office配置为在打开文档后加载和运行代码。 如果您需要在加载项可见之前注册事件处理程序、预加载任务窗格数据、同步 UI 或执行其他任务，这将非常有用。

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>将外接程序配置为在文档打开时加载

以下代码将外接程序配置为在打开文档时加载并开始运行。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> `setStartupBehavior`方法是异步的。

## <a name="place-startup-code-in-officeinitialize"></a>将启动代码放在 Office.initialize 中

当外接程序配置为在打开文档时加载时，它将立即运行。 `Office.initialize`将调用事件处理程序。 将启动代码放在 `Office.initialize` 或 `Office.onReady` 事件处理程序中。

以下Excel代码展示了如何为活动工作表中的更改事件注册事件处理程序。 如果您将外接程序配置为在文档打开时加载，则此代码将在文档打开时注册事件处理程序。 您可以在打开任务窗格之前处理更改事件。

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}
```

下面的PowerPoint代码展示了如何为文档文档中的选择更改事件注册PowerPoint处理程序。 如果您将外接程序配置为在文档打开时加载，则此代码将在文档打开时注册事件处理程序。 您可以在打开任务窗格之前处理更改事件。

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>在打开文档时将外接程序配置为无加载行为

以下代码将外接程序配置为在打开文档时不启动。 相反，它将在用户以某种方式参与时启动，例如选择功能区按钮或打开任务窗格。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>获取当前加载行为

若要确定当前启动行为是什么，请运行以下函数，该函数返回 `Office.StartupBehavior` 对象。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>另请参阅

- [将 Office 加载项配置为使用共享 JavaScript 运行时](configure-your-add-in-to-use-a-shared-runtime.md)
- [在自定义函数和任务Excel之间共享数据和事件教程](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [使用 Excel JavaScript API 处理事件](../excel/excel-add-ins-events.md)
