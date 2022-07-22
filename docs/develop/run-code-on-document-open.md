---
title: 文档打开时在 Office 加载项中运行代码
description: 了解如何在打开文档时在 Office 加载项加载项中运行代码。
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a1c3277a349dc4054da5f089c62331296590021
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958437"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>文档打开时在 Office 加载项中运行代码

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

可将 Office 加载项配置为在打开文档后立即加载和运行代码。 如果需要注册事件处理程序、为任务窗格预加载数据、同步 UI 或在外接程序可见之前执行其他任务，这非常有用。

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>将加载项配置为在文档打开时加载

以下代码将加载项配置为加载并在打开文档时开始运行。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> 该 `setStartupBehavior` 方法是异步的。

## <a name="place-startup-code-in-officeinitialize"></a>将启动代码置于 Office.initialize 中

将外接程序配置为在打开文档时加载时，它将立即运行。 `Office.initialize`将调用事件处理程序。 将启动代码放在或`Office.onReady`事件处理程序中`Office.initialize`。

以下 Excel 外接程序代码演示如何注册活动工作表中更改事件的事件处理程序。 如果将外接程序配置为在打开文档时加载，则此代码将在打开文档时注册事件处理程序。 可以在打开任务窗格之前处理更改事件。

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
    await Excel.run(async (context) => {    
        await context.sync();
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
  });
}
```

下面的 PowerPoint 外接程序代码演示如何为 PowerPoint 文档中的选择更改事件注册事件处理程序。 如果将外接程序配置为在打开文档时加载，则此代码将在打开文档时注册事件处理程序。 可以在打开任务窗格之前处理更改事件。

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

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>为打开文档时没有负载行为配置外接程序

以下代码配置加载项在打开文档时不启动。 相反，它会在用户以某种方式参与时启动，例如选择功能区按钮或打开任务窗格。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>获取当前加载行为

若要确定当前启动行为是什么，请运行以下方法，该方法返回一个 `Office.StartupBehavior` 对象。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>另请参阅

- [将 Office 加载项配置为使用共享 JavaScript 运行时](configure-your-add-in-to-use-a-shared-runtime.md)
- [在 Excel 自定义函数和任务窗格教程之间共享数据和事件](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [使用 Excel JavaScript API 处理事件](../excel/excel-add-ins-events.md)
