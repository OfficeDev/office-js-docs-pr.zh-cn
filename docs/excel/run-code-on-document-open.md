---
title: 在文档打开时，在 Excel 外接程序中运行代码（预览）
description: 在文档打开时，在 Excel 外接程序中运行代码。
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 5b8c646a1154540244b1f5e0ac47ad8eaec1801f
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284124"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens-preview"></a>在文档打开时，在 Excel 外接程序中运行代码（预览）

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

您可以将 Excel 加载项配置为在文档打开后立即加载和运行代码。 如果需要注册事件处理程序、预先加载任务窗格的数据、同步 UI 或在外接程序可见之前执行其他任务，这将非常有用。

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>将外接程序配置为在文档打开时加载

下面的代码将加载项配置为在文档打开时加载并开始运行。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> 方法`setStartupBehavior`是异步的。

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>为打开的文档配置无加载行为的外接程序

以下代码将外接程序配置为在文档打开时启动。 而是在用户以某种方式（例如，选择功能区按钮或打开任务窗格）时启动。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>获取当前加载行为

若要确定当前启动行为是什么，请运行以下函数，该函数将返回 StartupBehavior 对象。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>如何在文档打开时运行代码

将外接程序配置为在打开文档时加载时，它将立即运行。 将`Office.initialize`调用事件处理程序。 将启动代码放在`Office.initialize`事件处理程序中。

下面的代码演示如何为活动工作表中的更改事件注册事件处理程序。 如果将加载项配置为在打开文档时加载，此代码将在文档打开时注册事件处理程序。 您可以在打开任务窗格之前处理更改事件。


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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

## <a name="see-also"></a>另请参阅

- [在 Excel 自定义函数和任务窗格教程之间共享数据和事件教程](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)