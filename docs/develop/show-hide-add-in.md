---
title: 显示或隐藏 Office 加载项的任务窗格
description: 了解如何在加载项持续运行时以编程方式隐藏或显示加载项的用户界面。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: b2f0efa95f4ce71fc73d9834cfc165cfdd85dc8f
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773753"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>显示或隐藏 Office 加载项的任务窗格

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

可以通过调用 函数来Office加载项的任务 `Office.addin.showAsTaskpane()` 窗格。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

前面的代码假定存在一个名为 **CurrentQuarterSales** Excel工作表的方案。 只要激活此工作表，加载项就会使任务窗格可见。 方法 `onCurrentQuarter` 是一个处理程序，用于[Office。为工作表注册的 Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onActivated)事件。

您还可以通过调用 函数隐藏任务 `Office.addin.hide()` 窗格。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

前面的代码是一个为应用程序注册的[Office。Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onDeactivated)事件。

## <a name="additional-details-on-showing-the-task-pane"></a>有关显示任务窗格的其他详细信息

调用 时，Office将在任务窗格中显示你分配为资源 ID 的文件 () `Office.addin.showAsTaskpane()` `resid` 任务窗格的值。 此值 `resid` 可通过打开文件并位于 元素manifest.xml来分配 `<SourceLocation>` 或 `<Action xsi:type="ShowTaskpane">` 更改。
 (请参阅[配置Office加载项以使用共享](configure-your-add-in-to-use-a-shared-runtime.md)运行时了解其他详细信息。) 

由于 `Office.addin.showAsTaskpane()` 是异步方法，因此代码将继续运行，直到函数完成。 等待使用关键字或方法完成此操作 `await` `then()` ，具体取决于你使用的 JavaScript 语法。

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>将外接程序配置为使用共享运行时

若要使用 `showAsTaskpane()` 和 `hide()` 方法，加载项必须使用共享运行时。 有关详细信息，请参阅[将Office加载项配置为使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="preservation-of-state-and-event-listeners"></a>状态和事件侦听器的保留

`hide()`和 `showAsTaskpane()` 方法仅 *更改任务* 窗格的可见性。 它们不会卸载或重新加载 (或重新初始化其状态) 。

请考虑以下方案：任务窗格设计有选项卡。 首次 **启动** 加载项时，"主页"选项卡将打开。 假设用户 **打开设置选项卡**，稍后任务窗格中的代码调用以 `hide()` 响应某些事件。 稍后代码调用以 `showAsTaskpane()` 响应另一个事件。 任务窗格将重新出现，并且"设置"选项卡仍处于选中状态。

![包含四个标签为"主页"、"设置、收藏夹"和"帐户"的任务窗格的屏幕截图。](../images/TaskpaneWithTabs.png)

此外，即使任务窗格处于隐藏状态，任务窗格中注册的任何事件侦听器也将继续运行。

请考虑以下方案：任务窗格具有针对名为 Sheet1 的工作表Excel和事件的 `Worksheet.onActivated` 注册 `Worksheet.onDeactivated` **处理程序**。 激活的处理程序导致任务窗格中出现一个绿色点。 停用的处理程序将点红 (，这是其默认状态) 。 假设代码在 Sheet1 未激活且点为 `hide()` 红色时调用。  隐藏任务窗格时 **，Sheet1** 将被激活。 稍后代码调用 `showAsTaskpane()` 以响应某些事件。 任务窗格打开时，该点为绿色，因为即使任务窗格处于隐藏状态，事件侦听器和处理程序也运行。

## <a name="handle-the-visibility-changed-event"></a>处理可见性更改事件

当代码使用 或 更改任务窗格的可见性时，Office `showAsTaskpane()` `hide()` 触发 `VisibilityModeChanged` 事件。 处理此事件可能很有用。 例如，假设任务窗格显示工作簿中所有工作表的列表。 如果在隐藏任务窗格时添加新工作表，则使任务窗格可见本身不会向列表中添加新工作表名称。 但是，您的代码可以响应该事件，以 `VisibilityModeChanged` 重新加载[workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) [Worksheet.name](/javascript/api/excel/excel.worksheet#name)中所有工作表的 Worksheet.name 属性，如下面的示例代码所示。

若要为事件注册处理程序，请不要像在大多数 JavaScript 上下文中一样使用"add handler"Office方法。 相反，有一个特殊的函数要传递给你的处理程序[：Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onVisibilityModeChanged_listener_)。 示例如下。 请注意，该属性 `args.visibilityMode` 的类型为 [VisibilityMode](/javascript/api/office/office.visibilitymode)。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

函数返回另一个 *取消注册处理程序* 的函数。 下面是一个简单但不稳固的示例。

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

方法是异步的，并返回一个承诺，这意味着代码需要等待承诺的实现，然后才能调用 `onVisibilityModeChanged` **取消注册** 处理程序。

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

deregister 函数也是异步的，并返回一个承诺。 因此，如果您有在取消注册完成之前不应运行的代码，则应该等待取消注册函数返回的承诺。

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>另请参阅

- [将 Office 加载项配置为使用共享 JavaScript 运行时](configure-your-add-in-to-use-a-shared-runtime.md)
- [文档打开时在 Office 加载项中运行代码](run-code-on-document-open.md)
