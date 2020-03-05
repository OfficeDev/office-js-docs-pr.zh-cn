---
title: 在共享运行时中显示或隐藏 Office 外接程序
description: 了解如何在连续运行时以编程方式隐藏或显示外接程序的 UI
ms.date: 03/02/2020
localization_priority: Normal
ms.openlocfilehash: c028823be165723cad3c0b314b53fe7e618188b2
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/04/2020
ms.locfileid: "42413790"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime-preview"></a>在共享运行时中显示或隐藏 Office 外接程序（预览）

Office 外接程序可以包含以下任何部分：

- 任务窗格
- 不带 UI 的函数文件
- Excel 自定义函数

默认情况下，每个部件都在自己的独立 JavaScript 运行时中运行，其中包含其自己的全局对象和全局变量。 

具有两个或更多个部件的外接程序可以共享一个通用的 JavaScript 运行时。 此共享运行时功能启用在外接程序运行时隐藏和重新打开任务窗格的新预览 Api。

> [!INCLUDE [Information about using preview APIs](../includes/excel-shared-runtime-preview-note.md)]

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a>将外接程序配置为使用共享运行时

若要将外接程序配置为使用共享运行时，请参阅[configure The Office 外接程序以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="show-and-hide-the-task-pane"></a>显示和隐藏任务窗格

新的 Api 位于`Office.addin`属性中。 若要显示任务窗格，您的代码`Office.addin.showAsTaskpane()`将调用。 Office 将在任务窗格中显示分配给任务窗格的资源 ID （`resid`）的页面。 这是`resid`分配给清单`<SourceLocation>` `<Action xsi:type="ShowTaskpane">`中的的的。 （请参阅[配置 Office 外接程序以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。）

这是一种异步方法，因此，如果后续代码在完成之前不应运行，则代码应等待它。 使用`await`关键字或`then()`方法等待这一完成，具体取决于您使用的 JavaScript 语法。 以下示例假定有一个名为**CurrentQuarterSales**的 Excel 工作表。 每当激活此工作表时，加载项都应显示任务窗格。 该方法`onCurrentQuarter`是已为工作表注册的[onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated)事件的处理程序。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

若要隐藏任务窗格，您的代码`Office.addin.hide()`将调用。 下面的示例是一个为[onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated)事件注册的处理程序。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a>保留状态和事件侦听器

`hide()`和`showAsTaskpane()`方法仅更改任务窗格的*可见性*。 它们不会卸载或重新加载它（或重新初始化其状态）。

请考虑以下方案：任务窗格是用选项卡设计的。 首次启动加载项时，"**主页**" 选项卡处于打开状态。 假设用户打开 "**设置**" 选项卡，随后，任务窗格中的代码将`hide()`调用以响应某个事件。 仍在以后的`showAsTaskpane()`代码调用，以响应另一个事件。 任务窗格将重新显示，并且 "**设置**" 选项卡仍处于选中状态。

![任务窗格的屏幕截图，其中有四个标签为 "主页"、"设置"、"收藏夹" 和 "帐户"。](../images/TaskpaneWithTabs.png)

此外，即使任务窗格处于隐藏状态，在任务窗格中注册的任何事件侦听器也将继续运行。

请考虑以下方案：任务窗格有一个 Excel `Worksheet.onActivated`和`Worksheet.onDeactivated`一个名为**Sheet1**的工作表的事件的已注册处理程序。 激活的处理程序导致在任务窗格中显示一个绿色点。 已停用的处理程序会将点变为红色（这是其默认状态）。 假设该代码在`hide()` **Sheet1**未激活且点为红色时调用。 在任务窗格处于隐藏状态时， **Sheet1**处于激活状态。 后续代码调用`showAsTaskpane()`以响应某个事件。 任务窗格打开时，点为绿色，因为即使任务窗格被隐藏，也会运行事件侦听器和处理程序。

### <a name="handle-visibility-changed-event"></a>处理可见性更改事件

当您的代码通过`showAsTaskpane()` or `hide()`更改任务窗格的可见性时，Office 将`VisibilityModeChanged`触发该事件。 处理此事件可能很有用。 例如，假设任务窗格显示工作簿中所有工作表的列表。 如果在任务窗格处于隐藏状态时添加了一个新的工作表，使任务窗格可见，则它本身不会将新的工作表名称添加到列表中。 但您的代码可以响应`VisibilityModeChanged`事件以重新加载工作簿中所有工作表的[Worksheet.name](/javascript/api/excel/excel.worksheet#name)属性[。工作表](/javascript/api/excel/excel.workbook#worksheets)集合，如下面的示例代码所示。

若要注册事件的处理程序，请不要像在大多数 Office JavaScript 上下文中那样使用 "添加处理程序" 方法。 相反，有一个特殊的函数，您可以将其传递给处理程序： [onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)。 示例如下。 请注意， `args.visibilityMode`属性的类型为[VisibilityMode](/javascript/api/office/office.visibilitymode)。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

函数返回*deregisters*处理程序的另一个函数。 下面是一个简单但不可靠的示例：

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

方法是异步的，这意味着，如果代码`onVisibilityModeChanged`调用返回的取消*注册*处理程序，则应确保`onVisibilityModeChanged`在调用取消注册处理程序之前已完成。 `onVisibilityModeChanged` 执行此操作的一种方法是在`await`方法调用中使用关键字，如下面的示例所示。

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

如果您只想使用 ES2015 JavaScript，则代码可以使用`then`方法等待，直到返回的承诺对象已解决，并将返回的函数分配给全局变量，如以下示例中所示。

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

取消注册的功能本身是异步的。 因此，如果您有不应在注销完成之后运行的代码，则必须使用`await`关键字或`then`方法（如以下示例中所示）来等待取消注册功能。

取消注册处理程序：

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
