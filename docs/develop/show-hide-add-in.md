---
title: 显示或隐藏 Office 加载项的任务窗格
description: 了解如何在加载项持续运行时以编程方式隐藏或显示其用户界面。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 95f8c716bf1a0331fe47bc74e5aad49c17b65437
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660128"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>显示或隐藏 Office 加载项的任务窗格

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

可以通过调用 `Office.addin.showAsTaskpane()` 该函数来显示 Office 加载项的任务窗格。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

前面的代码假定存在名为 **CurrentQuarterSales** 的 Excel 工作表的场景。 每当激活此工作表时，加载项将使任务窗格可见。 该方法 `onCurrentQuarter` 是已为工作表注册的 [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) 事件的处理程序。

也可以通过调用 `Office.addin.hide()` 函数来隐藏任务窗格。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

前面的代码是为 [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) 事件注册的处理程序。

## <a name="additional-details-on-showing-the-task-pane"></a>有关显示任务窗格的其他详细信息

调用 `Office.addin.showAsTaskpane()`时，Office 会在任务窗格中显示分配为资源 ID 的文件 (`resid` 任务窗格的) 值。 可以通过打开 **manifest.xml** 文件并在元素中定位来 **\<SourceLocation\>** 分配或更改此`resid``<Action xsi:type="ShowTaskpane">`值。
 (请参阅 [配置 Office 加载项以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md) 获取其他详细信息。) 

由于 `Office.addin.showAsTaskpane()` 是异步方法，因此代码将继续运行，直到函数完成。 请使用关键字或`then()`方法等待此完成`await`，具体取决于使用的 JavaScript 语法。

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>将外接程序配置为使用共享运行时

若要使用这些 `showAsTaskpane()` 方法和 `hide()` 方法，外接程序必须使用共享运行时。 有关详细信息，请参阅 [配置 Office 加载项以使用共享运行时](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="preservation-of-state-and-event-listeners"></a>保留状态和事件侦听器

`showAsTaskpane()`和`hide()`方法仅更改任务窗格的 *可见性*。 他们不卸载或重新加载它 (或重新初始化其状态) 。

请考虑以下方案：任务窗格是使用选项卡设计的。 首次启动加载项时，“ **开始** ”选项卡处于打开状态。 假设用户打开 **“设置”** 选项卡，然后在任务窗格中调用 `hide()` 代码以响应某些事件。 稍后还会调用 `showAsTaskpane()` 代码以响应另一个事件。 任务窗格将重新出现，并且“ **设置”** 选项卡仍处于选中状态。

![任务窗格的屏幕截图，其中包含四个标签为“开始”、“设置”、“收藏夹”和“帐户”的选项卡。](../images/TaskpaneWithTabs.png)

此外，即使任务窗格处于隐藏状态，在任务窗格中注册的任何事件侦听器仍会继续运行。

请考虑以下方案：任务窗格具有 Excel `Worksheet.onActivated` 的已注册处理程序和 `Worksheet.onDeactivated` 名为 **Sheet1 的工作表** 的事件。 激活的处理程序会导致任务窗格中显示一个绿点。 停用的处理程序将点红色 (，这是其默认状态) 。 假设当 **Sheet1** 未激活且点为红色时，代码会调用`hide()`。 隐藏任务窗格时，将激活 **Sheet1** 。 稍后的代码调用 `showAsTaskpane()` 以响应某些事件。 当任务窗格打开时，该点为绿色，因为即使任务窗格已隐藏，事件侦听器和处理程序也已运行。

## <a name="handle-the-visibility-changed-event"></a>处理可见性更改事件

当代码更改任务窗格 `showAsTaskpane()` 的可见性时 `hide()`，Office 会触发 `VisibilityModeChanged` 该事件。 处理此事件可能很有用。 例如，假设任务窗格显示工作簿中所有工作表的列表。 如果在隐藏任务窗格时添加了新的工作表，则使任务窗格可见本身不会将新工作表名称添加到列表中。 但是，代码可以响应`VisibilityModeChanged`事件以重新加载 [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) 集合中所有工作表的 [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) 属性，如下面的示例代码所示。

若要为事件注册处理程序，请不要像在大多数 Office JavaScript 上下文中那样使用“添加处理程序”方法。 相反，有一个特殊函数可将处理程序传递给： [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1))。 示例如下。 请注意，该 `args.visibilityMode` 属性为 [VisibilityMode](/javascript/api/office/office.visibilitymode) 类型。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

该函数返回另一个 *取消注册* 处理程序的函数。 下面是一个简单但不可靠的示例。

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

该 `onVisibilityModeChanged` 方法是异步的，并返回一个承诺，这意味着代码需要等待承诺的履行，然后才能调用 **取消注册** 处理程序。

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

取消注册函数也是异步的，并返回一个承诺。 因此，如果代码在取消注册完成之前不应运行，则应等待取消注册函数返回的承诺。

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>另请参阅

- [将 Office 加载项配置为使用共享 JavaScript 运行时](configure-your-add-in-to-use-a-shared-runtime.md)
- [文档打开时在 Office 加载项中运行代码](run-code-on-document-open.md)
