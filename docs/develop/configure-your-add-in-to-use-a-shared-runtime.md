---
title: 将 Office 外接程序配置为使用共享运行时
description: 将 Office 外接程序配置为使用共享运行时来支持其他功能区、任务窗格和自定义函数功能。
ms.date: 07/18/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: e6b10cc2d342d95a8542146ecbd95d750322421f
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422934"
---
# <a name="configure-your-office-add-in-to-use-a-shared-runtime"></a>将 Office 外接程序配置为使用共享运行时

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

可以将 Office 加载项配置为在单个 [共享运行时](../testing/runtimes.md#shared-runtime)中运行其所有代码。 这可在加载项中实现更好的协调，并且可从加载项的所有部分访问 DOM 和 CORS。 它还能启用其他功能，例如文档打开时运行代码，或者启用或禁用功能区按钮。 若要将加载项配置为使用共享运行时，请按照本文中的说明进行操作。

## <a name="create-the-add-in-project"></a>创建加载项项目

如果要启动新项目，请使用 [适用于 Office 加载项的Yeoman 生成器](yeoman-generator-overview.md)创建 Excel、PowerPoint 或 Word 加载项项目。

运行命令 `yo office --projectType taskpane --name "my office add in" --host <host> --js true`，其中 `<host>` 是以下值之一。

- Excel
- Powerpoint
- Word

> [!IMPORTANT]
> `--name` 参数值必须采用双引号，即使没有空格也是如此。

对于 **--projecttype**、**--name**、**--js** 命令行选项，你可以使用不同的选项。 有关选项的完整列表，请参阅 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

生成器将创建项目并安装支持的 Node 组件。 还可以使用本文中的步骤更新 Visual Studio 项目以使用共享运行时。 但是，可能需要更新清单的 XML 架构。 有关详细信息，请参阅 [排除 Office 加载项开发错误故障](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

## <a name="configure-the-manifest"></a>配置清单

对于新项目或现有项目，请按照以下步骤将其配置为使用共享运行时。 以下步骤能确保你使用[适用于 Office 加载项的 Yeoman 生成器](yeoman-generator-overview.md)生成你的项目。

1. 启动 Visual Studio Code 并打开加载项项目。
1. 打开 **manifest.xml** 文件。
1. 对于 Excel 或 PowerPoint 外接程序，请更新要求部分，以包括[共享运行时](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)。 请务必删除 `CustomFunctionsRuntime` 要求（如果存在）。 XML 应该如下所示。

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

    > [!NOTE]
    > 不要将 `SharedRuntime` 要求集添加到 Word 加载项的清单。 加载加载项时会导致错误，这是一个目前已知的问题。

1. 查找 **\<VersionOverrides\>** 部分并添加以下 **\<Runtimes\>** 部分。 生存期需要 **较长**，以便在关闭任务窗格时加载项代码仍可运行。 `resid` 值是 **Taskpane.Url**，它引用 **manifest.xml** 文件底部附近的 `<bt:Urls>` 部分中指定的 **taskpane.html** 文件位置。

    > [!IMPORTANT]
    > 必须按照以下 XML 中显示的确切顺序在 **\<Host\>** 元素之后输入 **\<Runtimes\>** 部分。

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
         <Runtimes>
           <Runtime resid="Taskpane.Url" lifetime="long" />
         </Runtimes>
       ...
       </Host>
   ```

1. 如果已生成带自定义函数的 Excel 加载项，请查找 **\<Page\>** 元素。 然后将源位置从 **Functions.Page.Url** 更改为 **Taskpane.Url**。

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. 查找 **\<FunctionFile\>** 标记并将 `resid` 从 **Commands.Url** 更改为 **Taskpane.Url**。 请注意，如果你没有操作命令，则将不会具有 **\<FunctionFile\>** 条目，并且可跳过此步骤。

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. 保存 **manifest.xml** 文件。

## <a name="configure-the-webpackconfigjs-file"></a>配置 webpack.config.js 文件

**webpack.config.js** 将生成多个运行时加载程序。 需要对其进行修改，以便仅通过 **taskpane.html** 文件加载共享运行时。

1. 启动 Visual Studio Code 并打开生成的加载项项目。
1. 打开 **webpack.config.js** 文件。
1. 如果你的 **webpack.config.js** 文件有以下 **functions.html** 插件代码，请将其删除。

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. 如果你的 **webpack.config.js** 文件有以下 **commands.html** 插件代码，请将其删除。

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. 如果你的项目使用 **functions** 或 **commands** 区块，请将其添加到如下所示的区块列表中（以下代码适用于你的项目使用上述两种区块时）。

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. 保存更改并重新生成项目。

   ```command line
   npm run build
   ```

> [!NOTE]
> 如果你的项目有 **functions.html** 文件或 **commands.html** 文件，可将其删除。 **taskpane.html** 将通过刚才进行的 Webpack 更新将 **functions.js** 和 **commands.js** 代码加载到共享运行时。

## <a name="test-your-office-add-in-changes"></a>测试 Office 加载项更改

可以使用以下说明确认是否正确使用共享运行时。

1. 打开 **taskpane.js** 文件。
1. 使用以下代码替换文件的全部内容。 这将显示任务窗格已被打开次数的计数。 仅在共享运行时支持添加 onVisibilityModeChanged 事件。

    ```javascript
    /*global document, Office*/

    let _count = 0;

    Office.onReady(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      updateCount(); // Update count on first open.
      Office.addin.onVisibilityModeChanged(function (args) {
        if (args.visibilityMode === "Taskpane") {
          updateCount(); // Update count on subsequent opens.
        }
      });
    });

    function updateCount() {
      _count++;
      document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
    }
    ```

1. 保存更改并运行项目。

   ```command line
   npm start
   ```

每次打开任务窗格时，其打开次数的计数都将递增。 **_count** 的值不会丢失，因为即使任务窗格关闭，共享运行时也会使代码保持运行状态。

## <a name="runtime-lifetime"></a>运行时生存期

添加元素 **\<Runtime\>** 时，还指定值为或`short`值的`long`生存期。 将此值设置为 `long` 以利用相关功能，例如在文档打开时启动加载项，在关闭任务窗格后继续运行代码，或从自定义函数中使用 CORS 和 DOM。

> [!NOTE]
> 默认生存期值为 `short`，但我们建议在 Excel、PowerPoint、Word 加载项中使用 `long`。如果在此例中将运行时设置为 `short`，则当按下某个功能区按钮时，加载项将启动，但在功能区处理程序运行完毕后，它可能会关闭。 同样，打开任务窗格时，加载项将启动，但关闭任务窗格时，加载项可能会关闭。

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> 如果外接程序包含 **\<Runtimes\>** 共享运行时) 所需的清单 (元素，并且满足将 Microsoft Edge 与基于 WebView2 (Chromium的) 配合使用的条件，则它使用该 WebView2 控件。 如果不满足条件，则使用 Internet Explorer 11，而不考虑 Windows 或 Microsoft 365 版本。 有关详细信息，请参阅 [运行时](/javascript/api/manifest/runtimes) 和 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

## <a name="about-the-shared-runtime"></a>关于共享运行时

在 Windows 或 Mac 上，外接程序将在单独的运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。 这会产生一些局限性，例如无法轻松共享全局数据，也不能通过自定义函数访问所有 CORS 功能。

但是，可以将 Office 外接程序配置为在同一运行时中共享代码 (也称为共享运行时) 。 这可在加载项中实现更好的协调，并且可从加载项的所有部分访问任务窗格 DOM 和 CORS。

配置共享运行时可实现以下方案。

- Office 加载项可使用其他 UI 功能。
  - [启用和禁用加载项命令](../design/disable-add-in-commands.md)
  - [文档打开时在 Office 加载项中运行代码](run-code-on-document-open.md)
  - [显示或隐藏 Office 加载项的任务窗格](show-hide-add-in.md)
- 以下内容仅适用于 Excel 加载项。
  - [将自定义键盘快捷方式添加到 Office 加载项（预览）](../design/keyboard-shortcuts.md)
  - [在 Office 加载项中创建自定义上下文选项卡（预览）](../design/contextual-tabs.md)
  - 自定义函数将具有完整的 CORS 支持。
  - 自定义函数可调用 Office.js API 以读取电子表格文档数据。

对于 Windows 上的 Office，如果满足 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述，共享运行时使用 WebView2（基于 Chromium）的Microsoft Edge。否则，它使用 Internet Explorer 11。 此外，外接程序在功能区上显示的任何按钮都将在同一共享运行时中运行。 下图显示了自定义函数、功能区 UI 和任务窗格代码如何在同一运行时中运行。

![自定义函数、任务窗格和功能区按钮的示意图，这些按钮都在 Excel 的共享浏览器运行时中运行。](../images/custom-functions-in-browser-runtime.png)

### <a name="debug"></a>调试

使用共享运行时时，目前不能使用 Visual Studio Code 在 Windows 版 Excel 中调试自定义函数。 你需要改为使用开发人员工具。 有关详细信息，请参阅使用适用于 Internet Explorer 的开发人员工具[调试外接程序，](../testing/debug-add-ins-using-f12-tools-ie.md)或[使用 Microsoft Edge (基于 Chromium ) 中的开发人员工具调试外接程序](../testing/debug-add-ins-using-devtools-edge-chromium.md)。

### <a name="multiple-task-panes"></a>多个任务窗格

如果计划使用共享运行时，请勿将你的加载项设计为使用多个任务窗格。 共享运行时仅支持使用一个任务窗格。 请注意，不含 `<TaskpaneID>` 的任何任务窗格都被视为不同的任务窗格。

## <a name="see-also"></a>另请参阅

- [从自定义函数中调用 Excel API](../excel/call-excel-apis-from-custom-function.md)
- [将自定义键盘快捷方式添加到 Office 加载项（预览）](../design/keyboard-shortcuts.md)
- [在 Office 加载项中创建自定义上下文选项卡（预览）](../design/contextual-tabs.md)
- [启用和禁用加载项命令](../design/disable-add-in-commands.md)
- [文档打开时在 Office 加载项中运行代码](run-code-on-document-open.md)
- [显示或隐藏 Office 加载项的任务窗格](show-hide-add-in.md)
- [教程：在 Excel 自定义函数和任务窗格之间共享数据和事件](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Office 加载项中的运行时](../testing/runtimes.md)
