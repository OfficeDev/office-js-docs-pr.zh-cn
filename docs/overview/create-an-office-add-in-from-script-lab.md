---
title: 从脚本实验室代码创建独立的 Office 外接程序
description: 了解如何将代码片段从脚本实验室移动到 Yo Office 项目
ms.topic: how-to
ms.date: 04/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 725ce9b44c55b46e6d0ab0c085973947fcf88201
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810146"
---
# <a name="create-a-standalone-office-add-in-from-your-script-lab-code"></a>从脚本实验室代码创建独立的 Office 外接程序

如果在脚本实验室中创建了代码片段，则可能需要将其转换为独立加载项。 可以将代码从脚本实验室复制到由 Office 加载项的 [Yeoman 生成器生成的项目](../develop/yeoman-generator-overview.md)（也称为“Yo Office”）。 然后，可以继续将代码作为外接程序进行开发，最终可以将其部署到其他人。

本文中的步骤指的是 [Visual Studio Code](https://code.visualstudio.com/)，但你可以使用任何你喜欢的代码编辑器。

## <a name="create-a-new-yo-office-project"></a>创建新的 Yo Office 项目

需要创建独立的外接程序项目，该项目将成为代码片段代码的新开发位置。

运行命令 `yo office --projectType taskpane --ts true --host <host> --name "basic-sample"`，其中 `<host>` 是以下值之一。

- Excel
- outlook
- Powerpoint
- Word

> [!IMPORTANT]
> `--name` 参数值必须采用双引号，即使没有空格也是如此。

上一个命令创建名为 “**basic-sample**” 的新项目文件夹。 它配置为在指定的主机中运行，并使用 TypeScript。 默认情况下，脚本实验室使用 TypeScript，但大多数代码片段都是 JavaScript。 如果愿意，可以生成 Yo Office JavaScript 项目，但只需确保复制的任何代码都是 JavaScript。

## <a name="open-the-snippet-in-script-lab"></a>在脚本实验室中打开代码片段

使用脚本实验室中的现有代码片段了解如何将代码片段复制到 Yo Office 生成的项目。

1. 打开 Office（Word、Excel、PowerPoint 或 Outlook），然后打开脚本实验室。
1. 选择 **脚本实验室** > **代码**。 如果在 Outlook 中工作，请打开电子邮件以查看功能区上的脚本实验室。
1. 在“脚本实验室”任务窗格中，选择 **示例**。 然后，根据你正在使用的 Office 主机选择一个基本示例。
    - 对于 Excel 或 Word，请选择 **基本 API 调用 （TypeScript）** 示例。
    - 对于 Outlook，请选择 **使用外接程序设置** 示例。
    - 对于 PowerPoint，请选择 **基本 API 调用 （Ofice 2013）** 示例。

## <a name="copy-snippet-code-to-visual-studio-code"></a>将代码片段代码复制到Visual Studio代码

现在，可以在 VS Code 中将代码片段中的代码复制到 Yo Office 项目。

- 在VS Code中，打开 **基本示例** 项目。

在后续步骤中，你将从脚本实验室中的多个选项卡复制代码。

:::image type="content" source="../images/script-lab-script-tabs.png" alt-text="脚本实验室中选项卡的屏幕截图。":::

### <a name="copy-task-pane-code"></a>复制任务窗格代码

1. 在VS Code中，打开 **/src/taskpane/taskpane.ts** 文件。 如果使用的是 JavaScript 项目，则文件名 **taskpane.js**。
1. 在“脚本实验室”中，选择 **脚本** 选项卡。
1. 将 **脚本** 选项卡中的所有代码复制到剪贴板。 将 **JavaScript) 的 taskpane.ts** (或 **taskpane.js** 的全部内容替换为复制的代码。

### <a name="copy-task-pane-html"></a>复制任务窗格 HTML

1. 在VS Code中，打开 **/src/taskpane/taskpane.html** 文件。
1. 在“脚本实验室”中，选择“ **HTML** ”选项卡。
1. 将 **HTML** 选项卡中的所有 HTML 复制到剪贴板。 将 `<body>` 标记中的所有 HTML 替换为复制的 HTML。

### <a name="copy-task-pane-css"></a>复制任务窗格 CSS

1. 在VS Code中，打开 **/src/taskpane/taskpane.css** 文件。
1. 在“脚本实验室”中，选择“ **CSS** ”选项卡。
1. 将 **CSS** 选项卡中的所有 CSS 复制到剪贴板。 将 **taskpane.css** 的全部内容替换为复制的 CSS。
1. 保存对前面步骤中更新的文件所做的所有更改。

## <a name="add-jquery-support"></a>添加 jQuery 支持

脚本实验室在代码片段中使用 jQuery。 需要将此依赖项添加到 Yo Office 项目才能成功运行代码。

1. 打开 **taskpane.html** 文件，并将以下脚本标记添加到 `<head>` 部分。

    ```html
     <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-3.3.1.js"></script>
    ```

    > [!NOTE]
    > jQuery 的特定版本可能有所不同。 可以通过选择“ **库”** 选项卡来确定脚本实验室使用的版本。

1. 在 VS Code 中打开终端并输入以下命令。

    ```command&nbsp;line
    npm install --save-dev jquery@3.1.1
    npm install --save-dev @types/jquery@3.3.1
    ```

如果创建了具有其他库依赖项的代码片段，请务必将其添加到 Yo Office 项目。 在脚本实验室的“ **库** ”选项卡上查找所有库依赖项的列表。

## <a name="handle-initialization"></a>处理初始化

脚本实验室自动处理`Office.onReady`初始化。 需要修改代码，以提供自己的`Office.onReady`处理程序。

1. 打开 **taskpane.ts** （或适用于 JavaScript 的 **taskpane.js** ）文件。
1. 对于 Excel 或 Word，请替换：

    ```typescript
    $("#run").click(() => tryCatch(run));
    ```

    与:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(() => tryCatch(run));
      });
    });
    ```

1. 对于 Outlook，请替换：

    ```typescript
    $("#get").click(get);
    $("#set").click(set);
    $("#save").click(save);
    ```

    与:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#get").click(get);
        $("#set").click(set);
        $("#save").click(save);
      });
    });
    ```

1. 对于 PowerPoint，请替换：

    ```typescript
    $("#run").click(run);
    ```

    与:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(run);
      });
    });
    ```

1. 保存文件。

## <a name="custom-functions"></a>自定义函数

如果代码片段使用自定义函数，则需要使用 Yo Office 自定义函数模板。 若要将自定义函数转换为独立加载项，请执行以下步骤。

1. 运行命令 `yo office --projectType excel-functions --ts true --name "functions-sample"`。

    > [!IMPORTANT]
    > `--name` 参数值必须采用双引号，即使没有空格也是如此。

1. 打开 Excel，然后打开“脚本实验室”。
1. 选择 **脚本实验室** > **代码**。
1. 在“脚本实验室”任务窗格中，选择 **示例**，然后选择 **基本自定义函数** 示例。
1. 打开 **/src/functions/functions.ts** 文件。 如果使用的是 JavaScript 项目，则文件名 **functions.js**。
1. 在“脚本实验室”中，选择 **脚本** 选项卡。
1. 将 **脚本** 选项卡中的所有代码复制到剪贴板。 将代码粘贴到 **functions.ts** (或 **javaScript)** functions.js的顶部，以及复制的代码。
1. 保存文件。

## <a name="test-the-standalone-add-in"></a>测试独立加载项

完成所有步骤后，运行并测试独立加载项。 运行以下命令以开始使用。

```command&nbsp;line
npm start
```

Office 将启动，你可以从功能区打开加载项的任务窗格。 恭喜！ 现在，你可以继续将外接程序构建为独立项目。

## <a name="console-logging"></a>控制台日志记录

脚本实验室中的许多代码片段将输出写入任务窗格底部的控制台部分。 Yo Office 项目没有控制台部分。 所有 `console.log*` 语句都将写入默认调试控制台（如浏览器开发人员工具）。 如果希望输出转到任务窗格，则需要更新代码。
