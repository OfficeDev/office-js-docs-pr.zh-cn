---
ms.date: 06/14/2021
title: 将 Office 加载项配置为使用共享 JavaScript 运行时
ms.prod: non-product-specific
description: 将 Office 加载项配置为使用共享 JavaScript 运行时，以支持其他功能区、任务窗格和自定义函数功能。
localization_priority: Priority
ms.openlocfilehash: 9874d0fef2dc4966f106d1d88e4e897469300c0b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076215"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="f714c-103">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="f714c-103">Configure your Office Add-in to use a shared JavaScript runtime</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="f714c-104">你可以将 Office 加载项配置为在单个共享 JavaScript 运行时（也称为共享运行时）中运行它的所有代码。</span><span class="sxs-lookup"><span data-stu-id="f714c-104">You can configure your Office Add-in to run all of its code in a single shared JavaScript runtime (also known as a shared runtime).</span></span> <span data-ttu-id="f714c-105">这可在加载项中实现更好的协调，并且可从加载项的所有部分访问 DOM 和 CORS。</span><span class="sxs-lookup"><span data-stu-id="f714c-105">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="f714c-106">它还能启用其他功能，例如文档打开时运行代码，或者启用或禁用功能区按钮。</span><span class="sxs-lookup"><span data-stu-id="f714c-106">It also enables additional features such as running code when the document opens, or enabling or disabling ribbon buttons.</span></span> <span data-ttu-id="f714c-107">若要将加载项配置为使用共享 JavaScript 运行时，请按照本文中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="f714c-107">To configure your add-in to use a shared JavaScript runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="f714c-108">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="f714c-108">Create the add-in project</span></span>

<span data-ttu-id="f714c-109">如果要启动新项目，请按照以下步骤使用[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 或 PowerPoint 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="f714c-109">If you are starting a new project, follow these steps to use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create an Excel or PowerPoint add-in project.</span></span>

<span data-ttu-id="f714c-110">执行下列操作之一：</span><span class="sxs-lookup"><span data-stu-id="f714c-110">Do one of the following:</span></span>

- <span data-ttu-id="f714c-111">要生成带自定义函数的 Excel 加载项，请运行命令 `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`。</span><span class="sxs-lookup"><span data-stu-id="f714c-111">To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.</span></span>

    <span data-ttu-id="f714c-112">或者</span><span class="sxs-lookup"><span data-stu-id="f714c-112">or</span></span>

- <span data-ttu-id="f714c-113">要生成 PowerPoint 加载项，请运行命令 `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`。</span><span class="sxs-lookup"><span data-stu-id="f714c-113">To generate a PowerPoint add-in, run the command `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.</span></span>

<span data-ttu-id="f714c-114">生成器将创建项目并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="f714c-114">The generator will create the project and install supporting Node components.</span></span>

> [!NOTE]
> <span data-ttu-id="f714c-115">还可以使用本文中的步骤更新现有Visual Studio项目以使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="f714c-115">You can also use the steps in this article to update an existing Visual Studio project to use the shared runtime.</span></span> <span data-ttu-id="f714c-116">但是，可能需要更新清单的 XML 架构。</span><span class="sxs-lookup"><span data-stu-id="f714c-116">However, you may need to update the XML schemas for the manifest.</span></span> <span data-ttu-id="f714c-117">有关详细信息，请参阅 [排除 Office 加载项开发错误故障](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。</span><span class="sxs-lookup"><span data-stu-id="f714c-117">For more information, see [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f714c-118">配置清单</span><span class="sxs-lookup"><span data-stu-id="f714c-118">Configure the manifest</span></span>

<span data-ttu-id="f714c-119">对于新项目或现有项目，请按照以下步骤将其配置为使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="f714c-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span> <span data-ttu-id="f714c-120">以下步骤能确保你使用[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)生成你的项目。</span><span class="sxs-lookup"><span data-stu-id="f714c-120">These steps assume you have generated your project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="f714c-121">启动 Visual Studio Code 并打开你生成的 Excel 或 PowerPoint 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="f714c-121">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
1. <span data-ttu-id="f714c-122">打开 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="f714c-122">Open the **manifest.xml** file.</span></span>
1. <span data-ttu-id="f714c-123">如果生成 Excel 加载项，请更新“要求”部分，以使用[共享运行时](../reference/requirement-sets/shared-runtime-requirement-sets.md)，而不是自定义函数运行时。</span><span class="sxs-lookup"><span data-stu-id="f714c-123">If you generated an Excel add-in, update the requirements section to use the [shared runtime](../reference/requirement-sets/shared-runtime-requirement-sets.md) instead of the custom function runtime.</span></span> <span data-ttu-id="f714c-124">XML 应该如下所示。</span><span class="sxs-lookup"><span data-stu-id="f714c-124">The XML should appear as follows.</span></span>

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

1. <span data-ttu-id="f714c-125">找到 `<VersionOverrides>` 部分并添加 `<Host ...>` 标记内的以下 `<Runtimes>` 部分。</span><span class="sxs-lookup"><span data-stu-id="f714c-125">Find the `<VersionOverrides>` section and add the following `<Runtimes>` section just inside the `<Host ...>` tag.</span></span> <span data-ttu-id="f714c-126">生存期需要 **较长**，以便在关闭任务窗格时加载项代码仍可运行。</span><span class="sxs-lookup"><span data-stu-id="f714c-126">The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed.</span></span> <span data-ttu-id="f714c-127">`resid` 值是 **Taskpane.Url**，它引用 **manifest.xml** 文件底部附近的 ` <bt:Urls>` 部分中指定的 **taskpane.html** 文件位置。</span><span class="sxs-lookup"><span data-stu-id="f714c-127">The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.</span></span>

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

1. <span data-ttu-id="f714c-128">如果你生成带自定义函数的 Excel 加载项，请查找 `<Page>` 元素。</span><span class="sxs-lookup"><span data-stu-id="f714c-128">If you generated an Excel add-in with custom functions, find the `<Page>` element.</span></span> <span data-ttu-id="f714c-129">然后将源位置从 **Functions.Page.Url** 更改为 **Taskpane.Url**。</span><span class="sxs-lookup"><span data-stu-id="f714c-129">Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. <span data-ttu-id="f714c-130">查找 `<FunctionFile ...>` 标记并将 `resid` 从 **Commands.Url** 更改为  **Taskpane.Url**。</span><span class="sxs-lookup"><span data-stu-id="f714c-130">Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**.</span></span> <span data-ttu-id="f714c-131">请注意，如果你没有操作命令，则不会有 **FunctionFile** 条目，可跳过此步骤。</span><span class="sxs-lookup"><span data-stu-id="f714c-131">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. <span data-ttu-id="f714c-132">保存 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="f714c-132">Save the **manifest.xml** file.</span></span>

## <a name="configure-the-webpackconfigjs-file"></a><span data-ttu-id="f714c-133">配置 webpack.config.js 文件</span><span class="sxs-lookup"><span data-stu-id="f714c-133">Configure the webpack.config.js file</span></span>

<span data-ttu-id="f714c-134">**webpack.config.js** 将生成多个运行时加载程序。</span><span class="sxs-lookup"><span data-stu-id="f714c-134">The **webpack.config.js** will build multiple runtime loaders.</span></span> <span data-ttu-id="f714c-135">你需要对其进行修改，以通过 **taskpane.html** 文件仅加载共享 JavaScript 运行时。 </span><span class="sxs-lookup"><span data-stu-id="f714c-135">You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.</span></span>

1. <span data-ttu-id="f714c-136">启动 Visual Studio Code 并打开你生成的 Excel 或 PowerPoint 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="f714c-136">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
1. <span data-ttu-id="f714c-137">打开 **webpack.config.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="f714c-137">Open the **webpack.config.js** file.</span></span>
1. <span data-ttu-id="f714c-138">如果你的 **webpack.config.js** 文件有以下 **functions.html** 插件代码，请将其删除。</span><span class="sxs-lookup"><span data-stu-id="f714c-138">If your **webpack.config.js** file has the following **functions.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. <span data-ttu-id="f714c-139">如果你的 **webpack.config.js** 文件有以下 **commands.html** 插件代码，请将其删除。</span><span class="sxs-lookup"><span data-stu-id="f714c-139">If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. <span data-ttu-id="f714c-140">如果你的项目使用 **functions** 或 **commands** 区块，请将其添加到如下所示的区块列表中（以下代码适用于你的项目使用上述两种区块时）。</span><span class="sxs-lookup"><span data-stu-id="f714c-140">If your project used either the **functions** or **commands** chunks, add them to the chunks list as shown next (the following code is for if your project used both chunks).</span></span>

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. <span data-ttu-id="f714c-141">保存更改并重新生成项目。</span><span class="sxs-lookup"><span data-stu-id="f714c-141">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

> [!NOTE]
> <span data-ttu-id="f714c-142">如果你的项目有 **functions.html** 文件或 **commands.html** 文件，可将其删除。</span><span class="sxs-lookup"><span data-stu-id="f714c-142">If your project has a **functions.html** file or **commands.html** file, they can be removed.</span></span> <span data-ttu-id="f714c-143">**taskpane.html** 将通过你刚才进行的 webpack 更新将 **functions.js** 和 **commands.js** 代码加载到 共享 JavaScript 运行时中。</span><span class="sxs-lookup"><span data-stu-id="f714c-143">The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.</span></span>

## <a name="test-your-office-add-in-changes"></a><span data-ttu-id="f714c-144">测试 Office 加载项更改</span><span class="sxs-lookup"><span data-stu-id="f714c-144">Test your Office Add-in changes</span></span>

<span data-ttu-id="f714c-145">你可以通过使用以下指令，确认你正在正确使用共享 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="f714c-145">You can confirm that you are using the shared JavaScript runtime correctly by using the following instructions.</span></span>

1. <span data-ttu-id="f714c-146">打开 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="f714c-146">Open the **manifest.xml** file.</span></span>
1. <span data-ttu-id="f714c-147">找到 `<Control xsi:type="Button" id="TaskpaneButton">` 部分并更改以下 `<Action ...>` XML。</span><span class="sxs-lookup"><span data-stu-id="f714c-147">Find the `<Control xsi:type="Button" id="TaskpaneButton">` section and change the following `<Action ...>` XML.</span></span>

    <span data-ttu-id="f714c-148">来自：</span><span class="sxs-lookup"><span data-stu-id="f714c-148">from:</span></span>

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    <span data-ttu-id="f714c-149">更改为：</span><span class="sxs-lookup"><span data-stu-id="f714c-149">to:</span></span>

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. <span data-ttu-id="f714c-150">打开 **./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="f714c-150">Open the **./src/commands/commands.js** file.</span></span>
1. <span data-ttu-id="f714c-151">将 **操作** 函数替换成以下代码。</span><span class="sxs-lookup"><span data-stu-id="f714c-151">Replace the **action** function with the code below.</span></span> <span data-ttu-id="f714c-152">这将更新函数，以打开并修改任务窗格按钮，从而增加一个计数器。</span><span class="sxs-lookup"><span data-stu-id="f714c-152">This will update the function to open and modify the task pane button to increment a counter.</span></span> <span data-ttu-id="f714c-153">使用一个命令打开并访问任务窗格 DOM 仅适用于共享 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="f714c-153">Opening and accessing the task pane DOM from a command only works with the shared JavaScript runtime.</span></span>

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. <span data-ttu-id="f714c-154">保存更改并运行项目。</span><span class="sxs-lookup"><span data-stu-id="f714c-154">Save your changes and run the project.</span></span>

   ```command line
   npm start
   ```

<span data-ttu-id="f714c-155">每次选择加载项按钮，它都会将 **运行** 按钮文本更改为 **转到** ，并在其后增加一个计数器。</span><span class="sxs-lookup"><span data-stu-id="f714c-155">Each time you select the add-ins button, it will change the **run** button text to **go** and increment a counter after it.</span></span>

## <a name="runtime-lifetime"></a><span data-ttu-id="f714c-156">运行时生存期</span><span class="sxs-lookup"><span data-stu-id="f714c-156">Runtime lifetime</span></span>

<span data-ttu-id="f714c-157">添加 `Runtime` 元素时，还需要指定值为 `long` 或 `short` 的生存期。</span><span class="sxs-lookup"><span data-stu-id="f714c-157">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="f714c-158">将此值设置为 `long` 以利用相关功能，例如在文档打开时启动加载项，在关闭任务窗格后继续运行代码，或从自定义函数中使用 CORS 和 DOM。</span><span class="sxs-lookup"><span data-stu-id="f714c-158">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

> [!NOTE]
> <span data-ttu-id="f714c-159">默认生存期值为`short`，但我们建议在 Excel 加载项中使用`long`。如果在此例中将运行时设置为`short`，则当按下某个功能区按钮时，Excel 加载项将启动，但在功能区处理程序运行完毕后，它可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="f714c-159">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="f714c-160">同样，打开任务窗格时，加载项将启动，但关闭任务窗格时，加载项可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="f714c-160">Similarly, your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> <span data-ttu-id="f714c-161">如果加载启动项包括清单中的 `Runtimes` 元素（共享运行时所需），它将使用 Internet Explorer 11，而不考虑 Windows 或 Microsoft 365 版本。</span><span class="sxs-lookup"><span data-stu-id="f714c-161">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="f714c-162">有关详细信息，请参阅[运行时](../reference/manifest/runtimes.md)。</span><span class="sxs-lookup"><span data-stu-id="f714c-162">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="about-the-shared-javascript-runtime"></a><span data-ttu-id="f714c-163">关于共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="f714c-163">About the shared JavaScript runtime</span></span>

<span data-ttu-id="f714c-164">在 Windows 或 Mac 上，加载项将在单独的 JavaScript 运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。</span><span class="sxs-lookup"><span data-stu-id="f714c-164">On Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="f714c-165">这会产生一些局限性，例如无法轻松共享全局数据，也不能通过自定义函数访问所有 CORS 功能。</span><span class="sxs-lookup"><span data-stu-id="f714c-165">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="f714c-166">但是，你可以将 Office 加载项配置为在同一 JavaScript 运行时（也称为共享运行时）中共享代码。</span><span class="sxs-lookup"><span data-stu-id="f714c-166">However, you can configure your Office Add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="f714c-167">这可在加载项中实现更好的协调，并且可从加载项的所有部分访问任务窗格 DOM 和 CORS。</span><span class="sxs-lookup"><span data-stu-id="f714c-167">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="f714c-168">配置共享运行时可实现以下方案。</span><span class="sxs-lookup"><span data-stu-id="f714c-168">Configuring a shared runtime enables the following scenarios.</span></span>

- <span data-ttu-id="f714c-169">Office 加载项可使用其他 UI 功能：</span><span class="sxs-lookup"><span data-stu-id="f714c-169">Your Office Add-in can use additional UI features:</span></span>
  - [<span data-ttu-id="f714c-170">将自定义键盘快捷方式添加到 Office 加载项（预览）</span><span class="sxs-lookup"><span data-stu-id="f714c-170">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
  - [<span data-ttu-id="f714c-171">在 Office 加载项中创建自定义上下文选项卡（预览）</span><span class="sxs-lookup"><span data-stu-id="f714c-171">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
  - [<span data-ttu-id="f714c-172">启用和禁用加载项命令</span><span class="sxs-lookup"><span data-stu-id="f714c-172">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
  - [<span data-ttu-id="f714c-173">文档打开时在 Office 加载项中运行代码</span><span class="sxs-lookup"><span data-stu-id="f714c-173">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
  - [<span data-ttu-id="f714c-174">显示或隐藏 Office 加载项的任务窗格</span><span class="sxs-lookup"><span data-stu-id="f714c-174">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- <span data-ttu-id="f714c-175">对于 Excel 加载项：</span><span class="sxs-lookup"><span data-stu-id="f714c-175">For Excel add-ins:</span></span>
  - <span data-ttu-id="f714c-176">自定义函数将具有完整的 CORS 支持。</span><span class="sxs-lookup"><span data-stu-id="f714c-176">Custom functions will have full CORS support.</span></span>
  - <span data-ttu-id="f714c-177">自定义函数可调用 Office.js API 以读取电子表格文档数据。</span><span class="sxs-lookup"><span data-stu-id="f714c-177">Custom functions can call Office.js APIs to read spreadsheet document data.</span></span>

<span data-ttu-id="f714c-178">对于 Windows 版 Office，共享运行时需要 Microsoft Internet Explorer 11 浏览器实例，如 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。此外，加载项在功能区上显示的任何按钮都将在同一共享运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="f714c-178">For Office on Windows, the shared runtime requires a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="f714c-179">下图显示了自定义函数、功能区 UI 和任务窗格代码如何在同一 JavaScript 运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="f714c-179">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Excel 中自定义函数、任务窗格和功能区按钮均在共享 IE 浏览器运行时中运行的图表。](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a><span data-ttu-id="f714c-181">调试</span><span class="sxs-lookup"><span data-stu-id="f714c-181">Debugging</span></span>

<span data-ttu-id="f714c-182">使用共享运行时时，目前不能使用 Visual Studio Code 在 Windows 版 Excel 中调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="f714c-182">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="f714c-183">你需要改为使用开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="f714c-183">You'll need to use developer tools instead.</span></span> <span data-ttu-id="f714c-184">有关详细信息，请参阅[使用 Windows 10 上的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)。</span><span class="sxs-lookup"><span data-stu-id="f714c-184">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

### <a name="multiple-task-panes"></a><span data-ttu-id="f714c-185">多个任务窗格</span><span class="sxs-lookup"><span data-stu-id="f714c-185">Multiple task panes</span></span>

<span data-ttu-id="f714c-186">如果计划使用共享运行时，请勿将你的加载项设计为使用多个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f714c-186">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="f714c-187">共享运行时仅支持使用一个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f714c-187">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="f714c-188">请注意，不含 `<TaskpaneID>` 的任何任务窗格都被视为不同的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f714c-188">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="f714c-189">向我们提供反馈</span><span class="sxs-lookup"><span data-stu-id="f714c-189">Give us feedback</span></span>

<span data-ttu-id="f714c-190">我们非常乐意听取有关此功能的反馈。</span><span class="sxs-lookup"><span data-stu-id="f714c-190">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="f714c-191">如果你发现此功能存在任何 bug、问题或具有相关请求，请通过在 [office-js repo](https://github.com/OfficeDev/office-js) 中创建 GitHub 问题来告诉我们。</span><span class="sxs-lookup"><span data-stu-id="f714c-191">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="f714c-192">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f714c-192">See also</span></span>

- [<span data-ttu-id="f714c-193">从自定义函数中调用 Excel API</span><span class="sxs-lookup"><span data-stu-id="f714c-193">Call Excel APIs from a custom function</span></span>](../excel/call-excel-apis-from-custom-function.md)
- [<span data-ttu-id="f714c-194">将自定义键盘快捷方式添加到 Office 加载项（预览）</span><span class="sxs-lookup"><span data-stu-id="f714c-194">Add custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
- [<span data-ttu-id="f714c-195">在 Office 加载项中创建自定义上下文选项卡（预览）</span><span class="sxs-lookup"><span data-stu-id="f714c-195">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
- [<span data-ttu-id="f714c-196">启用和禁用加载项命令</span><span class="sxs-lookup"><span data-stu-id="f714c-196">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
- [<span data-ttu-id="f714c-197">文档打开时在 Office 加载项中运行代码</span><span class="sxs-lookup"><span data-stu-id="f714c-197">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
- [<span data-ttu-id="f714c-198">显示或隐藏 Office 加载项的任务窗格</span><span class="sxs-lookup"><span data-stu-id="f714c-198">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- [<span data-ttu-id="f714c-199">教程：在 Excel 自定义函数和任务窗格之间共享数据和事件</span><span class="sxs-lookup"><span data-stu-id="f714c-199">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
