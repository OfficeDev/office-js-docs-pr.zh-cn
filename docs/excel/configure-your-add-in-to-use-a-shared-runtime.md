---
ms.date: 08/25/2020
title: 将 Excel 加载项配置为共享浏览器运行时
ms.prod: excel
description: 将 Excel 加载项配置为共享浏览器运行时并在同一运行时中运行功能区、任务窗格和自定义函数代码。
localization_priority: Priority
ms.openlocfilehash: be4e79ae54376a9574ffb0669681c2fba7cd158c
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996275"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="f6595-103">将 Excel 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="f6595-103">Configure your Excel add-in to use a shared JavaScript runtime</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f6595-104">运行 Windows 版 Excel 或 Mac 版 Excel 时，加载项将在单独的 JavaScript 运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。</span><span class="sxs-lookup"><span data-stu-id="f6595-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="f6595-105">这会产生一些局限性，例如无法轻松共享全局数据，也不能从自定义函数访问所有 CORS 功能。</span><span class="sxs-lookup"><span data-stu-id="f6595-105">This creates limitations such as not being able to easily share global data, and not having access to all CORS functionality from a custom function.</span></span>

<span data-ttu-id="f6595-106">但是，你可以将 Excel 加载项配置为在共享 JavaScript 运行时中共享代码。</span><span class="sxs-lookup"><span data-stu-id="f6595-106">However, you can configure your Excel add-in to share code in a shared JavaScript runtime.</span></span> <span data-ttu-id="f6595-107">这可在加载项中实现更好的协调，并且可从加载项的所有部分访问 DOM 和 CORS。</span><span class="sxs-lookup"><span data-stu-id="f6595-107">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="f6595-108">它还允许在文档打开时运行代码，或在关闭任务窗格后继续运行代码。</span><span class="sxs-lookup"><span data-stu-id="f6595-108">It also enables you to run code when the document opens, or to run code while the task pane is closed.</span></span> <span data-ttu-id="f6595-109">若要将加载项配置为使用共享运行时，请按照本文中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="f6595-109">To configure your add-in to use a shared runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="f6595-110">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="f6595-110">Create the add-in project</span></span>

<span data-ttu-id="f6595-111">如果要启动新项目，请按照以下步骤使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="f6595-111">If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="f6595-112">运行下面的命令，使用下面的答案回答提示问题：</span><span class="sxs-lookup"><span data-stu-id="f6595-112">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="f6595-113">选择项目类型： **Excel 自定义函数加载项项目**</span><span class="sxs-lookup"><span data-stu-id="f6595-113">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="f6595-114">选择脚本类型： **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="f6595-114">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="f6595-115">你想要如何命名加载项？ **我的 Office 加载项**</span><span class="sxs-lookup"><span data-stu-id="f6595-115">What do you want to name your add-in? **My Office Add-in**</span></span>

![回答 Office 中的提示问题以创建加载项项目的屏幕截图。](../images/yo-office-excel-project.png)

<span data-ttu-id="f6595-117">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="f6595-117">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f6595-118">配置清单</span><span class="sxs-lookup"><span data-stu-id="f6595-118">Configure the manifest</span></span>

<span data-ttu-id="f6595-119">对于新项目或现有项目，请按照以下步骤将其配置为使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="f6595-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span>

1. <span data-ttu-id="f6595-120">启动 Visual Studio Code 并打开“ **我的 Office 加载项** ”项目。</span><span class="sxs-lookup"><span data-stu-id="f6595-120">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="f6595-121">打开 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="f6595-121">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="f6595-122">找到 `<VersionOverrides>` 部分并添加以下 `<Runtimes>` 部分。</span><span class="sxs-lookup"><span data-stu-id="f6595-122">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="f6595-123">生存期需要 **较长** ，以便在关闭任务窗格时自定义函数仍可正常工作。</span><span class="sxs-lookup"><span data-stu-id="f6595-123">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span> <span data-ttu-id="f6595-124">resid 是 `ContosoAddin.Url`，它在后面的资源部分中引用字符串。</span><span class="sxs-lookup"><span data-stu-id="f6595-124">The resid is `ContosoAddin.Url` which references a string in the resources section later.</span></span> <span data-ttu-id="f6595-125">可使用所需的任何 resid 值，但它应匹配加载项元素中其他元素的 resid。</span><span class="sxs-lookup"><span data-stu-id="f6595-125">You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="f6595-126">在 `<Page>` 元素中，将源位置从 **Functions.Page.Url** 更改为 **ContosoAddin.Url** 。</span><span class="sxs-lookup"><span data-stu-id="f6595-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="f6595-127">此 resid 匹配 `<Runtime>` resid 元素。</span><span class="sxs-lookup"><span data-stu-id="f6595-127">This resid matches the `<Runtime>` resid element.</span></span> <span data-ttu-id="f6595-128">请注意，如果你没有自定义函数，则不会有 **页面** 条目，可跳过此步骤。</span><span class="sxs-lookup"><span data-stu-id="f6595-128">Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="f6595-129">在 `<DesktopFormFactor>` 部分中，将 **FunctionFile** 从 **Commands.Url** 更改为使用 **ContosoAddin.Url** 。</span><span class="sxs-lookup"><span data-stu-id="f6595-129">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span> <span data-ttu-id="f6595-130">请注意，如果你没有操作命令，则不会有 **FunctionFile** 条目，可跳过此步骤。</span><span class="sxs-lookup"><span data-stu-id="f6595-130">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="f6595-131">在 `<Action>` 部分中，将源位置从 **Taskpane.Url** 更改为 **ContosoAddin.Url** 。</span><span class="sxs-lookup"><span data-stu-id="f6595-131">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="f6595-132">请注意，如果你没有任务窗格，则不会有 **ShowTaskpane** 操作，可跳过此步骤。</span><span class="sxs-lookup"><span data-stu-id="f6595-132">Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="f6595-133">为 **ContosoAddin.Url** 添加新的 **Url id** ，它指向 **taskpane.html** 。</span><span class="sxs-lookup"><span data-stu-id="f6595-133">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/dist/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="f6595-134">确保 taskpane.html 有一个参考 dist/functions.js file 文件的 `<script>` 标记。</span><span class="sxs-lookup"><span data-stu-id="f6595-134">Make sure the taskpane.html has a `<script>` tag that references the dist/functions.js file.</span></span> <span data-ttu-id="f6595-135">示例如下。</span><span class="sxs-lookup"><span data-stu-id="f6595-135">The following is an example.</span></span>

   ```html
   <script type="text/javascript" src="/dist/functions.js" ></script>
   ```

   > [!NOTE]
   > <span data-ttu-id="f6595-136">如果加载项使用 Webpack 和 HtmlWebpackPlugin 插入脚本标记，与 Yeoman 生成器创建的加载项一样（请参阅上面的[创建加载项项目](#create-the-add-in-project)），则必须确保 functions.js 模块包含在 `chunks` 数组中，如下例所示。</span><span class="sxs-lookup"><span data-stu-id="f6595-136">If the add-in uses Webpack and the HtmlWebpackPlugin to insert script tags, as add-ins created by the Yeoman generator do (see [Create the add-in project](#create-the-add-in-project) above), then you must ensure that the functions.js module is included in the `chunks` array as in the following example.</span></span>
   >
   > ```javascript
   > new HtmlWebpackPlugin({
   >     filename: "taskpane.html",
   >     template: "./src/taskpane/taskpane.html",
   >     chunks: ["polyfill", "taskpane", "functions"]
   > }),
   >```

9. <span data-ttu-id="f6595-137">保存更改并重新生成项目。</span><span class="sxs-lookup"><span data-stu-id="f6595-137">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a><span data-ttu-id="f6595-138">运行时生存期</span><span class="sxs-lookup"><span data-stu-id="f6595-138">Runtime lifetime</span></span>

<span data-ttu-id="f6595-139">添加 `Runtime` 元素时，还需要指定值为 `long` 或 `short` 的生存期。</span><span class="sxs-lookup"><span data-stu-id="f6595-139">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="f6595-140">将此值设置为 `long` 以利用相关功能，例如在文档打开时启动加载项，在关闭任务窗格后继续运行代码，或从自定义函数中使用 CORS 和 DOM。</span><span class="sxs-lookup"><span data-stu-id="f6595-140">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

>[!NOTE]
> <span data-ttu-id="f6595-141">默认生存期值为`short`，但我们建议在 Excel 加载项中使用`long`。如果在此例中将运行时设置为`short`，则当按下某个功能区按钮时，Excel 加载项将启动，但在功能区处理程序运行完毕后，它可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="f6595-141">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="f6595-142">同样，打开任务窗格时，加载项将启动，但在任务窗格关闭时可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="f6595-142">Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

>[!NOTE]
> <span data-ttu-id="f6595-143">如果加载启动项包括清单中的 `Runtimes` 元素（共享运行时所需），它将使用 Internet Explorer 11，而不考虑 Windows 或 Microsoft 365 版本。</span><span class="sxs-lookup"><span data-stu-id="f6595-143">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="f6595-144">有关详细信息，请参阅[运行时](../reference/manifest/runtimes.md)。</span><span class="sxs-lookup"><span data-stu-id="f6595-144">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="multiple-task-panes"></a><span data-ttu-id="f6595-145">多个任务窗格</span><span class="sxs-lookup"><span data-stu-id="f6595-145">Multiple task panes</span></span>

<span data-ttu-id="f6595-146">如果计划使用共享运行时，请勿将你的加载项设计为使用多个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f6595-146">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="f6595-147">共享运行时仅支持使用一个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f6595-147">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="f6595-148">请注意，不含 `<TaskpaneID>` 的任何任务窗格都被视为不同的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f6595-148">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f6595-149">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f6595-149">Next steps</span></span>

- <span data-ttu-id="f6595-150">有关在共享运行时中使用 Excel JavaScript API 和自定义 Excel 函数的详细信息，请参阅文章[从自定义函数中调用 Excel API](call-excel-apis-from-custom-function.md)。</span><span class="sxs-lookup"><span data-stu-id="f6595-150">Read the [Call Excel APIs from a custom function](call-excel-apis-from-custom-function.md) article for details on using the Excel JavaScript APIs and custom Excel functions in a shared runtime.</span></span>
- <span data-ttu-id="f6595-151">探索模式和实践示例[管理功能区和任务窗格 UI，并在文档打开时运行代码](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario)，以查看活动中的共享 JavaScript 运行时的更大示例。</span><span class="sxs-lookup"><span data-stu-id="f6595-151">Explore the patterns-and-practices sample [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) to see a larger example of the shared JavaScript runtime in action.</span></span>
- <span data-ttu-id="f6595-152">有关向项目添加自定义键盘快捷方式的信息，请阅读 [Office 加载项中的自定义键盘快捷方式](../design/keyboard-shortcuts.md)。</span><span class="sxs-lookup"><span data-stu-id="f6595-152">Read the [Custom keyboard shortcuts in Office Add-ins](../design/keyboard-shortcuts.md) for information about adding custom keyboard shortcuts to your project.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6595-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f6595-153">See also</span></span>

- [<span data-ttu-id="f6595-154">概述：在共享 JavaScript 运行时中运行加载项代码</span><span class="sxs-lookup"><span data-stu-id="f6595-154">Overview: Run your add-in code in a shared JavaScript runtime</span></span>](custom-functions-shared-overview.md)
