---
ms.date: 11/04/2019
title: 教程：在 Excel 自定义函数和任务窗格之间共享数据和事件（预览）
ms.prod: excel
description: 在 Excel 中，在自定义函数和任务窗格之间共享数据和事件
localization_priority: Priority
ms.openlocfilehash: dcd4bced7e1419a57256f4ec54e3ff72c0edf9ef
ms.sourcegitcommit: 42bcf9059327a8d71a7ab223805aea68be9ed6b5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/04/2019
ms.locfileid: "37962099"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="7674e-103">教程：在 Excel 自定义函数和任务窗格之间共享数据和事件（预览）</span><span class="sxs-lookup"><span data-stu-id="7674e-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

<span data-ttu-id="7674e-104">Excel 自定义函数和任务窗格共享全局数据，并可实现相互之间的函数调用。</span><span class="sxs-lookup"><span data-stu-id="7674e-104">Excel custom functions and the task pane share global data, and can make function calls into each other.</span></span> <span data-ttu-id="7674e-105">若要配置项目以便自定义函数可与任务窗格配合使用，请按照本文中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="7674e-105">To configure your project so that custom functions can work with the task pane, follow the instructions in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="7674e-106">本文中所述的功能目前处于预览阶段，可能会发生更改。</span><span class="sxs-lookup"><span data-stu-id="7674e-106">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="7674e-107">暂不支持在生产环境中使用。</span><span class="sxs-lookup"><span data-stu-id="7674e-107">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="7674e-108">本文中的预览功能仅适用于 Windows 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="7674e-108">The preview features in this article are only available on Excel on Windows.</span></span> <span data-ttu-id="7674e-109">若要试用预览功能，需[加入 Office 预览体验计划](https://insider.office.com/zh-CN/join)。</span><span class="sxs-lookup"><span data-stu-id="7674e-109">To try the preview features, you will need to [join Office Insider](https://insider.office.com/zh-CN/join).</span></span>  <span data-ttu-id="7674e-110">试用预览版功能的好方法是使用 Office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="7674e-110">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="7674e-111">如果还没有 Office 365 订阅，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个订阅。</span><span class="sxs-lookup"><span data-stu-id="7674e-111">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="7674e-112">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="7674e-112">Create the add-in project</span></span>

<span data-ttu-id="7674e-113">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="7674e-113">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="7674e-114">运行下面的命令，使用下面的答案回答提示问题：</span><span class="sxs-lookup"><span data-stu-id="7674e-114">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="7674e-115">选择项目类型： **Excel 自定义函数加载项项目**</span><span class="sxs-lookup"><span data-stu-id="7674e-115">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="7674e-116">选择脚本类型： **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="7674e-116">Choose a script type: </span></span>
- <span data-ttu-id="7674e-117">你想要如何命名加载项？ **我的 Office 加载项**</span><span class="sxs-lookup"><span data-stu-id="7674e-117">What do you want to name your add-in? **My Office Add-in**</span></span>

![回答 Office 中的提示问题以创建加载项项目的屏幕截图。](../images/yo-office-excel-project.png)

<span data-ttu-id="7674e-119">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="7674e-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="7674e-120">配置清单</span><span class="sxs-lookup"><span data-stu-id="7674e-120">Configure the add-in manifest</span></span>

1. <span data-ttu-id="7674e-121">启动 Visual Studio Code 并打开“**我的 Office 加载项**”项目。</span><span class="sxs-lookup"><span data-stu-id="7674e-121">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="7674e-122">打开 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="7674e-122">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="7674e-123">更改 `<Requirements>` 部分以使用 **CustomFunctionsRuntime** 版本 **1.2**，如以下代码所示。</span><span class="sxs-lookup"><span data-stu-id="7674e-123">Change the `<Requirements>` section to use **CustomFunctionsRuntime** version **1.2** as shown in the following code.</span></span>
    
    ```xml
    <Requirements> 
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. <span data-ttu-id="7674e-124">在工作簿的 `<Host>` 元素下，添加以下 `<Runtimes>` 部分。</span><span class="sxs-lookup"><span data-stu-id="7674e-124">Under the `<Host>` element for the workbook, add the following `<Runtimes>` section.</span></span> <span data-ttu-id="7674e-125">生存期需要**较长**，以便在关闭任务窗格时自定义函数仍可正常工作。</span><span class="sxs-lookup"><span data-stu-id="7674e-125">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>
    
    ```xml
    <Hosts>
    <Host xsi:type="Workbook">
    <Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
    </Runtimes>
    ```
    
5. <span data-ttu-id="7674e-126">在 `<Page>` 元素中，将源位置从 **Functions.Page.Url** 更改为 **TaskPaneAndCustomFunction.Url**。</span><span class="sxs-lookup"><span data-stu-id="7674e-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. <span data-ttu-id="7674e-127">在 `<DesktopFormFactor>` 部分中，将 **FunctionFile** 从 **Commands.Url** 更改为使用 **TaskPaneAndCustomFunction.Url**。</span><span class="sxs-lookup"><span data-stu-id="7674e-127">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. <span data-ttu-id="7674e-128">在 `<Action>` 部分中，将源位置从 **Taskpane.Url** 更改为 **TaskPaneAndCustomFunction.Url**。</span><span class="sxs-lookup"><span data-stu-id="7674e-128">In the `<Action>` section, change the source location from **Taskpane.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. <span data-ttu-id="7674e-129">为 **TaskPaneAndCustomFunction.Url** 添加新的 **Url id**，它指向 **taskpane.html**。</span><span class="sxs-lookup"><span data-stu-id="7674e-129">Add a new **Url id** for **TaskPaneAndCustomFunction.Url** that points to **taskpane.html**.</span></span>
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. <span data-ttu-id="7674e-130">保存更改并重新生成项目。</span><span class="sxs-lookup"><span data-stu-id="7674e-130">Save your changes and run the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="7674e-131">共享自定义函数和任务窗格代码之间的状态</span><span class="sxs-lookup"><span data-stu-id="7674e-131">Share state between custom function and task pane code</span></span> 

<span data-ttu-id="7674e-132">由于自定义函数在与任务窗格代码相同的上下文中运行，因此可以直接共享状态，无需使用 **Storage** 对象。</span><span class="sxs-lookup"><span data-stu-id="7674e-132">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="7674e-133">以下说明介绍了如何在自定义函数和任务窗格代码之间共享全局变量。</span><span class="sxs-lookup"><span data-stu-id="7674e-133">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="7674e-134">创建用于获取或存储共享状态的自定义函数</span><span class="sxs-lookup"><span data-stu-id="7674e-134">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="7674e-135">在 Visual Studio Code 中，打开文件 **src/functions/functions.js**。</span><span class="sxs-lookup"><span data-stu-id="7674e-135">In Visual Studio Code, open the file **config\config.json**.</span></span>
2. <span data-ttu-id="7674e-136">在第 1 行，将以下代码插入到最顶部。</span><span class="sxs-lookup"><span data-stu-id="7674e-136">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="7674e-137">这将初始化一个名为 **sharedState** 的全局变量。</span><span class="sxs-lookup"><span data-stu-id="7674e-137">This will initialize a global variable named **sharedState**.</span></span>
    
    ```js
    window.sharedState = "empty";
    ```
    
3. <span data-ttu-id="7674e-138">添加以下代码，创建将值存储到 **sharedState** 变量的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="7674e-138">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>
    
    ```js
    /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
    function storeValue(sharedValue) {
    window.sharedState = sharedValue;
    return "value stored";
    }
    ```
    
4. <span data-ttu-id="7674e-139">添加以下代码，创建获取 **sharedState** 变量的当前值的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="7674e-139">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

    ```js
    /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
    function getValue() {
    return window.sharedState;
    }
    ```
    
5. <span data-ttu-id="7674e-140">保存文件。</span><span class="sxs-lookup"><span data-stu-id="7674e-140">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="7674e-141">创建任务窗格控件以处理全局数据</span><span class="sxs-lookup"><span data-stu-id="7674e-141">Create task pane controls to work with global data</span></span> 

1. <span data-ttu-id="7674e-142">打开 file**src/taskpane/taskpane.html**。</span><span class="sxs-lookup"><span data-stu-id="7674e-142">Open the file**src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="7674e-143">关闭 `</main>` 元素后，添加以下 HTML。</span><span class="sxs-lookup"><span data-stu-id="7674e-143">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="7674e-144">该 HTML 创建两个用于获取或存储全局数据的文本框和按钮。</span><span class="sxs-lookup"><span data-stu-id="7674e-144">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
    <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>
    <p>Store new value to shared state</p>
    <div>
    <input type="text" id="storeBox" />
    <button onclick="storeSharedValue()">Store</button>
    </div>
     
    <p>Get shared state value</p>
    <div>
    <input type="text" id="getBox" />
    <button onclick="getSharedValue()">Get</button>
    </div>
    ```
    
3. <span data-ttu-id="7674e-145">在 `<body>` 元素前，添加以下脚本。</span><span class="sxs-lookup"><span data-stu-id="7674e-145">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="7674e-146">当用户想存储或获取全局数据时，此代码将处理按钮单击事件。</span><span class="sxs-lookup"><span data-stu-id="7674e-146">This code will handle the button click events when the user wants to store or get global data.</span></span>
    
    ```js
    <script>
    function storeSharedValue() {
    let sharedValue = document.getElementById('storeBox').value;
    window.sharedState = sharedValue;
    }
    
    function getSharedValue() {
    document.getElementById('getBox').value = window.sharedState;
    }</script>
    ```
    
4. <span data-ttu-id="7674e-147">保存文件。</span><span class="sxs-lookup"><span data-stu-id="7674e-147">Save the file.</span></span>
5. <span data-ttu-id="7674e-148">生成项目</span><span class="sxs-lookup"><span data-stu-id="7674e-148">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="7674e-149">在自定义函数和任务窗格之间尝试共享数据</span><span class="sxs-lookup"><span data-stu-id="7674e-149">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="7674e-150">使用以下命令启动项目。</span><span class="sxs-lookup"><span data-stu-id="7674e-150">Start the migration by using the following command.</span></span>

    ```command&nbsp;line
    npm run start
    ```

<span data-ttu-id="7674e-151">Excel 启动后，可使用“任务窗格”按钮来存储或获取共享数据。</span><span class="sxs-lookup"><span data-stu-id="7674e-151">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="7674e-152">在自定义函数的单元格中输入 `=CONTOSO.GETVALUE()`，以检索相同的共享数据。</span><span class="sxs-lookup"><span data-stu-id="7674e-152">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="7674e-153">或使用 `=CONTOSO.STOREVALUE(“new value”)` 将共享数据更改为新值。</span><span class="sxs-lookup"><span data-stu-id="7674e-153">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="7674e-154">如本文所示配置项目，可在自定义函数和任务窗格之间共享上下文。</span><span class="sxs-lookup"><span data-stu-id="7674e-154">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="7674e-155">不支持从自定义函数调用 Office API。</span><span class="sxs-lookup"><span data-stu-id="7674e-155">Calling Office APIs from custom functions is not supported.</span></span> <span data-ttu-id="7674e-156">如果需要与文档交互，在 [onCalculated 事件](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details)中实现对 Office API 的调用。</span><span class="sxs-lookup"><span data-stu-id="7674e-156">If you need to interact with the document, implement calls to the Office APIs in the [onCalculated event](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details).</span></span>

