---
title: 教程：在 Excel 自定义函数和任务窗格之间共享数据和事件
description: 在 Excel 中，在自定义函数和任务窗格之间共享数据和事件
ms.date: 05/17/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6d96b10d6cd6e9bb7909b9d6d64b9a65fcac5b3a
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275599"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a><span data-ttu-id="6a91f-103">教程：在 Excel 自定义函数和任务窗格之间共享数据和事件</span><span class="sxs-lookup"><span data-stu-id="6a91f-103">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>

<span data-ttu-id="6a91f-104">你可配置 Excel 加载项以使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="6a91f-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="6a91f-105">这样，就可以共享全局数据，或在任务窗格和自定义函数之间发送事件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-105">This makes it possible to shared global data, or send events between the task pane and custom functions.</span></span>

<span data-ttu-id="6a91f-106">对于大多数自定义函数方案，建议使用共享运行时，除非有特定的原因需要使用非任务窗格（无 UI）自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6a91f-106">For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function.</span></span>

<span data-ttu-id="6a91f-107">本教程假定您熟悉使用 Yo Office 生成器创建外接项目。</span><span class="sxs-lookup"><span data-stu-id="6a91f-107">This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects.</span></span> <span data-ttu-id="6a91f-108">如果还没有，请考虑填写[Excel 自定义函数教程](./excel-tutorial-create-custom-functions.md)。</span><span class="sxs-lookup"><span data-stu-id="6a91f-108">Consider completing the [Excel custom functions tutorial](./excel-tutorial-create-custom-functions.md), if you haven't already.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="6a91f-109">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="6a91f-109">Create the add-in project</span></span>

<span data-ttu-id="6a91f-110">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="6a91f-110">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="6a91f-111">运行下面的命令，使用下面的答案回答提示问题：</span><span class="sxs-lookup"><span data-stu-id="6a91f-111">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="6a91f-112">选择项目类型： **Excel 自定义函数加载项项目**</span><span class="sxs-lookup"><span data-stu-id="6a91f-112">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="6a91f-113">选择脚本类型： **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="6a91f-113">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="6a91f-114">你想要如何命名加载项？ **我的 Office 加载项**</span><span class="sxs-lookup"><span data-stu-id="6a91f-114">What do you want to name your add-in? **My Office Add-in**</span></span>

![回答 Office 中的提示问题以创建加载项项目的屏幕截图。](../images/yo-office-excel-project.png)

<span data-ttu-id="6a91f-116">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-116">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="6a91f-117">配置清单</span><span class="sxs-lookup"><span data-stu-id="6a91f-117">Configure the manifest</span></span>

1. <span data-ttu-id="6a91f-118">启动 Visual Studio Code 并打开“**我的 Office 加载项**”项目。</span><span class="sxs-lookup"><span data-stu-id="6a91f-118">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="6a91f-119">打开 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="6a91f-120">找到 `<VersionOverrides>` 部分并添加以下 `<Runtimes>` 部分。</span><span class="sxs-lookup"><span data-stu-id="6a91f-120">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="6a91f-121">生存期需要**较长**，以便在关闭任务窗格时自定义函数仍可正常工作。</span><span class="sxs-lookup"><span data-stu-id="6a91f-121">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="6a91f-122">在 `<Page>` 元素中，将源位置从 **Functions.Page.Url** 更改为 **ContosoAddin.Url**。</span><span class="sxs-lookup"><span data-stu-id="6a91f-122">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="6a91f-123">在 `<DesktopFormFactor>` 部分中，将 **FunctionFile** 从 **Commands.Url** 更改为使用 **ContosoAddin.Url**。</span><span class="sxs-lookup"><span data-stu-id="6a91f-123">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="6a91f-124">在 `<Action>` 部分中，将源位置从 **Taskpane.Url** 更改为 **ContosoAddin.Url**。</span><span class="sxs-lookup"><span data-stu-id="6a91f-124">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="6a91f-125">为 **ContosoAddin.Url** 添加新的 **Url id**，它指向 **taskpane.html**。</span><span class="sxs-lookup"><span data-stu-id="6a91f-125">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="6a91f-126">保存更改并重新生成项目。</span><span class="sxs-lookup"><span data-stu-id="6a91f-126">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="6a91f-127">共享自定义函数和任务窗格代码之间的状态</span><span class="sxs-lookup"><span data-stu-id="6a91f-127">Share state between custom function and task pane code</span></span>

<span data-ttu-id="6a91f-128">由于自定义函数在与任务窗格代码相同的上下文中运行，因此可以直接共享状态，无需使用 **Storage** 对象。</span><span class="sxs-lookup"><span data-stu-id="6a91f-128">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="6a91f-129">以下说明介绍了如何在自定义函数和任务窗格代码之间共享全局变量。</span><span class="sxs-lookup"><span data-stu-id="6a91f-129">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="6a91f-130">创建用于获取或存储共享状态的自定义函数</span><span class="sxs-lookup"><span data-stu-id="6a91f-130">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="6a91f-131">在 Visual Studio Code 中，打开文件 **src/functions/functions.js**。</span><span class="sxs-lookup"><span data-stu-id="6a91f-131">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="6a91f-132">在第 1 行，将以下代码插入到最顶部。</span><span class="sxs-lookup"><span data-stu-id="6a91f-132">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="6a91f-133">这将初始化一个名为 **sharedState** 的全局变量。</span><span class="sxs-lookup"><span data-stu-id="6a91f-133">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="6a91f-134">添加以下代码，创建将值存储到 **sharedState** 变量的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6a91f-134">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="6a91f-135">添加以下代码，创建获取 **sharedState** 变量的当前值的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6a91f-135">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="6a91f-136">保存此文件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-136">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="6a91f-137">创建任务窗格控件以处理全局数据</span><span class="sxs-lookup"><span data-stu-id="6a91f-137">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="6a91f-138">打开 **src/taskpane/taskpane.html** 文件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-138">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="6a91f-139">在 `</head>` 元素前，添加以下脚本元素。</span><span class="sxs-lookup"><span data-stu-id="6a91f-139">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="6a91f-140">关闭 `</main>` 元素后，添加以下 HTML。</span><span class="sxs-lookup"><span data-stu-id="6a91f-140">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="6a91f-141">该 HTML 创建两个用于获取或存储全局数据的文本框和按钮。</span><span class="sxs-lookup"><span data-stu-id="6a91f-141">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
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

4. <span data-ttu-id="6a91f-142">在 `<body>` 元素前，添加以下脚本。</span><span class="sxs-lookup"><span data-stu-id="6a91f-142">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="6a91f-143">当用户想存储或获取全局数据时，此代码将处理按钮单击事件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-143">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="6a91f-144">保存文件。</span><span class="sxs-lookup"><span data-stu-id="6a91f-144">Save the file.</span></span>
6. <span data-ttu-id="6a91f-145">生成项目</span><span class="sxs-lookup"><span data-stu-id="6a91f-145">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="6a91f-146">在自定义函数和任务窗格之间尝试共享数据</span><span class="sxs-lookup"><span data-stu-id="6a91f-146">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="6a91f-147">使用以下命令启动项目。</span><span class="sxs-lookup"><span data-stu-id="6a91f-147">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="6a91f-148">Excel 启动后，可使用“任务窗格”按钮来存储或获取共享数据。</span><span class="sxs-lookup"><span data-stu-id="6a91f-148">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="6a91f-149">在自定义函数的单元格中输入 `=CONTOSO.GETVALUE()`，以检索相同的共享数据。</span><span class="sxs-lookup"><span data-stu-id="6a91f-149">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="6a91f-150">或使用 `=CONTOSO.STOREVALUE("new value")` 将共享数据更改为新值。</span><span class="sxs-lookup"><span data-stu-id="6a91f-150">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="6a91f-151">如本文所示配置项目，可在自定义函数和任务窗格之间共享上下文。</span><span class="sxs-lookup"><span data-stu-id="6a91f-151">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="6a91f-152">可以从自定义函数中调用一些 Office Api。</span><span class="sxs-lookup"><span data-stu-id="6a91f-152">Calling some Office APIs from custom functions is possible.</span></span> <span data-ttu-id="6a91f-153">有关更多详细信息，[请参阅通过自定义函数调用 Microsoft Excel api](../excel/call-excel-apis-from-custom-function.md) 。</span><span class="sxs-lookup"><span data-stu-id="6a91f-153">[See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.</span></span>
