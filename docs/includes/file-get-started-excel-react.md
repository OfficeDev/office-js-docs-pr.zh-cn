# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="663c1-101">使用 React 生成 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="663c1-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="663c1-102">本文将逐步介绍如何使用 React 和 Excel JavaScript API 生成 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="663c1-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="663c1-103">环境</span><span class="sxs-lookup"><span data-stu-id="663c1-103">Environment</span></span>

- <span data-ttu-id="663c1-104">**Office 桌面**：确保你安装了最新版本的 Office。</span><span class="sxs-lookup"><span data-stu-id="663c1-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="663c1-105">加载项命令需要内部版本 16.0.6769.0000 或更高版本（推荐 **16.0.6868.0000**）。</span><span class="sxs-lookup"><span data-stu-id="663c1-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="663c1-106">学习如何 [安装最新版本的 Office 应用程序](http://aka.ms/latestoffice)。</span><span class="sxs-lookup"><span data-stu-id="663c1-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="663c1-107">**Office Online**：没有额外的设置。</span><span class="sxs-lookup"><span data-stu-id="663c1-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="663c1-108">请注意，对工作/学校帐户的 Office Online 命令的支持处于预览状态。</span><span class="sxs-lookup"><span data-stu-id="663c1-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="663c1-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="663c1-109">Prerequisites</span></span>

- <span data-ttu-id="663c1-110">全局安装 [Create React App](https://github.com/facebookincubator/create-react-app)。</span><span class="sxs-lookup"><span data-stu-id="663c1-110">Install [Create React App](https://github.com/facebookincubator/create-react-app) globally.</span></span>

    ```bash
    npm install -g create-react-app
    ```

- <span data-ttu-id="663c1-111">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="663c1-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a><span data-ttu-id="663c1-112">生成新的 React 应用</span><span class="sxs-lookup"><span data-stu-id="663c1-112">Generate a new React app</span></span>

<span data-ttu-id="663c1-113">使用 Create React App 生成 React 应用。</span><span class="sxs-lookup"><span data-stu-id="663c1-113">Use Create React App to generate your React app.</span></span> <span data-ttu-id="663c1-114">在终端运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="663c1-114">From the terminal, run the following command:</span></span>

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a><span data-ttu-id="663c1-115">生成清单文件并旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="663c1-115">Generate the manifest file and sideload the add-in</span></span>

<span data-ttu-id="663c1-116">每个加载项都需要用于定义其设置和功能的清单文件。</span><span class="sxs-lookup"><span data-stu-id="663c1-116">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="663c1-117">转到应用程序文件夹。</span><span class="sxs-lookup"><span data-stu-id="663c1-117">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="663c1-118">使用 Yeoman 生成器生成加载项的清单文件。</span><span class="sxs-lookup"><span data-stu-id="663c1-118">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="663c1-119">运行下面的命令，再回答提示问题，如以下屏幕截图所示：</span><span class="sxs-lookup"><span data-stu-id="663c1-119">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="663c1-120">**选择一个项目类型：** `Office Add-in containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="663c1-120">**Choose a project type:** `Office Add-in containing the manifest only`</span></span>
    - <span data-ttu-id="663c1-121">**要将你的外接程序命名为什么?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="663c1-121">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="663c1-122">**要支持哪一个 Office 客户端应用程序?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="663c1-122">**Which Office client application would you like to support?:** `Excel`</span></span>


    <span data-ttu-id="663c1-123">完成向导后，可以使用清单文件和资源文件来构建项目。</span><span class="sxs-lookup"><span data-stu-id="663c1-123">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>
    
    ![Yeoman 生成器](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="663c1-125">如果系统提示覆盖 **package.json**，请回答“否”****（不覆盖）。</span><span class="sxs-lookup"><span data-stu-id="663c1-125">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

3. <span data-ttu-id="663c1-126">请按照运行加载项所用平台对应的说明操作，以在 Excel 中旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="663c1-126">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="663c1-127">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="663c1-127">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="663c1-128">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="663c1-128">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="663c1-129">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="663c1-129">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

## <a name="update-the-app"></a><span data-ttu-id="663c1-130">更新应用</span><span class="sxs-lookup"><span data-stu-id="663c1-130">Update the app</span></span>

1. <span data-ttu-id="663c1-131">打开“public/index.html”****，紧靠 `</head>` 标记前面添加以下 `<script>` 标记，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="663c1-131">Open **public/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. <span data-ttu-id="663c1-132">打开“src/index.js”****，将 `ReactDOM.render(<App />, document.getElementById('root'));` 替换为以下代码，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="663c1-132">Open **src/index.js**, replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following code, and save the file.</span></span> 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. <span data-ttu-id="663c1-133">打开“src/App.js”****，将文件内容替换为以下代码，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="663c1-133">Open **src/App.js**, replace file contents with the following code, and save the file.</span></span> 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. <span data-ttu-id="663c1-134">打开“src/App.css”****，将文件内容替换为以下 CSS 代码，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="663c1-134">Open **src/App.css**, replace file contents with the following CSS code, and save the file.</span></span> 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## <a name="try-it-out"></a><span data-ttu-id="663c1-135">试用</span><span class="sxs-lookup"><span data-stu-id="663c1-135">Try it out</span></span>

1. <span data-ttu-id="663c1-136">通过终端运行下面的命令，以启动开发人员服务器。</span><span class="sxs-lookup"><span data-stu-id="663c1-136">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="663c1-137">Windows：</span><span class="sxs-lookup"><span data-stu-id="663c1-137">Windows:</span></span>
    ```bash
    set HTTPS=true&&npm start
    ```

    <span data-ttu-id="663c1-138">先决条件</span><span class="sxs-lookup"><span data-stu-id="663c1-138">macOS:</span></span>
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > <span data-ttu-id="663c1-p105">此时，浏览器窗口打开，其中包含加载项。请关闭此窗口。</span><span class="sxs-lookup"><span data-stu-id="663c1-p105">A browser window will open with the add-in in it. Close this window.</span></span>

2. <span data-ttu-id="663c1-141">在 Excel 中，依次选择“主页”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="663c1-141">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="663c1-143">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="663c1-143">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="663c1-144">在任务窗格中，选择“设置颜色”**** 按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="663c1-144">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="663c1-146">后续步骤</span><span class="sxs-lookup"><span data-stu-id="663c1-146">Next steps</span></span>

<span data-ttu-id="663c1-p106">恭喜！已使用 React 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="663c1-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="663c1-149">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="663c1-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="663c1-150">另请参阅</span><span class="sxs-lookup"><span data-stu-id="663c1-150">See also</span></span>

* [<span data-ttu-id="663c1-151">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="663c1-151">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="663c1-152">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="663c1-152">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="663c1-153">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="663c1-153">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="663c1-154">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="663c1-154">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
