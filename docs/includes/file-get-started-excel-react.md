# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="b54a0-101">使用 React 生成 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="b54a0-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="b54a0-102">在本文中，您将了解使用 React 和 Excel JavaScript API 生成 Excel 加载项的过程。</span><span class="sxs-lookup"><span data-stu-id="b54a0-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="b54a0-103">环境</span><span class="sxs-lookup"><span data-stu-id="b54a0-103">Environment</span></span>

- <span data-ttu-id="b54a0-104">**Office 桌面**：确保你安装了最新版本的 Office。</span><span class="sxs-lookup"><span data-stu-id="b54a0-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="b54a0-105">加载项命令需要内部版本 16.0.6769.0000 或更高版本（推荐 **16.0.6868.0000**）。</span><span class="sxs-lookup"><span data-stu-id="b54a0-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="b54a0-106">学习如何 [安装最新版本的 Office 应用程序](http://aka.ms/latestoffice)。</span><span class="sxs-lookup"><span data-stu-id="b54a0-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="b54a0-107">**Office Online**：没有额外的设置。</span><span class="sxs-lookup"><span data-stu-id="b54a0-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="b54a0-108">请注意，对工作/学校帐户的 Office Online 命令的支持处于预览状态。</span><span class="sxs-lookup"><span data-stu-id="b54a0-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b54a0-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="b54a0-109">Prerequisites</span></span>

- [<span data-ttu-id="b54a0-110">Node.js</span><span class="sxs-lookup"><span data-stu-id="b54a0-110">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="b54a0-111">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="b54a0-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="b54a0-112">创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="b54a0-112">Create the web app</span></span>

1. <span data-ttu-id="b54a0-113">在本地驱动器上创建一个文件夹，并命名为“my-addin”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="b54a0-113">Create a folder on your local drive and name it **my-addin**.</span></span> <span data-ttu-id="b54a0-114">将在其中创建应用程序文件。</span><span class="sxs-lookup"><span data-stu-id="b54a0-114">This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="b54a0-115">转到应用程序文件夹。</span><span class="sxs-lookup"><span data-stu-id="b54a0-115">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="b54a0-116">使用 Yeoman 生成器生成加载项清单文件。</span><span class="sxs-lookup"><span data-stu-id="b54a0-116">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="b54a0-117">运行下面的命令，再回答提示问题，如以下屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="b54a0-117">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="b54a0-118">**选择一个项目类型：** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="b54a0-118">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="b54a0-119">**要将你的外接程序命名为什么?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="b54a0-119">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="b54a0-120">**要支持哪一个 Office 客户端应用?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="b54a0-120">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman 生成器](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="b54a0-122">完成向导后，生成器将创建项目并安装 Node 支持组件。</span><span class="sxs-lookup"><span data-stu-id="b54a0-122">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

4.  <span data-ttu-id="b54a0-123">打开 **src/components/App.tsx**，搜索注释“更新填充颜色”，然后将填充颜色从“黄色”更改为“蓝色”，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="b54a0-123">Open **src/components/App.tsx**, search for the comment "Update the fill color," then change the fill color from 'yellow' to 'blue', and save the file.</span></span> 

    ```js
    range.format.fill.color = 'blue'

    ```

5. <span data-ttu-id="b54a0-124">在\*\* src / components / App.tsx \*\*中的`render` 函数的`return` 块中，将 `<Herolist>` 更新到下面的代码中，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="b54a0-124">In the `return` block of the `render` function within **src/components/App.tsx**, update the `<Herolist>` to the code below, and save the file.</span></span> 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. <span data-ttu-id="b54a0-125">按照[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中的步骤操作，信任开发计算机操作系统的证书。</span><span class="sxs-lookup"><span data-stu-id="b54a0-125">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

7. <span data-ttu-id="b54a0-126">旁加载加载项以便在 Excel 中将显示。</span><span class="sxs-lookup"><span data-stu-id="b54a0-126">Sideload your add-in so it will appear in Excel.</span></span> <span data-ttu-id="b54a0-127">在终端中，运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="b54a0-127">In the terminal run the following command:</span></span> 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a><span data-ttu-id="b54a0-128">试用</span><span class="sxs-lookup"><span data-stu-id="b54a0-128">Try it out</span></span>

1. <span data-ttu-id="b54a0-129">通过终端运行下面的命令，以启动开发人员服务器。</span><span class="sxs-lookup"><span data-stu-id="b54a0-129">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="b54a0-130">Windows：</span><span class="sxs-lookup"><span data-stu-id="b54a0-130">Windows:</span></span>
    ```bash
    npm start
    ```

2. <span data-ttu-id="b54a0-131">在 Excel 中，依次选择“主页”\*\*\*\* 选项卡和功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="b54a0-131">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="b54a0-133">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="b54a0-133">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="b54a0-134">在任务窗格中，选择 **“设置颜色”** 按钮，将选定区域的颜色设置为l蓝色。</span><span class="sxs-lookup"><span data-stu-id="b54a0-134">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="b54a0-136">后续步骤</span><span class="sxs-lookup"><span data-stu-id="b54a0-136">Next steps</span></span>

<span data-ttu-id="b54a0-p106">恭喜！已使用 React 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="b54a0-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b54a0-139">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="b54a0-139">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="b54a0-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b54a0-140">See also</span></span>

* [<span data-ttu-id="b54a0-141">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="b54a0-141">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="b54a0-142">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="b54a0-142">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="b54a0-143">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="b54a0-143">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="b54a0-144">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="b54a0-144">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
