# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="3df1b-101">使用 React 生成 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="3df1b-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="3df1b-102">在本文中，你将完成使用 React 和 Excel JavaScript API 生成 Excel 加载项的过程。</span><span class="sxs-lookup"><span data-stu-id="3df1b-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3df1b-103">先决条件</span><span class="sxs-lookup"><span data-stu-id="3df1b-103">Prerequisites</span></span>

- [<span data-ttu-id="3df1b-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="3df1b-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="3df1b-105">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="3df1b-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="3df1b-106">创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="3df1b-106">Create the web app</span></span>

1. <span data-ttu-id="3df1b-107">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="3df1b-107">Use the Yeoman generator to create an Outlook add-in project.</span></span> <span data-ttu-id="3df1b-108">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="3df1b-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="3df1b-109">**选择项目类型:** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="3df1b-109">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="3df1b-110">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="3df1b-110">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="3df1b-111">**要支持哪一个 Office 客户端应用?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="3df1b-111">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman 生成器](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="3df1b-113">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="3df1b-114">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="3df1b-114">Navigate to the root folder of the project in the Terminal app, and from Terminal run:</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="3df1b-115">更新代码</span><span class="sxs-lookup"><span data-stu-id="3df1b-115">Update the code</span></span>

1. <span data-ttu-id="3df1b-116">在代码编辑器中，打开文件 **src/styles.less**，将以下样式添加到文件的末尾，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-116">In your code editor, open the file **src/styles.less**, add the following styles to the end of the file, and save the file.</span></span>

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
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. <span data-ttu-id="3df1b-117">Office 加载项 Yeoman 生成器创建的项目模板包含此快速入门不需要的 React 组件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-117">The project template that the Office Add-ins Yeoman generator created includes a React component that is not needed for this quick start.</span></span> <span data-ttu-id="3df1b-118">删除文件 **src/components/HeroList.tsx**。</span><span class="sxs-lookup"><span data-stu-id="3df1b-118">Delete the file **src/components/HeroList.tsx**.</span></span>

3. <span data-ttu-id="3df1b-119">打开文件 **src/components/Header.tsx**，将全部内容替换为以下代码，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-119">Open the file **src/components/Header.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. <span data-ttu-id="3df1b-120">在 **src / components** 文件夹中创建一个名为 **Content.tsx** 的新 React 组件，添加以下代码，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-120">Create a new React component named **Content.tsx** in the **src/components** folder, add the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. <span data-ttu-id="3df1b-121">打开文件 **src/components/App.tsx**，将全部内容替换为以下代码，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-121">Open the file **src/components/App.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="3df1b-122">更新清单</span><span class="sxs-lookup"><span data-stu-id="3df1b-122">Update the manifest</span></span>

1. <span data-ttu-id="3df1b-123">打开文件 **manifest.xml** 以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="3df1b-123">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="3df1b-124">`ProviderName` 元素具有占位符值。</span><span class="sxs-lookup"><span data-stu-id="3df1b-124">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="3df1b-125">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="3df1b-125">Replace it with your name.</span></span>

3. <span data-ttu-id="3df1b-126">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="3df1b-126">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="3df1b-127">将它替换为“A task pane add-in for Excel”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3df1b-127">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="3df1b-128">保存文件。</span><span class="sxs-lookup"><span data-stu-id="3df1b-128">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="3df1b-129">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="3df1b-129">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="3df1b-130">试用</span><span class="sxs-lookup"><span data-stu-id="3df1b-130">Try it out</span></span>

1. <span data-ttu-id="3df1b-131">请按照运行加载项所用平台对应的说明操作，以在 Excel 中旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="3df1b-131">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="3df1b-132">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="3df1b-132">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="3df1b-133">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="3df1b-133">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="3df1b-134">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="3df1b-134">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="3df1b-135">在 Excel 中，依次选择**主页**选项卡和功能区中的**显示任务窗格**按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="3df1b-135">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="3df1b-137">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="3df1b-137">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="3df1b-138">在任务窗格中，选择**设置颜色**按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="3df1b-138">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="3df1b-140">后续步骤</span><span class="sxs-lookup"><span data-stu-id="3df1b-140">Next steps</span></span>

<span data-ttu-id="3df1b-p105">恭喜！已使用 React 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="3df1b-p105">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3df1b-143">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="3df1b-143">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="3df1b-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3df1b-144">See also</span></span>

* [<span data-ttu-id="3df1b-145">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="3df1b-145">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="3df1b-146">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="3df1b-146">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="3df1b-147">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="3df1b-147">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="3df1b-148">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="3df1b-148">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
