---
title: 使用 React 生成 Excel 任务窗格加载项
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 1c0f2f4af1ee14aaf7d581509733e26013657590
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308027"
---
# <a name="build-an-excel-task-pane-add-in-using-react"></a><span data-ttu-id="a6935-102">使用 React 生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="a6935-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="a6935-103">本文将逐步介绍如何使用 React 和 Excel JavaScript API 生成 Excel 任务加载项。</span><span class="sxs-lookup"><span data-stu-id="a6935-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a6935-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="a6935-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="a6935-105">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="a6935-105">Create the add-in project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="a6935-106">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="a6935-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="a6935-107">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="a6935-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="a6935-108">**选择项目类型:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="a6935-108">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="a6935-109">**选择脚本类型:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="a6935-109">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="a6935-110">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="a6935-110">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="a6935-111">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="a6935-111">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="a6935-112">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="a6935-112">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="a6935-113">浏览项目</span><span class="sxs-lookup"><span data-stu-id="a6935-113">Explore the project</span></span>

<span data-ttu-id="a6935-114">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="a6935-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="a6935-115">如果想要浏览加载项项目的主要组件，请在代码编辑器中打开项目并检查下面列出的文件。</span><span class="sxs-lookup"><span data-stu-id="a6935-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="a6935-116">准备好试用加载项时，请转至下一部分。</span><span class="sxs-lookup"><span data-stu-id="a6935-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="a6935-117">项目根目录中的 **manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="a6935-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="a6935-118">**./src/taskpane/taskpane.html** 文件定义任务窗格的 HTML 框架，而 **./src/taskpane/components** 文件夹内的文件定义任务窗格 UI 的各个部分。</span><span class="sxs-lookup"><span data-stu-id="a6935-118">The **./src/taskpane/taskpane.html** file defines the HTML framework of the task pane, and the files within the **./src/taskpane/components** folder define the various parts of the task pane UI.</span></span>
- <span data-ttu-id="a6935-119">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="a6935-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="a6935-120">**./src/taskpane/components/App.tsx** 文件包含用于加快任务窗格与 Excel 之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="a6935-120">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="a6935-121">试用</span><span class="sxs-lookup"><span data-stu-id="a6935-121">Try it out</span></span>

1. <span data-ttu-id="a6935-122">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="a6935-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="a6935-123">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="a6935-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="a6935-125">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="a6935-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="a6935-126">在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。</span><span class="sxs-lookup"><span data-stu-id="a6935-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="a6935-128">后续步骤</span><span class="sxs-lookup"><span data-stu-id="a6935-128">Next steps</span></span>

<span data-ttu-id="a6935-129">祝贺，你已使用 React 成功创建了 Excel 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="a6935-129">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="a6935-130">接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="a6935-130">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a6935-131">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="a6935-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="a6935-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a6935-132">See also</span></span>

* [<span data-ttu-id="a6935-133">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="a6935-133">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="a6935-134">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="a6935-134">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="a6935-135">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="a6935-135">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="a6935-136">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="a6935-136">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
