---
title: 使用 Angular 生成 Excel 任务窗格加载项
description: ''
ms.date: 09/06/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: ed805cf5d19a38d543b7fcbba49508dd3f2d6d97
ms.sourcegitcommit: ce7e7087a4550b9c090dc565fee5eac08a2985a2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/06/2019
ms.locfileid: "36782252"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="89b84-102">使用 Angular 生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="89b84-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="89b84-103">本文将逐步介绍如何使用 Angular 和 Excel JavaScript API 生成 Excel 任务加载项。</span><span class="sxs-lookup"><span data-stu-id="89b84-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="89b84-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="89b84-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="89b84-105">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="89b84-105">Create the add-in project</span></span>

<span data-ttu-id="89b84-106">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="89b84-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="89b84-107">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="89b84-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="89b84-108">**选择项目类型:** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="89b84-108">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="89b84-109">**选择脚本类型:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="89b84-109">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="89b84-110">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="89b84-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="89b84-111">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="89b84-111">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman 生成器](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="89b84-113">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="89b84-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="89b84-114">浏览项目</span><span class="sxs-lookup"><span data-stu-id="89b84-114">Explore the project</span></span>

<span data-ttu-id="89b84-115">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="89b84-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="89b84-116">如果想要浏览加载项项目的主要组件，请在代码编辑器中打开项目并检查下面列出的文件。</span><span class="sxs-lookup"><span data-stu-id="89b84-116">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="89b84-117">准备好试用加载项时，请转至下一部分。</span><span class="sxs-lookup"><span data-stu-id="89b84-117">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="89b84-118">项目根目录中的 **manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="89b84-118">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="89b84-119">**./src/taskpane/app/app.component.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="89b84-119">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="89b84-120">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="89b84-120">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="89b84-121">**./src/taskpane/app/app.component.ts** 文件包含用于加快任务窗格与 Excel 之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="89b84-121">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="89b84-122">试用</span><span class="sxs-lookup"><span data-stu-id="89b84-122">Try it out</span></span>

1. <span data-ttu-id="89b84-123">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="89b84-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="89b84-124">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="89b84-124">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="89b84-126">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="89b84-126">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="89b84-127">在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。</span><span class="sxs-lookup"><span data-stu-id="89b84-127">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="89b84-129">后续步骤</span><span class="sxs-lookup"><span data-stu-id="89b84-129">Next steps</span></span>

<span data-ttu-id="89b84-130">祝贺，你已使用 Angular 成功创建了 Excel 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="89b84-130">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="89b84-131">接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="89b84-131">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="89b84-132">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="89b84-132">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="89b84-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="89b84-133">See also</span></span>

* [<span data-ttu-id="89b84-134">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="89b84-134">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="89b84-135">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="89b84-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="89b84-136">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="89b84-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="89b84-137">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="89b84-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
