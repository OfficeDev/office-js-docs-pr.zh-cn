---
title: 使用 Angular 生成 Excel 任务窗格加载项
description: ''
ms.date: 12/24/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: c2565c305833335e589a53655397bcb68654d22a
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851556"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="5d8e0-102">使用 Angular 生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="5d8e0-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="5d8e0-103">本文将逐步介绍如何使用 Angular 和 Excel JavaScript API 生成 Excel 任务加载项。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5d8e0-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="5d8e0-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="5d8e0-105">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="5d8e0-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="5d8e0-106">**选择项目类型:** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="5d8e0-106">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="5d8e0-107">**选择脚本类型:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="5d8e0-107">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="5d8e0-108">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="5d8e0-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="5d8e0-109">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="5d8e0-109">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman 生成器](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="5d8e0-111">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="5d8e0-112">浏览项目</span><span class="sxs-lookup"><span data-stu-id="5d8e0-112">Explore the project</span></span>

<span data-ttu-id="5d8e0-113">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="5d8e0-114">如果想要浏览加载项项目的主要组件，请在代码编辑器中打开项目并检查下面列出的文件。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-114">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="5d8e0-115">准备好试用加载项时，请转至下一部分。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-115">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="5d8e0-116">项目根目录中的 **manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-116">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="5d8e0-117">**./src/taskpane/app/app.component.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-117">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="5d8e0-118">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="5d8e0-119">**./src/taskpane/app/app.component.ts** 文件包含用于加快任务窗格与 Excel 之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-119">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="5d8e0-120">试用</span><span class="sxs-lookup"><span data-stu-id="5d8e0-120">Try it out</span></span>

1. <span data-ttu-id="5d8e0-121">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-121">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="5d8e0-122">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-122">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="5d8e0-124">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-124">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="5d8e0-125">在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-125">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="5d8e0-127">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5d8e0-127">Next steps</span></span>

<span data-ttu-id="5d8e0-128">祝贺，你已使用 Angular 成功创建了 Excel 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="5d8e0-128">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="5d8e0-129">接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="5d8e0-129">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5d8e0-130">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="5d8e0-130">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="5d8e0-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5d8e0-131">See also</span></span>

* [<span data-ttu-id="5d8e0-132">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="5d8e0-132">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="5d8e0-133">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="5d8e0-133">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
* <span data-ttu-id="5d8e0-134">[开发 Office 加载项](../develop/develop-overview.md)</span><span class="sxs-lookup"><span data-stu-id="5d8e0-134">[](../develop/develop-overview.md)Develop Office Add-ins with Angular</span></span>
* [<span data-ttu-id="5d8e0-135">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="5d8e0-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="5d8e0-136">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="5d8e0-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="5d8e0-137">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="5d8e0-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)