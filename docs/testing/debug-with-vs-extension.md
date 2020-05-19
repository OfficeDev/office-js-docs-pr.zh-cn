---
title: 适用于 Visual Studio Code 的 Microsoft Office 外接程序调试器扩展
description: 使用 Visual Studio Code extension Microsoft Office 加载项调试器调试 Office 外接程序。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 57a59029ee9bb9791829d9d3583ce8b85e417b16
ms.sourcegitcommit: 71a44405e42b4798a8354f7f96d84548ae7a00f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44280358"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="062ed-103">适用于 Visual Studio Code 的 Microsoft Office 外接程序调试器扩展</span><span class="sxs-lookup"><span data-stu-id="062ed-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="062ed-104">通过 Visual Studio Code 的 Microsoft Office 外接程序调试器扩展，你可以针对边缘运行时调试 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="062ed-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="062ed-105">此调试模式是动态的，允许您在代码运行时设置断点。</span><span class="sxs-lookup"><span data-stu-id="062ed-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="062ed-106">在调试器附加时，您可以立即看到代码中的更改，而不会丢失您的调试会话。</span><span class="sxs-lookup"><span data-stu-id="062ed-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="062ed-107">您的代码更改也会保留，以便您可以看到对代码进行多个更改的结果。</span><span class="sxs-lookup"><span data-stu-id="062ed-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="062ed-108">下图显示了此扩展在操作中。</span><span class="sxs-lookup"><span data-stu-id="062ed-108">The following image shows this extension in action.</span></span>

![Office Addin 调试器扩展调试 Excel 外接程序的某个部分](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="062ed-110">先决条件</span><span class="sxs-lookup"><span data-stu-id="062ed-110">Prerequisites</span></span>

- <span data-ttu-id="062ed-111">[Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）</span><span class="sxs-lookup"><span data-stu-id="062ed-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="062ed-112">Node.js （版本 10 +）</span><span class="sxs-lookup"><span data-stu-id="062ed-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="062ed-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="062ed-113">Windows 10</span></span>
- [<span data-ttu-id="062ed-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="062ed-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="062ed-115">这些说明假定您有使用命令行的经验，了解基本 JavaScript，并已在使用 Yo Office 生成器之前创建了 Office 外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="062ed-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="062ed-116">如果你之前未执行此操作，请考虑访问我们的一个教程，如此[Excel Office 外接教程教程](../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="062ed-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="062ed-117">安装和使用调试器</span><span class="sxs-lookup"><span data-stu-id="062ed-117">Install and use the debugger</span></span>

1. <span data-ttu-id="062ed-118">如果需要创建外接程序项目，请[使用 Yo Office 生成器创建一个](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator)外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="062ed-118">If you need to create an add-in project, [use the Yo Office generator to create one](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span></span> <span data-ttu-id="062ed-119">按照命令行中的提示设置项目。</span><span class="sxs-lookup"><span data-stu-id="062ed-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="062ed-120">您可以根据需要选择任意语言或项目类型。</span><span class="sxs-lookup"><span data-stu-id="062ed-120">You can choose any language or type of project to suit your needs.</span></span>

> <span data-ttu-id="062ed-121">!便笺如果已有一个项目，请跳过步骤1并转到步骤2。</span><span class="sxs-lookup"><span data-stu-id="062ed-121">![NOTE] If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="062ed-122">以管理员身份打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="062ed-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="062ed-123">![Windows 10 中的命令提示符选项，包括 "以管理员身份运行"](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="062ed-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="062ed-124">导航到您的项目目录。</span><span class="sxs-lookup"><span data-stu-id="062ed-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="062ed-125">运行以下命令，以管理员身份在 Visual Studio Code 中打开项目。</span><span class="sxs-lookup"><span data-stu-id="062ed-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="062ed-126">在 Visual Studio Code 打开后，手动导航到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="062ed-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="062ed-127">若要以管理员身份打开 Visual Studio Code，请选择 "以**管理员身份运行**" 选项，在 Windows 中搜索 Visual studio code 之后打开它。</span><span class="sxs-lookup"><span data-stu-id="062ed-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="062ed-128">在 VS 代码中，选择**CTRL + SHIFT + X**打开扩展栏。</span><span class="sxs-lookup"><span data-stu-id="062ed-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="062ed-129">搜索 "Microsoft Office 外接程序调试器" 扩展并安装它。</span><span class="sxs-lookup"><span data-stu-id="062ed-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="062ed-130">在项目的 ". vscode" 文件夹中，打开 "**启动. json** " 文件。</span><span class="sxs-lookup"><span data-stu-id="062ed-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="062ed-131">将以下代码添加到 `configurations` 部分：</span><span class="sxs-lookup"><span data-stu-id="062ed-131">Add the following code to the `configurations` section:</span></span>

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. <span data-ttu-id="062ed-132">在刚刚复制的 JSON 部分中，找到 "url" 部分。</span><span class="sxs-lookup"><span data-stu-id="062ed-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="062ed-133">在此 URL 中，需要将大写的主机文本替换为 Office 加载项的主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="062ed-133">In this URL, you will need to replace the uppercase HOST text with the host application for your Office add-in.</span></span> <span data-ttu-id="062ed-134">例如，如果您的 Office 外接程序适用于 excel，则 URL 值将为 " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $ 16.01 $ en-us $ \$ \$ \$ 0"。</span><span class="sxs-lookup"><span data-stu-id="062ed-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="062ed-135">打开命令提示符，并确保您在项目的根文件夹中。</span><span class="sxs-lookup"><span data-stu-id="062ed-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="062ed-136">运行命令 `npm start` 以启动开发服务器。</span><span class="sxs-lookup"><span data-stu-id="062ed-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="062ed-137">当加载项在 Office 客户端中加载时，打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="062ed-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="062ed-138">返回到 Visual Studio Code，然后选择 "**查看 > 调试**" 或 enter **CTRL + SHIFT + D**切换到 "调试" 视图。</span><span class="sxs-lookup"><span data-stu-id="062ed-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="062ed-139">从 "调试" 选项中，选择 "**附加到 Office 外接程序**"。从菜单中选择 " **F5** " 或选择 "**调试-> 启动调试**" 以开始调试。</span><span class="sxs-lookup"><span data-stu-id="062ed-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="062ed-140">在项目的任务窗格文件中设置断点。</span><span class="sxs-lookup"><span data-stu-id="062ed-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="062ed-141">您可以通过悬停在代码行旁边并选择显示的红色圆圈，在 VS 代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="062ed-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![对 VS 代码中的一行代码显示红色圆圈](../images/set-breakpoint.jpg)

12. <span data-ttu-id="062ed-143">运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="062ed-143">Run your add-in.</span></span> <span data-ttu-id="062ed-144">您将看到断点已命中，您可以检查局部变量。</span><span class="sxs-lookup"><span data-stu-id="062ed-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="062ed-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="062ed-145">See also</span></span>

* [<span data-ttu-id="062ed-146">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="062ed-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="062ed-147">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="062ed-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="062ed-148">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="062ed-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
