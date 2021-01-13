---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用Visual Studio调试Microsoft Office代码扩展来调试 Office 外接程序。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 83791d5d60238288e3059809b8b8c02b1f4f768f
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840109"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="03329-103">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="03329-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="03329-104">通过Microsoft Office外接程序调试器扩展Visual Studio代码，您可以针对边缘运行时调试 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="03329-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="03329-105">此调试模式是动态的，允许您在代码运行时设置断点。</span><span class="sxs-lookup"><span data-stu-id="03329-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="03329-106">在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。</span><span class="sxs-lookup"><span data-stu-id="03329-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="03329-107">代码更改也会持续存在，因此你可以看到对代码进行多次更改的结果。</span><span class="sxs-lookup"><span data-stu-id="03329-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="03329-108">下图显示了此扩展的操作。</span><span class="sxs-lookup"><span data-stu-id="03329-108">The following image shows this extension in action.</span></span>

![Office 加载项调试程序扩展调试 Excel 加载项的一部分](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="03329-110">先决条件</span><span class="sxs-lookup"><span data-stu-id="03329-110">Prerequisites</span></span>

- <span data-ttu-id="03329-111">[Visual Studio必须](https://code.visualstudio.com/) (管理员角色运行代码) </span><span class="sxs-lookup"><span data-stu-id="03329-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="03329-112">Node.js (版本 10 及以上版本) </span><span class="sxs-lookup"><span data-stu-id="03329-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="03329-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="03329-113">Windows 10</span></span>
- [<span data-ttu-id="03329-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="03329-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="03329-115">这些说明假定你具有使用命令行的经验，了解基本 JavaScript，并且已创建 Office 加载项项目，然后再使用 Yo Office 生成器。</span><span class="sxs-lookup"><span data-stu-id="03329-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="03329-116">如果之前尚未这样做，请考虑访问我们的教程之一，如此 Excel Office [加载项教程](../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="03329-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="03329-117">安装和使用调试器</span><span class="sxs-lookup"><span data-stu-id="03329-117">Install and use the debugger</span></span>

1. <span data-ttu-id="03329-118">如果需要创建加载项项目，请使用 Yo Office 生成器 [创建一个](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。</span><span class="sxs-lookup"><span data-stu-id="03329-118">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="03329-119">按照命令行中的提示设置项目。</span><span class="sxs-lookup"><span data-stu-id="03329-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="03329-120">可以选择任何语言或项目类型以满足你的需求。</span><span class="sxs-lookup"><span data-stu-id="03329-120">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="03329-121">如果已有项目，请跳过步骤 1 并移至步骤 2。</span><span class="sxs-lookup"><span data-stu-id="03329-121">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="03329-122">以管理员角色打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="03329-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="03329-123">![命令提示符选项，包括 Windows 10 中的"以管理员方式运行"](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="03329-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="03329-124">导航到项目目录。</span><span class="sxs-lookup"><span data-stu-id="03329-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="03329-125">运行以下命令以管理员Visual Studio代码打开项目。</span><span class="sxs-lookup"><span data-stu-id="03329-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="03329-126">打开Visual Studio代码后，手动导航到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="03329-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="03329-127">若要以Visual Studio方式打开"代码"，请在 Windows中搜索 Visual Studio代码后，选择"以管理员方式运行"选项。</span><span class="sxs-lookup"><span data-stu-id="03329-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="03329-128">在 VS Code 中，选择 **Ctrl + Shift + X** 以打开扩展栏。</span><span class="sxs-lookup"><span data-stu-id="03329-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="03329-129">搜索"Microsoft Office调试器"扩展并安装它。</span><span class="sxs-lookup"><span data-stu-id="03329-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="03329-130">在项目的 .vscode 文件夹中，打开launch.js **文件** 。</span><span class="sxs-lookup"><span data-stu-id="03329-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="03329-131">将以下代码添加到 `configurations` 此部分：</span><span class="sxs-lookup"><span data-stu-id="03329-131">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="03329-132">在刚复制的 JSON 部分中，找到"url"部分。</span><span class="sxs-lookup"><span data-stu-id="03329-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="03329-133">在此 URL 中，您需要将大写的 HOST 文本替换为托管 Office 外接程序的应用程序。</span><span class="sxs-lookup"><span data-stu-id="03329-133">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office add-in.</span></span> <span data-ttu-id="03329-134">例如，如果您的 Office 外接程序适用于 Excel，则 URL 值为 https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0"。</span><span class="sxs-lookup"><span data-stu-id="03329-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="03329-135">打开命令提示符并确保你位于项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="03329-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="03329-136">运行命令 `npm start` 以启动开发服务器。</span><span class="sxs-lookup"><span data-stu-id="03329-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="03329-137">当加载项在 Office 客户端中加载时，打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="03329-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="03329-138">返回到Visual Studio代码，然后选择"> **调试** "或输入 **Ctrl + Shift + D** 以切换到调试视图。</span><span class="sxs-lookup"><span data-stu-id="03329-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="03329-139">从"调试"选项中，选择 **"附加到 Office 加载项"。** 选择 **F5** 或选择>开始 **调试** "-开始调试"。</span><span class="sxs-lookup"><span data-stu-id="03329-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="03329-140">在项目的任务窗格文件中设置断点。</span><span class="sxs-lookup"><span data-stu-id="03329-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="03329-141">通过在代码行旁边悬停并选择出现的红色圆圈，可以在 VS Code 中设置断点。</span><span class="sxs-lookup"><span data-stu-id="03329-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![VS Code 中的一行代码上出现一个红色圆圈](../images/set-breakpoint.jpg)

12. <span data-ttu-id="03329-143">运行加载项。</span><span class="sxs-lookup"><span data-stu-id="03329-143">Run your add-in.</span></span> <span data-ttu-id="03329-144">你将看到已命中断点，你可以检查本地变量。</span><span class="sxs-lookup"><span data-stu-id="03329-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="03329-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="03329-145">See also</span></span>

* [<span data-ttu-id="03329-146">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="03329-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="03329-147">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="03329-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="03329-148">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="03329-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)