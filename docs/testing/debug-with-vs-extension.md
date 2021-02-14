---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用Visual Studio调试器Microsoft Office代码扩展来调试 Office 外接程序。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237929"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="6b969-103">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="6b969-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="6b969-104">借助 Microsoft Office 外接程序调试器扩展 for Visual Studio Code，您可以使用原始 WebView (EdgeHTML) 运行时针对 Microsoft Edge 调试 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="6b969-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="6b969-105">有关针对基于 Chromium (Microsoft Edge WebView2 进行) 的说明，请参阅 [本文](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="6b969-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="6b969-106">此调试模式是动态的，允许您在代码运行时设置断点。</span><span class="sxs-lookup"><span data-stu-id="6b969-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="6b969-107">在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。</span><span class="sxs-lookup"><span data-stu-id="6b969-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="6b969-108">代码更改也会持续存在，因此你可以看到对代码进行多次更改的结果。</span><span class="sxs-lookup"><span data-stu-id="6b969-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="6b969-109">下图显示了此扩展的操作。</span><span class="sxs-lookup"><span data-stu-id="6b969-109">The following image shows this extension in action.</span></span>

![Office 加载项调试程序扩展调试 Excel 加载项的一部分](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="6b969-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="6b969-111">Prerequisites</span></span>

- <span data-ttu-id="6b969-112">[Visual Studio必须](https://code.visualstudio.com/) (管理员角色运行代码) </span><span class="sxs-lookup"><span data-stu-id="6b969-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="6b969-113">Node.js (版本 10+) </span><span class="sxs-lookup"><span data-stu-id="6b969-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="6b969-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="6b969-114">Windows 10</span></span>
- [<span data-ttu-id="6b969-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="6b969-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="6b969-116">这些说明假定你具有使用命令行的经验，了解基本 JavaScript，并且已使用 Yo Office 生成器之前创建了 Office 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="6b969-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="6b969-117">如果之前尚未这样做，请考虑访问我们的教程之一，如本 Excel Office [加载项教程](../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="6b969-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="6b969-118">安装和使用调试器</span><span class="sxs-lookup"><span data-stu-id="6b969-118">Install and use the debugger</span></span>

1. <span data-ttu-id="6b969-119">如果需要创建加载项项目，请使用 Yo Office 生成器 [创建一个](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。</span><span class="sxs-lookup"><span data-stu-id="6b969-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="6b969-120">按照命令行中的提示设置项目。</span><span class="sxs-lookup"><span data-stu-id="6b969-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="6b969-121">可以选择任何语言或项目类型来满足您的需求。</span><span class="sxs-lookup"><span data-stu-id="6b969-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="6b969-122">如果已有项目，请跳过步骤 1 并移动到步骤 2。</span><span class="sxs-lookup"><span data-stu-id="6b969-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="6b969-123">以管理员角色打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="6b969-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="6b969-124">![命令提示符选项，包括 Windows 10 中的"以管理员方式运行"](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="6b969-124">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="6b969-125">导航到项目目录。</span><span class="sxs-lookup"><span data-stu-id="6b969-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="6b969-126">运行以下命令以管理员Visual Studio代码打开项目。</span><span class="sxs-lookup"><span data-stu-id="6b969-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="6b969-127">打开Visual Studio后，手动导航到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="6b969-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="6b969-128">若要以Visual Studio方式打开代码，请在 Windows 中搜索代码后Visual Studio代码时选择"以管理员方式运行"选项。</span><span class="sxs-lookup"><span data-stu-id="6b969-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="6b969-129">在 VS Code 中，选择 **Ctrl + Shift + X** 以打开扩展栏。</span><span class="sxs-lookup"><span data-stu-id="6b969-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="6b969-130">搜索"Microsoft Office调试器"扩展并安装它。</span><span class="sxs-lookup"><span data-stu-id="6b969-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="6b969-131">在项目的 .vscode 文件夹中，打开launch.js **文件。**</span><span class="sxs-lookup"><span data-stu-id="6b969-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="6b969-132">将以下代码添加到 `configurations` 该部分：</span><span class="sxs-lookup"><span data-stu-id="6b969-132">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="6b969-133">在刚复制的 JSON 部分中，查找"url"部分。</span><span class="sxs-lookup"><span data-stu-id="6b969-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="6b969-134">在此 URL 中，您需要将大写的 HOST 文本替换为托管 Office 外接程序的应用程序。</span><span class="sxs-lookup"><span data-stu-id="6b969-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="6b969-135">例如，如果 Office 外接程序适用于 Excel，则 URL 值为 https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0"。</span><span class="sxs-lookup"><span data-stu-id="6b969-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="6b969-136">打开命令提示符，并确保你位于项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="6b969-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="6b969-137">运行命令 `npm start` 以启动开发服务器。</span><span class="sxs-lookup"><span data-stu-id="6b969-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="6b969-138">当加载项在 Office 客户端中加载时，打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="6b969-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="6b969-139">返回到Visual Studio代码，然后选择 **">调试** "或输入 **Ctrl + Shift + D** 以切换到调试视图。</span><span class="sxs-lookup"><span data-stu-id="6b969-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="6b969-140">从"调试"选项中，选择 **"附加到 Office 加载项"。** 选择 **F5** 或从>开始 **调试** 以开始调试。</span><span class="sxs-lookup"><span data-stu-id="6b969-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="6b969-141">在项目的任务窗格文件中设置断点。</span><span class="sxs-lookup"><span data-stu-id="6b969-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="6b969-142">通过在代码行旁边悬停并选择出现的红色圆圈，可以在 VS Code 中设置断点。</span><span class="sxs-lookup"><span data-stu-id="6b969-142">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![VS Code 中的一行代码上显示一个红色圆圈](../images/set-breakpoint.jpg)

12. <span data-ttu-id="6b969-144">运行加载项。</span><span class="sxs-lookup"><span data-stu-id="6b969-144">Run your add-in.</span></span> <span data-ttu-id="6b969-145">你将看到断点已命中，你可以检查本地变量。</span><span class="sxs-lookup"><span data-stu-id="6b969-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="6b969-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6b969-146">See also</span></span>

* [<span data-ttu-id="6b969-147">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6b969-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="6b969-148">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="6b969-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="6b969-149">使用 Microsoft Edge WebView2 和基于 Chromium (Windows 调试加载项) </span><span class="sxs-lookup"><span data-stu-id="6b969-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
