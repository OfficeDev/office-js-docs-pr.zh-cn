---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用Visual Studio Code调试Microsoft Office调试器中的扩展Office调试外接程序。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 264a5d43a8b4f0faf7d6216664d30d7c8b64cccc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077118"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="79f1e-103">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="79f1e-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="79f1e-104">Microsoft Office外接程序调试器扩展 for Visual Studio Code 允许你使用原始 webView Microsoft Edge EdgeHTML Microsoft Edge运行时调试 Office 外接程序 (调试) 外接程序。</span><span class="sxs-lookup"><span data-stu-id="79f1e-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="79f1e-105">有关针对基于 WebView2 Microsoft Edge (Chromium进行) 的说明，[请参阅本文](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="79f1e-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="79f1e-106">此调试模式是动态的，允许在代码运行时设置断点。</span><span class="sxs-lookup"><span data-stu-id="79f1e-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="79f1e-107">在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。</span><span class="sxs-lookup"><span data-stu-id="79f1e-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="79f1e-108">代码更改也持续存在，因此可以看到对代码进行多次更改的结果。</span><span class="sxs-lookup"><span data-stu-id="79f1e-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="79f1e-109">下图显示了此扩展的操作。</span><span class="sxs-lookup"><span data-stu-id="79f1e-109">The following image shows this extension in action.</span></span>

![Office加载项调试器扩展调试加载项Excel部分。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="79f1e-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="79f1e-111">Prerequisites</span></span>

- <span data-ttu-id="79f1e-112">[Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）</span><span class="sxs-lookup"><span data-stu-id="79f1e-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="79f1e-113">Node.js （版本 10+）</span><span class="sxs-lookup"><span data-stu-id="79f1e-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="79f1e-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="79f1e-114">Windows 10</span></span>
- [<span data-ttu-id="79f1e-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="79f1e-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="79f1e-116">这些说明假定你拥有使用命令行的经验，了解基本 JavaScript，并且已创建一个 Office 加载项项目，然后才使用 Yo Office 生成器。</span><span class="sxs-lookup"><span data-stu-id="79f1e-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="79f1e-117">如果你之前没有这样做，请考虑访问我们的教程之一，Excel Office[外接程序教程](../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="79f1e-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="79f1e-118">安装和使用调试器</span><span class="sxs-lookup"><span data-stu-id="79f1e-118">Install and use the debugger</span></span>

1. <span data-ttu-id="79f1e-119">如果需要创建加载项项目，请使用[Yo Office生成器创建一个](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。</span><span class="sxs-lookup"><span data-stu-id="79f1e-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="79f1e-120">按照命令行中的提示设置项目。</span><span class="sxs-lookup"><span data-stu-id="79f1e-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="79f1e-121">可以选择任何语言或项目类型以满足你的需求。</span><span class="sxs-lookup"><span data-stu-id="79f1e-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="79f1e-122">如果已有项目，请跳过步骤 1 并移至步骤 2。</span><span class="sxs-lookup"><span data-stu-id="79f1e-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="79f1e-123">以管理员角色打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="79f1e-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="79f1e-124">![命令提示符选项，包括"以管理员Windows 10。](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="79f1e-124">![Command prompt options, including "run as administrator" in Windows 10.](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="79f1e-125">导航到项目目录。</span><span class="sxs-lookup"><span data-stu-id="79f1e-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="79f1e-126">运行以下命令以管理员Visual Studio Code打开项目。</span><span class="sxs-lookup"><span data-stu-id="79f1e-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="79f1e-127">打开Visual Studio Code后，手动导航到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="79f1e-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="79f1e-128">若要以Visual Studio Code方式打开文件，请选择"以管理员方式运行"选项，Visual Studio Code中搜索后打开Windows。</span><span class="sxs-lookup"><span data-stu-id="79f1e-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="79f1e-129">在 VS 代码中，选择 **CTRL + SHIFT + X** 打开扩展栏。</span><span class="sxs-lookup"><span data-stu-id="79f1e-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="79f1e-130">搜索"Microsoft Office加载项调试器"扩展并安装它。</span><span class="sxs-lookup"><span data-stu-id="79f1e-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="79f1e-131">在你的项目 .vscode 文件夹中打开 **launch.json** 文件。</span><span class="sxs-lookup"><span data-stu-id="79f1e-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="79f1e-132">将以下代码添加到 `configurations` 部分：</span><span class="sxs-lookup"><span data-stu-id="79f1e-132">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="79f1e-133">在刚刚复制的 JSON 部分中，找到"url"部分。</span><span class="sxs-lookup"><span data-stu-id="79f1e-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="79f1e-134">在此 URL 中，您需要将大写的 HOST 文本替换为托管您的外接程序Office应用程序。</span><span class="sxs-lookup"><span data-stu-id="79f1e-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="79f1e-135">例如，如果Office外接程序用于 Excel，则 URL 值将是 https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0"。</span><span class="sxs-lookup"><span data-stu-id="79f1e-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="79f1e-136">打开命令提示符，并确保位于项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="79f1e-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="79f1e-137">运行命令 `npm start` 以启动开发服务器。</span><span class="sxs-lookup"><span data-stu-id="79f1e-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="79f1e-138">当加载项在客户端Office时，打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="79f1e-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="79f1e-139">返回到"Visual Studio Code并选择"查看 **>调试"** 或输入 **Ctrl + Shift + D** 以切换到调试视图。</span><span class="sxs-lookup"><span data-stu-id="79f1e-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="79f1e-140">从"调试"选项中，选择"**附加到Office加载项"。** 从 **菜单中选择 F5** 或 **>** 调试 -开始调试"开始调试。</span><span class="sxs-lookup"><span data-stu-id="79f1e-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="79f1e-141">在项目的任务窗格文件中设置断点。</span><span class="sxs-lookup"><span data-stu-id="79f1e-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="79f1e-142">通过将鼠标悬停在代码行Visual Studio Code并选择出现的红色圆圈，可以在代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="79f1e-142">You can set breakpoints in Visual Studio Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![在代码行中出现红色圆圈Visual Studio Code。](../images/set-breakpoint.jpg)

12. <span data-ttu-id="79f1e-144">运行加载项。</span><span class="sxs-lookup"><span data-stu-id="79f1e-144">Run your add-in.</span></span> <span data-ttu-id="79f1e-145">你将看到已命中的断点，并且你可以检查本地变量。</span><span class="sxs-lookup"><span data-stu-id="79f1e-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="79f1e-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="79f1e-146">See also</span></span>

* [<span data-ttu-id="79f1e-147">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="79f1e-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="79f1e-148">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="79f1e-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="79f1e-149">使用 Windows 上的 Microsoft Edge WebView2 （基于 Chromium）调试加载项</span><span class="sxs-lookup"><span data-stu-id="79f1e-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
