---
title: 使用 Windows 上的 Microsoft Edge WebView2 （基于 Chromium）调试加载项
description: 了解如何在 VS 代码中使用适用于 Microsoft Edge 扩展的调试器来调试使用 Microsoft Edge WebView2（基于 Chromium）的 Office 加载项。
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 0908bb5040b49568006324600acacb5e36dbd1a5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238112"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a><span data-ttu-id="981bd-103">使用 Windows 上的 Microsoft Edge Chromium WebView2 调试加载项</span><span class="sxs-lookup"><span data-stu-id="981bd-103">Debug add-ins on Windows using Edge Chromium WebView2</span></span>

<span data-ttu-id="981bd-104">在 Windows 上正在运行的 Office 加载项可以使用 VS 代码中适用于 Microsoft Edge 扩展的调试器来对 Edge Chromium WebView2 运行时进行调试。</span><span class="sxs-lookup"><span data-stu-id="981bd-104">Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="981bd-105">先决条件</span><span class="sxs-lookup"><span data-stu-id="981bd-105">Prerequisites</span></span>

- <span data-ttu-id="981bd-106">[Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）</span><span class="sxs-lookup"><span data-stu-id="981bd-106">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="981bd-107">Node.js （版本 10+）</span><span class="sxs-lookup"><span data-stu-id="981bd-107">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="981bd-108">Windows 10</span><span class="sxs-lookup"><span data-stu-id="981bd-108">Windows 10</span></span>
- [<span data-ttu-id="981bd-109"> 适用于 Windows Insiders 的 Microsoft Edge Chromium</span><span class="sxs-lookup"><span data-stu-id="981bd-109">Microsoft Edge Chromium available to Windows Insiders</span></span>](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="981bd-110">安装和使用调试器</span><span class="sxs-lookup"><span data-stu-id="981bd-110">Install and use the debugger</span></span>

1. <span data-ttu-id="981bd-111">使用 [ 适用于 Office 加载项的 Yeoman 生成器 ](https://github.com/OfficeDev/generator-office) 创建项目。可以使用我们的任何一个快速入门指南，例如 [Outlook 加载项快速入门 ](../quickstarts/outlook-quickstart.md)，以做到这一点。</span><span class="sxs-lookup"><span data-stu-id="981bd-111">Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.</span></span>

> [!TIP]
> <span data-ttu-id="981bd-112">如果没有使用基于 Yeoman 生成器的加载项，需要调整一个注册表项。</span><span class="sxs-lookup"><span data-stu-id="981bd-112">If you aren't using a Yeoman generator based add-in, you need to adjust a registry key.</span></span> <span data-ttu-id="981bd-113">在你的项目根目录下，在命令行中运行以下命令： `office-add-in-debugging start <your manifest path>`。</span><span class="sxs-lookup"><span data-stu-id="981bd-113">While in the root folder of your project, run the following in the command line: `office-add-in-debugging start <your manifest path>`.</span></span>

2. <span data-ttu-id="981bd-114">在 VS 代码中打开项目。</span><span class="sxs-lookup"><span data-stu-id="981bd-114">Open your project in VS Code.</span></span> <span data-ttu-id="981bd-115">在 VS 代码中，选择 **CTRL + SHIFT + X** 打开扩展栏。</span><span class="sxs-lookup"><span data-stu-id="981bd-115">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="981bd-116">搜索“适用于 Microsoft Edge 的调试器”扩展并安装。</span><span class="sxs-lookup"><span data-stu-id="981bd-116">Search for the "Debugger for Microsoft Edge" extension and install it.</span></span>

3. <span data-ttu-id="981bd-117">在你的项目 **.vscode** 文件夹中打开 **launch.json** 文件。</span><span class="sxs-lookup"><span data-stu-id="981bd-117">In the **.vscode** folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="981bd-118">将以下代码添加到配置节：</span><span class="sxs-lookup"><span data-stu-id="981bd-118">Add the following code to the configurations section:</span></span>

```JSON
  {
      "name": "Debug Office Add-in (Edge Chromium)",
      "type": "edge",
      "request": "attach",
      "useWebView": "advanced",
      "port": 9229,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
    },
```

4. <span data-ttu-id="981bd-119">下一步，选择 **View > Debug** 或者输入 **CTRL + SHIFT + D** 以切换到调试视图。</span><span class="sxs-lookup"><span data-stu-id="981bd-119">Next, choose  **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

5. <span data-ttu-id="981bd-120">从调试选项中，为你的主机应用程序选择 Microsoft Edge Chromium 选项，例如 **Excel 桌面版（Microsoft Edge Chromium）**。</span><span class="sxs-lookup"><span data-stu-id="981bd-120">From the Debug options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**.</span></span> <span data-ttu-id="981bd-121">选择 **F5** 或从菜单选择 **Debug > Start Debugging** 以开始调试。</span><span class="sxs-lookup"><span data-stu-id="981bd-121">Select **F5** or choose **Debug > Start Debugging** from the menu to begin debugging.</span></span>

6. <span data-ttu-id="981bd-122">在主机应用程序（如 Excel）中，你的加载项现在可以使用了。</span><span class="sxs-lookup"><span data-stu-id="981bd-122">In the host application, such as Excel, your add-in is now ready to use.</span></span> <span data-ttu-id="981bd-123">选择 **显示任务窗格** 或运行其他加载项命令。</span><span class="sxs-lookup"><span data-stu-id="981bd-123">Select **Show Taskpane** or run any other add-in command.</span></span> <span data-ttu-id="981bd-124">此时将出现一个对话框，内容是：</span><span class="sxs-lookup"><span data-stu-id="981bd-124">A dialog box will appear, reading:</span></span>

> <span data-ttu-id="981bd-125">WebView 停止加载。</span><span class="sxs-lookup"><span data-stu-id="981bd-125">WebView Stop On Load.</span></span> 
> <span data-ttu-id="981bd-126">要调试 webview，请使用适用于 Microsoft Edge 扩展的 Microsoft 调试器将 VS 代码附加到 webview 实例，然后单击“确定”以继续。</span><span class="sxs-lookup"><span data-stu-id="981bd-126">To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue.</span></span> <span data-ttu-id="981bd-127">要防止今后出现此对话框，单击“取消”。</span><span class="sxs-lookup"><span data-stu-id="981bd-127">To prevent this dialog from appearing in the future, click Cancel."</span></span>

<span data-ttu-id="981bd-128">选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="981bd-128">Select **OK**.</span></span>

> [!NOTE]
> <span data-ttu-id="981bd-129">如果选择“**取消**”，则当加载项的此实例正在运行时，将不会再次显示该对话框。</span><span class="sxs-lookup"><span data-stu-id="981bd-129">If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running.</span></span> <span data-ttu-id="981bd-130">但如果重新启动加载项，则会再次看到该对话框。</span><span class="sxs-lookup"><span data-stu-id="981bd-130">However, if you restart your add-in, you'll see the dialog again.</span></span>

7. <span data-ttu-id="981bd-131">现在可以在你的项目代码中设置断点并进行调试。</span><span class="sxs-lookup"><span data-stu-id="981bd-131">You're now able to set breakpoints in your project's code and debug.</span></span>

## <a name="see-also"></a><span data-ttu-id="981bd-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="981bd-132">See also</span></span>

* [<span data-ttu-id="981bd-133">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="981bd-133">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="981bd-134">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="981bd-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
* [<span data-ttu-id="981bd-135">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="981bd-135">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)