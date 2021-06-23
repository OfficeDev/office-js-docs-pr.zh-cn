---
title: '调试基于事件Outlook加载项 (预览) '
description: 了解如何调试Outlook基于事件的激活的加载项。
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 8cabbb669d9b46e047efa7e79ae4225c1fc22689
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077090"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="80fd1-103">调试基于事件Outlook加载项 (预览) </span><span class="sxs-lookup"><span data-stu-id="80fd1-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="80fd1-104">本文提供了在外接程序中实现基于 [事件的](autolaunch.md) 激活时调试指南。</span><span class="sxs-lookup"><span data-stu-id="80fd1-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="80fd1-105">基于事件的激活功能当前处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="80fd1-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="80fd1-106">此调试功能仅在使用 Outlook 订阅Windows预览Microsoft 365支持。</span><span class="sxs-lookup"><span data-stu-id="80fd1-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="80fd1-107">有关详细信息，请参阅本文中的预览调试基于 [事件的](#preview-debugging-for-the-event-based-activation-feature) 激活功能部分。</span><span class="sxs-lookup"><span data-stu-id="80fd1-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="80fd1-108">本文将讨论启用调试的关键阶段。</span><span class="sxs-lookup"><span data-stu-id="80fd1-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="80fd1-109">标记加载项进行调试</span><span class="sxs-lookup"><span data-stu-id="80fd1-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="80fd1-110">配置Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="80fd1-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="80fd1-111">附加Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="80fd1-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="80fd1-112">Debug</span><span class="sxs-lookup"><span data-stu-id="80fd1-112">Debug</span></span>](#debug)

<span data-ttu-id="80fd1-113">有几种创建加载项项目的选项。</span><span class="sxs-lookup"><span data-stu-id="80fd1-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="80fd1-114">根据你使用的选项，步骤可能会有所不同。</span><span class="sxs-lookup"><span data-stu-id="80fd1-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="80fd1-115">在这种情况下，如果使用 Office 加载项的 Yeoman 生成器创建加载项项目 (，例如，通过执行基于事件的激活演练 [) ，](autolaunch.md)请按照 **yo office** 步骤操作，否则执行其他步骤。 </span><span class="sxs-lookup"><span data-stu-id="80fd1-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="80fd1-116">Visual Studio Code版本 1.56.1。</span><span class="sxs-lookup"><span data-stu-id="80fd1-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="80fd1-117">预览基于事件的激活功能调试</span><span class="sxs-lookup"><span data-stu-id="80fd1-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="80fd1-118">我们邀请你试用基于事件的激活功能调试功能！</span><span class="sxs-lookup"><span data-stu-id="80fd1-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="80fd1-119">请告诉我们你的方案，以及我们如何通过反馈提供反馈GitHub (请参阅此页面末尾的反馈部分) 。 </span><span class="sxs-lookup"><span data-stu-id="80fd1-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="80fd1-120">若要预览此功能，Outlook上Windows，最低要求版本为 16.0.13729.20000。</span><span class="sxs-lookup"><span data-stu-id="80fd1-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="80fd1-121">若要访问 Office beta 版本，请加入[Office 预览体验计划](https://insider.office.com)。</span><span class="sxs-lookup"><span data-stu-id="80fd1-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="80fd1-122">标记加载项进行调试</span><span class="sxs-lookup"><span data-stu-id="80fd1-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="80fd1-123">设置注册表项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="80fd1-124">`[Add-in ID]`是加载项清单中的 **ID。**</span><span class="sxs-lookup"><span data-stu-id="80fd1-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="80fd1-125">**yo office：** 在命令行窗口中，导航到加载项文件夹的根目录，然后运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="80fd1-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="80fd1-126">除了生成代码和启动本地服务器之外，此命令还应将此加载项的注册表项 `UseDirectDebugger` 设置为 `1` 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="80fd1-127">**其他**：在 `UseDirectDebugger` 下添加注册表项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="80fd1-128">将 `[Add-in ID]` 替换为外接程序清单中的 **Id。**</span><span class="sxs-lookup"><span data-stu-id="80fd1-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="80fd1-129">将注册表项设置为 `1` 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="80fd1-130">如果Outlook桌面 (桌面Outlook，请启动桌面) 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="80fd1-131">撰写新邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="80fd1-131">Compose a new message or appointment.</span></span> <span data-ttu-id="80fd1-132">应看到以下对话框。</span><span class="sxs-lookup"><span data-stu-id="80fd1-132">You should see the following dialog.</span></span> <span data-ttu-id="80fd1-133">*不要* 与对话框进行交互。</span><span class="sxs-lookup"><span data-stu-id="80fd1-133">Do *not* interact with the dialog yet.</span></span>

    ![调试基于事件的处理程序对话框的屏幕截图。](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="80fd1-135">配置Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="80fd1-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="80fd1-136">yo office</span><span class="sxs-lookup"><span data-stu-id="80fd1-136">yo office</span></span>

1. <span data-ttu-id="80fd1-137">返回到命令行窗口，打开Visual Studio Code。</span><span class="sxs-lookup"><span data-stu-id="80fd1-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="80fd1-138">In Visual Studio Code， open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span><span class="sxs-lookup"><span data-stu-id="80fd1-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="80fd1-139">保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="80fd1-139">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a><span data-ttu-id="80fd1-140">其他</span><span class="sxs-lookup"><span data-stu-id="80fd1-140">Other</span></span>

1. <span data-ttu-id="80fd1-141">在桌面 **文件夹中创建一** (调试文件夹) 。 </span><span class="sxs-lookup"><span data-stu-id="80fd1-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="80fd1-142">打开 Visual Studio 代码。</span><span class="sxs-lookup"><span data-stu-id="80fd1-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="80fd1-143">转到"**文件**  >  **""打开** 文件夹"，导航到刚创建的文件夹，然后选择"**选择文件夹"。**</span><span class="sxs-lookup"><span data-stu-id="80fd1-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="80fd1-144">在活动栏上，选择"调试" (Ctrl+Shift+D) 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![活动栏上的"调试"图标的屏幕截图。](../images/vs-code-debug.png)

1. <span data-ttu-id="80fd1-146">选择" **创建launch.js文件"** 链接。</span><span class="sxs-lookup"><span data-stu-id="80fd1-146">Select the **create a launch.json file** link.</span></span>

    ![Screenshot of link to create a launch.json file in Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="80fd1-148">在" **选择环境** "下拉列表中，选择" **边缘： 启动** "以创建launch.js文件。</span><span class="sxs-lookup"><span data-stu-id="80fd1-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="80fd1-149">将以下摘录添加到配置列表中。</span><span class="sxs-lookup"><span data-stu-id="80fd1-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="80fd1-150">保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="80fd1-150">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a><span data-ttu-id="80fd1-151">附加Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="80fd1-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="80fd1-152">若要查找外接程序的 **bundle.js，** 在 Windows 资源管理器中打开以下文件夹，并搜索在清单 (找到的) 。 </span><span class="sxs-lookup"><span data-stu-id="80fd1-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="80fd1-153">打开以此 ID 作为前缀的文件夹并复制其完整路径。</span><span class="sxs-lookup"><span data-stu-id="80fd1-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="80fd1-154">In Visual Studio Code， open **bundle.js** from that folder.</span><span class="sxs-lookup"><span data-stu-id="80fd1-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="80fd1-155">文件路径的模式应如下所示：</span><span class="sxs-lookup"><span data-stu-id="80fd1-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="80fd1-156">将断点bundle.js调试器停止的位置。</span><span class="sxs-lookup"><span data-stu-id="80fd1-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="80fd1-157">在 **"调试**"下拉列表中，选择名称 **"Direct Debugging"，** 然后选择"运行 **"。**</span><span class="sxs-lookup"><span data-stu-id="80fd1-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![Screenshot of selecting Direct Debugging from configuration options in the Visual Studio Code Debug dropdown.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="80fd1-159">Debug</span><span class="sxs-lookup"><span data-stu-id="80fd1-159">Debug</span></span>

1. <span data-ttu-id="80fd1-160">确认已附加调试程序后，返回到Outlook，在"调试基于 **事件的处理程序**"对话框中，选择"确定 **"。**</span><span class="sxs-lookup"><span data-stu-id="80fd1-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="80fd1-161">现在，你可以点击 Visual Studio Code 断点，从而可以调试基于事件的激活代码。</span><span class="sxs-lookup"><span data-stu-id="80fd1-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="80fd1-162">停止调试</span><span class="sxs-lookup"><span data-stu-id="80fd1-162">Stop debugging</span></span>

<span data-ttu-id="80fd1-163">若要停止调试当前桌面会话Outlook，在"调试基于 **事件的处理程序**"对话框中，选择"取消 **"。**</span><span class="sxs-lookup"><span data-stu-id="80fd1-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="80fd1-164">若要重新启用调试，请重新启动Outlook桌面。</span><span class="sxs-lookup"><span data-stu-id="80fd1-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="80fd1-165">若要阻止 **"调试** 基于事件的处理程序"对话框弹出并停止后续 Outlook 会话的调试，请删除关联的注册表项或将其值设置为 `0` ： `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` 。</span><span class="sxs-lookup"><span data-stu-id="80fd1-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="80fd1-166">另请参阅</span><span class="sxs-lookup"><span data-stu-id="80fd1-166">See also</span></span>

- [<span data-ttu-id="80fd1-167">配置Outlook加载项进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="80fd1-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="80fd1-168">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="80fd1-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
