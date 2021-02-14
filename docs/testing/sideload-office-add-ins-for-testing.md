---
title: 在 Office 网页版中旁加载 Office 加载项进行测试
description: 通过旁加载在 Office 网页中测试 Office 加载项。
ms.date: 02/11/2021
localization_priority: Normal
ms.openlocfilehash: f81fbc163135be5a616e7b44e604cb842da9870b
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238062"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="0edb7-103">在 Office 网页版中旁加载 Office 加载项进行测试</span><span class="sxs-lookup"><span data-stu-id="0edb7-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="0edb7-104">旁加载外接程序时，无需先将其放入加载项目录中即可安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="0edb7-104">When you sideload an add-in, you're able to install the add-in without first putting it in the add-in catalog.</span></span> <span data-ttu-id="0edb7-105">在测试和开发外接程序时，这非常有用，因为你可以看到外接程序的显示方式和运行方式。</span><span class="sxs-lookup"><span data-stu-id="0edb7-105">This is useful when testing and developing your add-in because you can see how your add-in will appear and function.</span></span>

<span data-ttu-id="0edb7-106">旁加载外接程序时，加载项清单存储在浏览器的本地存储中，因此，如果清除浏览器的缓存或切换到其他浏览器，必须再次旁加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="0edb7-106">When you sideload an add-in, the add-in's manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

<span data-ttu-id="0edb7-107">旁加载因主机应用程序 (例如 Excel) 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-107">Sideloading varies between host applications (for example, Excel).</span></span>

> [!NOTE]
> <span data-ttu-id="0edb7-p102">Excel、OneNote、PowerPoint 和 Word 支持本文中所述的旁加载。若要旁加载 Outlook 外接程序，请参阅 [旁加载 Outlook](../outlook/sideload-outlook-add-ins-for-testing.md)外接程序进行测试。</span><span class="sxs-lookup"><span data-stu-id="0edb7-p102">Sideloading as described in this article is supported on Excel, OneNote, PowerPoint, and Word. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="0edb7-110">在 Office 网页版中旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0edb7-110">Sideload an Office Add-in in Office on the web</span></span>

<span data-ttu-id="0edb7-111">仅 **Excel、OneNote、PowerPoint** 和 Word 支持 **此过程**。</span><span class="sxs-lookup"><span data-stu-id="0edb7-111">This process is supported for **Excel**, **OneNote**, **PowerPoint**, and **Word** only.</span></span> <span data-ttu-id="0edb7-112">有关其他主机应用程序，请参阅以下部分中的手动旁加载说明。</span><span class="sxs-lookup"><span data-stu-id="0edb7-112">For other host applications, see the manual sideloading instructions in the following section.</span></span> <span data-ttu-id="0edb7-113">此示例项目假定您使用的是使用适用于 Office 加载项的 [Yeoman 生成器创建的项目](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="0edb7-113">This example project assumes that you are using a project created with [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="0edb7-114">在[Web 上打开 Office。](https://office.live.com/)</span><span class="sxs-lookup"><span data-stu-id="0edb7-114">Open [Office on the web](https://office.live.com/).</span></span> <span data-ttu-id="0edb7-115">使用"**创建"** 选项，在 **Excel、OneNote、PowerPoint** 或 **Word** 中创建文档。  </span><span class="sxs-lookup"><span data-stu-id="0edb7-115">Using the **Create** option, make a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**.</span></span> <span data-ttu-id="0edb7-116">在此新文档中，选择功能 **区** 中的"共享"，选择 **"复制链接**"，然后复制 URL。</span><span class="sxs-lookup"><span data-stu-id="0edb7-116">In this new document, select **Share** in the ribbon, select **Copy Link**, and copy the URL.</span></span>

2. <span data-ttu-id="0edb7-117">在 yo office 项目文件的根目录中，打开package.js **文件** 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-117">In the root directory of your yo office project files, open the **package.json** file.</span></span> <span data-ttu-id="0edb7-118">在此 **文件的脚本** 部分中，创建 `"document"` 一个属性。</span><span class="sxs-lookup"><span data-stu-id="0edb7-118">Within the **scripts** section of this file, create a `"document"` property.</span></span> <span data-ttu-id="0edb7-119">粘贴您复制的 URL 作为属性的值 `"document"` 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-119">Paste the URL you copied as the value for the `"document"` property.</span></span> <span data-ttu-id="0edb7-120">例如，你的将如下所示：</span><span class="sxs-lookup"><span data-stu-id="0edb7-120">For example, yours will look something like this:</span></span>

    ```json
      "scripts": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > <span data-ttu-id="0edb7-121">如果您创建的外接程序不是使用 Yeoman 生成器，您可以通过将以下内容追加到现有 URL，将查询参数添加到文档的 URL 中：</span><span class="sxs-lookup"><span data-stu-id="0edb7-121">If you are creating an add-in not using our Yeoman generator, you can add query parameters to your document's URL, by appending the following to the existing URL:</span></span>

    - <span data-ttu-id="0edb7-122">开发服务器端口，如 `&wdaddindevserverport=3000` 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-122">The dev server port, such as `&wdaddindevserverport=3000`.</span></span>
    - <span data-ttu-id="0edb7-123">清单文件名，如 `&wdaddinmanifestfile=manifest1.xml` 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-123">The manifest file name, such as `&wdaddinmanifestfile=manifest1.xml`.</span></span>
    - <span data-ttu-id="0edb7-124">清单 GUID，如 `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-124">The manifest GUID, such as `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143`.</span></span>

    > <span data-ttu-id="0edb7-125">如果使用 Yeoman 生成器，则不需要添加此信息，因为 Yeoman 工具会自动追加此信息。</span><span class="sxs-lookup"><span data-stu-id="0edb7-125">If you are using the Yeoman generator, adding this information is not necessary as the Yeoman tooling appends this information automatically.</span></span>
    > <span data-ttu-id="0edb7-126">请注意，在这两种情况下，只能从 localhost 加载清单。</span><span class="sxs-lookup"><span data-stu-id="0edb7-126">Note that in both cases, however, you can only load manifests from localhost.</span></span>

3. <span data-ttu-id="0edb7-127">在从项目的根目录开始的命令行中，运行以下 `npm run start:web` 命令：</span><span class="sxs-lookup"><span data-stu-id="0edb7-127">In the command line starting at the root directory of your project, run the following command: `npm run start:web`.</span></span>

4. <span data-ttu-id="0edb7-128">首次使用此方法在 Web 上旁加载外接程序时，你将看到一个对话框，要求您启用开发人员模式。</span><span class="sxs-lookup"><span data-stu-id="0edb7-128">The first time you use this method to sideload an add-in on the web, you'll see a dialog asking you to enable developer mode.</span></span> <span data-ttu-id="0edb7-129">选中"现在启用 **开发人员模式"复选框，** 然后选择"**确定"。**</span><span class="sxs-lookup"><span data-stu-id="0edb7-129">Select the checkbox for **Enable Developer Mode now** and select **OK**.</span></span>

5. <span data-ttu-id="0edb7-130">你将看到第二个对话框，询问您是否希望从计算机注册 Office 外接程序清单。</span><span class="sxs-lookup"><span data-stu-id="0edb7-130">You will see a second dialog box, asking if you wish to register an Office Add-in manifest from your computer.</span></span> <span data-ttu-id="0edb7-131">应选择"**是"。**</span><span class="sxs-lookup"><span data-stu-id="0edb7-131">You should select **Yes**.</span></span>

6. <span data-ttu-id="0edb7-132">已安装加载项。</span><span class="sxs-lookup"><span data-stu-id="0edb7-132">Your add-in is installed.</span></span> <span data-ttu-id="0edb7-133">如果是加载项命令，它应显示在功能区或上下文菜单上。</span><span class="sxs-lookup"><span data-stu-id="0edb7-133">If it is an add-in command, it should appear on either the ribbon or the context menu.</span></span> <span data-ttu-id="0edb7-134">如果是任务窗格加载项，应显示任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0edb7-134">If it is a task pane add-in, the task pane should appear.</span></span>

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a><span data-ttu-id="0edb7-135">手动在 Office 网页中旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0edb7-135">Sideload an Office Add-in in Office on the web manually</span></span>

<span data-ttu-id="0edb7-136">此方法不使用命令行，并且只能在主机应用程序（如 Excel (）内使用) 。</span><span class="sxs-lookup"><span data-stu-id="0edb7-136">This method doesn't use the command line and can be accomplished using commands only within the host application (such as Excel).</span></span>

1. <span data-ttu-id="0edb7-137">在[Web 上打开 Office。](https://office.live.com/)</span><span class="sxs-lookup"><span data-stu-id="0edb7-137">Open [Office on the web](https://office.live.com/).</span></span> <span data-ttu-id="0edb7-138">在 **Excel、Word** 或 **PowerPoint** 中打开文档。</span><span class="sxs-lookup"><span data-stu-id="0edb7-138">Open a document in **Excel**, **Word**, or **PowerPoint**.</span></span> <span data-ttu-id="0edb7-139">在"**加载项**"部分功能区上的"插入"选项卡上，选择 **"Office 加载项"。**</span><span class="sxs-lookup"><span data-stu-id="0edb7-139">On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Office Add-ins**.</span></span>

1. <span data-ttu-id="0edb7-140">在 **"Office 加载项"对话框中**，选择 **"我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，然后上载 **"我的外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="0edb7-140">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>

    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

1. <span data-ttu-id="0edb7-142">**转到** 加载项清单文件，再选择“上传”。</span><span class="sxs-lookup"><span data-stu-id="0edb7-142">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

1. <span data-ttu-id="0edb7-p111">验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。</span><span class="sxs-lookup"><span data-stu-id="0edb7-p111">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
> <span data-ttu-id="0edb7-147">若要使用原始 WebView (EdgeHTML) Microsoft Edge 测试 Office 外接程序，需要执行其他配置步骤。</span><span class="sxs-lookup"><span data-stu-id="0edb7-147">To test your Office Add-in with Microsoft Edge with the original WebView (EdgeHTML), an additional configuration step is required.</span></span> <span data-ttu-id="0edb7-148">在 Windows 命令提示符中，运行以下 `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` 行：</span><span class="sxs-lookup"><span data-stu-id="0edb7-148">In a Windows Command Prompt, run the following line: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`.</span></span> <span data-ttu-id="0edb7-149">当 Office 使用基于 Chromium 的边缘 WebView2 时，不需要这样做。</span><span class="sxs-lookup"><span data-stu-id="0edb7-149">This is not required when Office is using the Chromium-based Edge WebView2.</span></span> <span data-ttu-id="0edb7-150">有关详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="0edb7-150">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

## <a name="sideload-an-office-add-in"></a><span data-ttu-id="0edb7-151">旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0edb7-151">Sideload an Office Add-in</span></span>

1. <span data-ttu-id="0edb7-152">登录到 Microsoft 365 帐户。</span><span class="sxs-lookup"><span data-stu-id="0edb7-152">Sign in to your Microsoft 365 account.</span></span>

2. <span data-ttu-id="0edb7-153">打开工具栏左端的应用程序启动器，选择 **Excel、Word** 或 **PowerPoint，** 然后新建文档。</span><span class="sxs-lookup"><span data-stu-id="0edb7-153">Open the App Launcher on the left end of the toolbar and select **Excel**, **Word**, or **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="0edb7-154">步骤 3 - 6 与上一部分 **在 Office 网页版中旁加载 Office 加载项** 相同。</span><span class="sxs-lookup"><span data-stu-id="0edb7-154">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="0edb7-155">使用 Visual Studio 时旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="0edb7-155">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="0edb7-156">如果你使用 Visual Studio开发外接程序，旁加载过程类似于手动旁加载到 Web。</span><span class="sxs-lookup"><span data-stu-id="0edb7-156">If you're using Visual Studio to develop your add-in, the process to sideload is similar to manual sideloading to the web.</span></span> <span data-ttu-id="0edb7-157">唯一的区别是，必须更新清单中 **SourceURL** 元素的值以包含部署加载项位置的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="0edb7-157">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="0edb7-158">虽然可以将加载项从 Visual Studio 旁加载到 Office 网页版，但无法从 Visual Studio 调试它们。</span><span class="sxs-lookup"><span data-stu-id="0edb7-158">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="0edb7-159">若要进行调试，需要使用浏览器调试工具。</span><span class="sxs-lookup"><span data-stu-id="0edb7-159">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="0edb7-160">有关详细信息，请参阅[在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)。</span><span class="sxs-lookup"><span data-stu-id="0edb7-160">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="0edb7-161">在 Visual Studio 中，通过选择 **视图** > **属性窗口** 来显示 **属性** 窗口。</span><span class="sxs-lookup"><span data-stu-id="0edb7-161">In Visual Studio, show the **Properties** window by choosing **View** > **Properties Window**.</span></span>
2. <span data-ttu-id="0edb7-162">在 **解决方案资源管理器** 中，选择 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="0edb7-162">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="0edb7-163">这将在 **属性** 窗口中显示项目的属性。</span><span class="sxs-lookup"><span data-stu-id="0edb7-163">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="0edb7-164">在“属性”窗口中复制 **SSL URL**。</span><span class="sxs-lookup"><span data-stu-id="0edb7-164">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="0edb7-165">在加载项项目中，打开清单 XML 文件。</span><span class="sxs-lookup"><span data-stu-id="0edb7-165">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="0edb7-166">请确保正在编辑源 XML。</span><span class="sxs-lookup"><span data-stu-id="0edb7-166">Be sure you are editing the source XML.</span></span> <span data-ttu-id="0edb7-167">对于某些项目类型，Visual Studio 将打开 XML 的可视视图，它不适用于下一步骤。</span><span class="sxs-lookup"><span data-stu-id="0edb7-167">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="0edb7-168">使用刚复制的 SSL URL 来搜索和替换 **~remoteAppUrl/** 的所有实例。</span><span class="sxs-lookup"><span data-stu-id="0edb7-168">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="0edb7-169">将看到多个替换，具体取决于项目类型。将显示新 URL，类似于 `https://localhost:44300/Home.html`。</span><span class="sxs-lookup"><span data-stu-id="0edb7-169">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="0edb7-170">保存 XML 文件。</span><span class="sxs-lookup"><span data-stu-id="0edb7-170">Save the XML file.</span></span>
7. <span data-ttu-id="0edb7-171">右键单击 Web 项目，然后选择 **调试** > **启动新实例**。</span><span class="sxs-lookup"><span data-stu-id="0edb7-171">Right click the web project and choose **Debug** > **Start new instance**.</span></span> <span data-ttu-id="0edb7-172">这将在不启动 Office 的情况下运行 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="0edb7-172">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="0edb7-173">从 Office 网页版，使用之前[在 Office 网页版中加载 Office 加载项](#sideload-an-office-add-in-in-office-on-the-web)中所述的步骤旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="0edb7-173">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="0edb7-174">删除旁加载的加载项</span><span class="sxs-lookup"><span data-stu-id="0edb7-174">Remove a sideloaded add-in</span></span>

<span data-ttu-id="0edb7-175">可以通过清除浏览器的缓存来删除之前旁加载的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0edb7-175">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="0edb7-176">如果您更改加载项清单 (例如，更新图标或外接程序命令) 文本的文件名，您可能需要清除 [Office](clear-cache.md) 缓存，然后使用更新后的清单重新旁加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="0edb7-176">If you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to [clear the Office cache](clear-cache.md) and then re-sideload the add-in using the updated manifest.</span></span> <span data-ttu-id="0edb7-177">执行此操作后，Office 将按照更新清单中所述的方式呈现该加载项。</span><span class="sxs-lookup"><span data-stu-id="0edb7-177">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="0edb7-178">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0edb7-178">See also</span></span>

- [<span data-ttu-id="0edb7-179">在 iPad 和 Mac 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0edb7-179">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="0edb7-180">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="0edb7-180">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)
- [<span data-ttu-id="0edb7-181">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="0edb7-181">Clear the Office cache</span></span>](clear-cache.md)
