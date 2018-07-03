---
title: 在 iPad 和 Mac 上调试 Office 外接程序
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: e9efae76aa3341eacfd73d6afcc3a3274536aa9d
ms.sourcegitcommit: 6fbf42723f9c1b72095700c20458fd0e8c572794
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2018
ms.locfileid: "19722329"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="95880-102">在 iPad 和 Mac 上调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="95880-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="95880-p101">您可以使用 Visual Studio 开发和调试 Windows 上的外接程序。但是，无法使用它调试 iPad 或 Mac 上的外接程序。由于外接程序使用 HTML 和 Javascript 开发，它们应旨在跨平台工作，但不同浏览器呈现您的 HTML 的方式可能存在细微差异。本文介绍如何调试在 iPad 或 Mac 上运行的外接程序。</span><span class="sxs-lookup"><span data-stu-id="95880-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span> 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="95880-106">使用 Mac 上的 Safari Web Inspector 进行调试</span><span class="sxs-lookup"><span data-stu-id="95880-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="95880-107">如果您拥有在任务窗格或内容加载项中显示 UI 的加载项，则可以使用 Safari Web Inspector 调试 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="95880-107">If you have add-in that shows UI in a taskpane or in a content add-in, you can debug an Office add-in using Safari Web Inspector.</span></span> 

<span data-ttu-id="95880-108">要能够在 Mac上调试 Office 外接程序，您必须具有 Mac OS 高级 Sierra 和 16.9.1 版本（内部版本 18012504）或更高版本的 Mac Office。</span><span class="sxs-lookup"><span data-stu-id="95880-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="95880-109">如果您没有 Office Mac 版本，可以通过加入 [Office 365 开发人员计划](https://aka.ms/o365devprogram)来获得。</span><span class="sxs-lookup"><span data-stu-id="95880-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="95880-110">要开始，打开一个终端并设置`OfficeWebAddinDeveloperExtras` 相关 Office 应用的属性如下：</span><span class="sxs-lookup"><span data-stu-id="95880-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="95880-111">然后，打开 Office 应用程序并插入您的加载项。</span><span class="sxs-lookup"><span data-stu-id="95880-111">Then, open the Office application and insert your add-in.</span></span> <span data-ttu-id="95880-112">用鼠标右键单击该加载项，您应该在上下文菜单中看到一个**检查元素**的选项。</span><span class="sxs-lookup"><span data-stu-id="95880-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="95880-113">选择该选项，会有检查器弹出，您可以在其中设置断点并调试加载项。</span><span class="sxs-lookup"><span data-stu-id="95880-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="95880-114">请注意，这是一项实验性功能，我们无法保证未来版本的 Office 应用程序中一定保留此功能。</span><span class="sxs-lookup"><span data-stu-id="95880-114">Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a><span data-ttu-id="95880-115">在 iPad 或 Mac 上使用 Vorlon.JS 进行调试</span><span class="sxs-lookup"><span data-stu-id="95880-115">Debugging with Vorlon.JS on a iPad or Mac</span></span>

<span data-ttu-id="95880-116">要在 iPad 或 Mac 上调试加载项，可以使用 Vorlon.JS，这是一个类似于 F12 工具的网页调试器。</span><span class="sxs-lookup"><span data-stu-id="95880-116">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="95880-117">它旨在实现远程工作，使您能够在不同设备上调试网页。</span><span class="sxs-lookup"><span data-stu-id="95880-117">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="95880-118">有关详细信息，请参阅 [Vorlon 网站](http://www.vorlonjs.com)。</span><span class="sxs-lookup"><span data-stu-id="95880-118">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="95880-119">安装并设置 Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="95880-119">Install and set up up Vorlon.JS on a Mac or iPad</span></span>  

1.  <span data-ttu-id="95880-120">以管理员身份登录到设备。</span><span class="sxs-lookup"><span data-stu-id="95880-120">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="95880-121">如果尚未安装 [Node.js](https://nodejs.org)，请执行安装。</span><span class="sxs-lookup"><span data-stu-id="95880-121">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span> 

3.  <span data-ttu-id="95880-p105">打开“**终端**”窗口，然后输入命令 `npm i -g vorlon`。该工具将安装到 `/usr/local/lib/node_modules/vorlon`。</span><span class="sxs-lookup"><span data-stu-id="95880-p105">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="95880-124">将 Vorlon.JS 配置为使用 HTTPS</span><span class="sxs-lookup"><span data-stu-id="95880-124">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="95880-p106">若要使用 Vorlon.JS 调试应用，请将 `<script>` 标记添加到应用的开始页，以便从已知位置加载 Vorlon.JS 脚本（有关详细信息，请参阅以下过程）。如果加载项受 SSL 保护 (HTTPS)，它使用的任何脚本都必须通过 HTTPS 服务器进行托管，包括 Vorlon.JS 脚本。因此，必须将 Vorlon.JS 配置为使用 SSL，这样才能结合使用 Vorlon.JS 和加载项。</span><span class="sxs-lookup"><span data-stu-id="95880-p106">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span> 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="95880-128">在**查找器**中，转到 `/usr/local/lib/node_modules/vorlon`，打开 `/Server` 文件夹的上下文菜单（右键单击），再选择“获取信息”****。</span><span class="sxs-lookup"><span data-stu-id="95880-128">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="95880-129">在“**服务器信息**”窗口的右下角选择挂锁图标来解锁该文件夹。</span><span class="sxs-lookup"><span data-stu-id="95880-129">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="95880-130">在窗口的“**共享和权限**”部分，将“**员工**”组的“**特权**”设置为“**读写**”。</span><span class="sxs-lookup"><span data-stu-id="95880-130">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="95880-131">再次选择挂锁图标以***重新锁定***文件夹。</span><span class="sxs-lookup"><span data-stu-id="95880-131">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="95880-132">返回**查找器**，展开 `/Server` 子文件夹，右键单击文件 `config.json`，然后选择“**获取信息**”。</span><span class="sxs-lookup"><span data-stu-id="95880-132">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="95880-p107">在“**config.json 信息**”窗口中，完全按照更改 `/Server` 父文件夹的方式来更改文件特权。请务必重新锁定并关闭窗口。</span><span class="sxs-lookup"><span data-stu-id="95880-p107">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="95880-p108">返回**查找器**，右键单击文件 `config.json`，选择“**打开方式**”，然后选择“**文本编辑**”。在文本编辑器中打开该文件。</span><span class="sxs-lookup"><span data-stu-id="95880-p108">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="95880-137">将 **useSSL** 属性的值更改为 `true`。</span><span class="sxs-lookup"><span data-stu-id="95880-137">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="95880-p109">在“**插件**”部分，使用 `OFFICE` 的 **id** 和 `Office Addin` 的**名称**查找插件。如果插件的“**启用**”属性还不是 `true`，请将其设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="95880-p109">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="95880-140">保存文件并关闭编辑器。</span><span class="sxs-lookup"><span data-stu-id="95880-140">Save the file and close the editor.</span></span>

11. <span data-ttu-id="95880-141">在**查找器**中，导航到 `/usr/local/lib/node_modules/vorlon`，右键单击 `Server` 子文件夹，然后选择“**文件夹的新终端**”。</span><span class="sxs-lookup"><span data-stu-id="95880-141">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
12. <span data-ttu-id="95880-p110">在“**终端**”窗口中，输入 `sudo vorlon`。系统将提示你输入管理员密码。Vorlon 服务器将启动。使“**终端**”窗口保持打开状态。</span><span class="sxs-lookup"><span data-stu-id="95880-p110">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="95880-p111">打开浏览器窗口，再转到 Vorlon.JS 界面 `https://localhost:1337`。当出现提示时，选择“始终”****，以信任安全证书。</span><span class="sxs-lookup"><span data-stu-id="95880-p111">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="95880-p112">如果没有看到提示，可能需要手动信任安全证书。证书文件是 `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`。请尝试执行以下步骤。如有疑问，请咨询 Macintosh 或 iPad 帮助人员。</span><span class="sxs-lookup"><span data-stu-id="95880-p112">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span> 
    >
    > 1. <span data-ttu-id="95880-152">关闭浏览器窗口，在运行 Vorlon 服务器的“终端”**** 窗口中，按 Control-C 停止服务器。</span><span class="sxs-lookup"><span data-stu-id="95880-152">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="95880-p113">在**查找器**中，右键单击 `server.crt` 文件并选择“**钥匙链访问**”。“**钥匙链访问**”窗口将打开。</span><span class="sxs-lookup"><span data-stu-id="95880-p113">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="95880-p114">在左侧的“**钥匙链**”列表中，如果尚未选择“**登录**”，请进行选择，然后再选择“**类别**”部分中的“**证书**”。将列出证书 **localhost**。</span><span class="sxs-lookup"><span data-stu-id="95880-p114">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="95880-p115">右键单击证书 **localhost**，并选择“**获取信息**”。**localhost** 窗口将打开。</span><span class="sxs-lookup"><span data-stu-id="95880-p115">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="95880-159">在“**信任**”部分，打开标记了“**使用此证书时**”的选择器，并选择“**始终相信**”。</span><span class="sxs-lookup"><span data-stu-id="95880-159">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="95880-p116">关闭 **localhost** 窗口。如果此操作成功，“**钥匙链访问**”窗口中的“**localhost**”证书图标将显示蓝色圆圈中带白色十字图案。</span><span class="sxs-lookup"><span data-stu-id="95880-p116">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="95880-162">配置外接程序用于 Vorlon.JS 调试</span><span class="sxs-lookup"><span data-stu-id="95880-162">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="95880-163">向外接程序的 home.html 文件（或主 HTML 文件）的 `<head>` 部分添加以下脚本标记：</span><span class="sxs-lookup"><span data-stu-id="95880-163">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. <span data-ttu-id="95880-164">将外接程序 Web 应用程序部署到可从 Mac 或 iPad 进行访问的 Web 服务器，如 Azure 网站。</span><span class="sxs-lookup"><span data-stu-id="95880-164">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span> 

3. <span data-ttu-id="95880-165">更新所有位置的外接程序 URL，其中 URL 出现在外接程序清单中。</span><span class="sxs-lookup"><span data-stu-id="95880-165">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="95880-166">将外接程序清单复制到 Mac 或 iPad 上的以下文件夹：`/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`，其中 *{host_name}* 为 Word、Excel、PowerPoint 或 Outlook。</span><span class="sxs-lookup"><span data-stu-id="95880-166">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="95880-167">检查 Vorlon.JS 中的外接程序</span><span class="sxs-lookup"><span data-stu-id="95880-167">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="95880-168">如果 Vorlon 服务器未运行，则在**查找器**中，导航到 `/usr/local/lib/node_modules/vorlon`，右键单击 `Server` 子文件夹，然后选择“**文件夹的新终端**”。</span><span class="sxs-lookup"><span data-stu-id="95880-168">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
2.  <span data-ttu-id="95880-p117">在“**终端**”窗口中，输入 `sudo vorlon`。系统将提示你输入管理员密码。Vorlon 服务器将启动。使“**终端**”窗口保持打开状态。</span><span class="sxs-lookup"><span data-stu-id="95880-p117">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="95880-173">打开浏览器窗口，然后转到 Vorlon.JS 界面 `https://localhost:1337`。</span><span class="sxs-lookup"><span data-stu-id="95880-173">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="95880-p118">旁加载外接程序。如果是针对 Excel、PowerPoint 或 Word，请按[在 iPad 和 Mac 上旁加载 Office 外接程序](sideload-an-office-add-in-on-ipad-and-mac.md)中所述进行旁加载。如果是 Outlook 外接程序，请按[旁加载 Outlook 外接程序用于测试](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing)进行旁加载。如果外接程序不使用外接程序命令，它将立即打开。否则，请选择按钮以打开外接程序。按钮位于“**主页**”选项卡或“**外接程序**”选项卡上，具体取决于 Office 主机应用程序版本。</span><span class="sxs-lookup"><span data-stu-id="95880-p118">Sideload the add-in. If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md). If it is an Outlook add-in, sideload it as described in [Sideload Outlook Add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). If the add-in does not use add-in commands, it will open immediately. Otherwise, choose the button to open the add-in. Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="95880-180">外接程序将在 Vorlon.JS（在 Vorlon.JS 界面左侧）的客户端列表中显示为 **{OS} - n**，*n* 代表数字，而 *{OS}* 表示设备类型，例如“Macintosh”。</span><span class="sxs-lookup"><span data-stu-id="95880-180">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span> 

![显示 Vorlon.js 界面的快照](../images/vorlon-interface.png)

<span data-ttu-id="95880-p119">Vorlon 工具具有多种插件。当前已启用的插件显示为工具顶部的选项卡。（可以通过选择左侧的齿轮图标启用更多插件。）这些插件类似于 F12 工具中的功能。例如，可以突出显示 DOM 元素，执行命令等。有关详细信息，请参阅 [Vorlon 文档核心插件](http://vorlonjs.com/documentation/#console)</span><span class="sxs-lookup"><span data-stu-id="95880-p119">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool. (You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools. For example, you can highlight DOM elements, execute commands, and more. For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console)</span></span> 

<span data-ttu-id="95880-p120">**Office 外接程序**插件为 Office.js 添加额外的功能，例如探索对象模型、执行 Office.js 调用和读取对象属性的值。有关说明，请参阅[调试 Office 外接程序的 VorlonJS 插件](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)。</span><span class="sxs-lookup"><span data-stu-id="95880-p120">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="95880-188">无法在 Vorlon.JS 中设置断点。</span><span class="sxs-lookup"><span data-stu-id="95880-188">There is no way to set break points in Vorlon.JS.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="95880-189">在 Mac 或 iPad 上清除 Office 应用程序缓存</span><span class="sxs-lookup"><span data-stu-id="95880-189">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="95880-p121">出于性能方面的考虑，外接程序通常在 Office for Mac 中缓存。通常情况下，将通过重载外接程序清除缓存。如果同一文档中存在多个外接程序，则重载后自动清除缓存的过程可能不可靠。</span><span class="sxs-lookup"><span data-stu-id="95880-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span> 

<span data-ttu-id="95880-193">在 Mac 上，通过删除 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的所有内容可以手动清除缓存。</span><span class="sxs-lookup"><span data-stu-id="95880-193">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

<span data-ttu-id="95880-p122">在 iPad 上，可以从外接程序中的 JavaScript 调用 `window.location.reload(true)` 来强制重载。或者，可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="95880-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
