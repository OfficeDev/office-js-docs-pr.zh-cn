---
title: 在 iPad 和 Mac 上调试 Office 外接程序
description: ''
ms.date: 03/21/2018
localization_priority: Priority
ms.openlocfilehash: 058f3cb4a4acc77a5c4fcd4559970187842c2c4b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388029"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a>在 iPad 和 Mac 上调试 Office 外接程序

您可以使用 Visual Studio 开发和调试 Windows 上的外接程序。但是，无法使用它调试 iPad 或 Mac 上的外接程序。由于外接程序使用 HTML 和 Javascript 开发，它们应旨在跨平台工作，但不同浏览器呈现您的 HTML 的方式可能存在细微差异。本文介绍如何调试在 iPad 或 Mac 上运行的外接程序。 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>在 Mac 上使用 Safari Web 检查器进行调试

如果有在任务窗格或内容加载项中显示 UI 的加载项，可以使用 Safari Web 检查器调试 Office 加载项。

要在 Mac 上调试 Office 加载项，必须拥有 Mac OS High Sierra 和 Mac Office 版本：16.9.1（内部版本 18012504）或更高版本。 如果没有 Office Mac 内部版本，可以通过加入 [Office 365 开发人员计划](https://aka.ms/o365devprogram)获取一个版本。

首先，打开终端，设置相关 Office 应用程序的 `OfficeWebAddinDeveloperExtras` 属性，如下所示：

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

然后，打开 Office 应用程序并插入加载项。 右键单击加载项，应在上下文菜单中看到一个“检查元素”**** 选项。  选择该选项，它将弹出检查器，可以在其中设置断点并调试加载项。

> [!NOTE]
> 请注意，这是一个实验性功能，我们不能保证在未来的 Office 应用程序版本中保留此功能。
>
> 如果试图使用检查器和对话框闪烁，请尝试以下解决方法：
> 1. 缩小对话框大小。
> 2. 选择“检查元素”****，这将在新窗口中打开。
> 3. 将对话框调整为原始大小。
> 4. 根据需要使用检查器。

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a>在 iPad 或 Mac 上使用 Vorlon.JS 进行调试

若要在 iPad 或 Mac 上调试加载项，可以使用 Vorlon.JS，一个类似于 F12 工具的 Web 页调试程序。 它旨在实现远程工作，使你能够在不同设备上调试网页。 有关详细信息，请参阅 [Vorlon 网站](http://www.vorlonjs.com)。  


### <a name="install-and-set-up-vorlonjs"></a>安装和设置 Vorlon.JS  

1.  以管理员身份登录到设备。

2.  如果尚未安装 [Node.js](https://nodejs.org)，请执行安装。

3.  打开“**终端**”窗口，然后输入命令 `npm i -g vorlon`。该工具将安装到 `/usr/local/lib/node_modules/vorlon`。


### <a name="configure-vorlonjs-to-use-https"></a>将 Vorlon.JS 配置为使用 HTTPS

若要使用 Vorlon.JS 调试应用，请将 `<script>` 标记添加到应用的开始页，以便从已知位置加载 Vorlon.JS 脚本（有关详细信息，请参阅以下过程）。如果加载项受 SSL 保护 (HTTPS)，它使用的任何脚本都必须通过 HTTPS 服务器进行托管，包括 Vorlon.JS 脚本。因此，必须将 Vorlon.JS 配置为使用 SSL，这样才能结合使用 Vorlon.JS 和加载项。

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  在**查找器**中，转到 `/usr/local/lib/node_modules/vorlon`，打开 `/Server` 文件夹的上下文菜单（右键单击），再选择“获取信息”****。

2.  在“**服务器信息**”窗口的右下角选择挂锁图标来解锁该文件夹。

3. 在窗口的“**共享和权限**”部分，将“**员工**”组的“**特权**”设置为“**读写**”。

4. 再次选择挂锁图标以***重新锁定***文件夹。

5. 返回**查找器**，展开 `/Server` 子文件夹，右键单击文件 `config.json`，然后选择“**获取信息**”。

6. 在“**config.json 信息**”窗口中，完全按照更改 `/Server` 父文件夹的方式来更改文件特权。请务必重新锁定并关闭窗口。

7. 返回**查找器**，右键单击文件 `config.json`，选择“**打开方式**”，然后选择“**文本编辑**”。在文本编辑器中打开该文件。

8. 将 **useSSL** 属性的值更改为 `true`。

9. 在“**插件**”部分，使用 `OFFICE` 的 **id** 和 `Office Addin` 的**名称**查找插件。如果插件的“**启用**”属性还不是 `true`，请将其设置为 `true`。

10. 保存文件并关闭编辑器。

11. 在**查找器**中，导航到 `/usr/local/lib/node_modules/vorlon`，右键单击 `Server` 子文件夹，然后选择“**文件夹的新终端**”。

12. 在“**终端**”窗口中，输入 `sudo vorlon`。系统将提示你输入管理员密码。Vorlon 服务器将启动。使“**终端**”窗口保持打开状态。

13. 打开浏览器窗口，再转到 Vorlon.JS 界面 `https://localhost:1337`。当出现提示时，选择“始终”****，以信任安全证书。

    > [!NOTE]
    > 如果没有看到提示，可能需要手动信任安全证书。证书文件是 `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`。请尝试执行以下步骤。如有疑问，请咨询 Macintosh 或 iPad 帮助人员。
    >
    > 1. 关闭浏览器窗口，在运行 Vorlon 服务器的“终端”**** 窗口中，按 Control-C 停止服务器。
    > 2. 在**查找器**中，右键单击 `server.crt` 文件并选择“**钥匙链访问**”。“**钥匙链访问**”窗口将打开。
    > 3. 在左侧的“**钥匙链**”列表中，如果尚未选择“**登录**”，请进行选择，然后再选择“**类别**”部分中的“**证书**”。将列出证书 **localhost**。
    > 4. 右键单击证书 **localhost**，并选择“**获取信息**”。**localhost** 窗口将打开。
    > 5. 在“**信任**”部分，打开标记了“**使用此证书时**”的选择器，并选择“**始终相信**”。 
    > 6. 关闭 **localhost** 窗口。如果此操作成功，“**钥匙链访问**”窗口中的“**localhost**”证书图标将显示蓝色圆圈中带白色十字图案。


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>配置外接程序用于 Vorlon.JS 调试

1. 向外接程序的 home.html 文件（或主 HTML 文件）的 `<head>` 部分添加以下脚本标记：

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>
    ```  

2. 将外接程序 Web 应用程序部署到可从 Mac 或 iPad 进行访问的 Web 服务器，如 Azure 网站。

3. 更新所有位置的外接程序 URL，其中 URL 出现在外接程序清单中。

4. 将外接程序清单复制到 Mac 或 iPad 上的以下文件夹：`/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`，其中 *{host_name}* 为 Word、Excel、PowerPoint 或 Outlook。


### <a name="inspect-an-add-in-in-vorlonjs"></a>检查 Vorlon.JS 中的外接程序

1. 如果 Vorlon 服务器未运行，则在**查找器**中，导航到 `/usr/local/lib/node_modules/vorlon`，右键单击 `Server` 子文件夹，然后选择“**文件夹的新终端**”。 

2.  在“**终端**”窗口中，输入 `sudo vorlon`。系统将提示你输入管理员密码。Vorlon 服务器将启动。使“**终端**”窗口保持打开状态。

3.  打开浏览器窗口，并转到 `https://localhost:1337`（即 Vorlon.JS 界面）。

4. 旁加载外接程序。 如果是 Excel、PowerPoint 或 Word 加载项，请按照[在 iPad 和 Mac 上旁加载 Office 加载项](sideload-an-office-add-in-on-ipad-and-mac.md)中的说明操作，执行旁加载。 如果是 Outlook 加载项，请按照[旁加载 Outlook 加载项以供测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)中的说明操作，执行旁加载。 如果加载项不使用加载项命令，它会立即打开。 否则，请选择用于打开加载项的按钮。 按钮位于“**主页**”选项卡或“**外接程序**”选项卡上，具体取决于 Office 主机应用程序版本。

外接程序将在 Vorlon.JS（在 Vorlon.JS 界面左侧）的客户端列表中显示为 **{OS} - n**，*n* 代表数字，而 *{OS}* 表示设备类型，例如“Macintosh”。

![显示 Vorlon.js 界面的屏幕截图](../images/vorlon-interface.png)

Vorlon 工具具有多种插件。当前已启用的插件显示为工具顶部的选项卡。 （可以通过选择左侧的齿轮图标启用更多插件。）这些插件类似于 F12 工具中的功能。 例如，可以突出显示 DOM 元素，执行命令等。 有关详细信息，请参阅 [Vorlon 文档核心插件](http://vorlonjs.com/documentation/#console)。

**Office 外接程序**插件为 Office.js 添加额外的功能，例如探索对象模型、执行 Office.js 调用和读取对象属性的值。有关说明，请参阅[调试 Office 外接程序的 VorlonJS 插件](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)。

> [!NOTE]
> 无法在 Vorlon.JS 中设置断点。


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>在 Mac 或 iPad 上清除 Office 应用程序缓存

出于性能方面的考虑，外接程序通常在 Office for Mac 中缓存。通常情况下，将通过重载外接程序清除缓存。如果同一文档中存在多个外接程序，则重载后自动清除缓存的过程可能不可靠。

在 Mac 上，通过删除 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的所有内容可以手动清除缓存。

在 iPad 上，可以从外接程序中的 JavaScript 调用 `window.location.reload(true)` 来强制重载。或者，可以重新安装 Office。
