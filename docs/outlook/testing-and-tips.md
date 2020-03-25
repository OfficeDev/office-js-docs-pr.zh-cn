---
title: 部署和安装 Outlook 加载项以进行测试
description: 创建清单文件，将加载项 UI 文件部署到 Web 服务器，在邮箱中安装加载项，然后测试加载项。
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 76688ad3e1eca2dda832a94c3a9ae815e37678bc
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890975"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="cbc28-103">部署和安装 Outlook 加载项以进行测试</span><span class="sxs-lookup"><span data-stu-id="cbc28-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="cbc28-104">作为开发 Outlook 外接程序的一个环节，您可能会发现自己在反复部署和安装外接程序以进行测试，这会涉及以下步骤：</span><span class="sxs-lookup"><span data-stu-id="cbc28-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:</span></span>

1. <span data-ttu-id="cbc28-105">创建描述外接程序的清单文件。</span><span class="sxs-lookup"><span data-stu-id="cbc28-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="cbc28-106">将外接程序 UI 文件部署到 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="cbc28-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="cbc28-107">在邮箱中安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="cbc28-108">测试加载项，对 UI 或清单文件进行适当更改，并重复步骤 2 和步骤 3 来测试更改。</span><span class="sxs-lookup"><span data-stu-id="cbc28-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="cbc28-109">[已弃用自定义窗格](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)，因此请确保正在使用[受支持的加载项扩展点](outlook-add-ins-overview.md#extension-points)。</span><span class="sxs-lookup"><span data-stu-id="cbc28-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="cbc28-110">创建加载项清单文件</span><span class="sxs-lookup"><span data-stu-id="cbc28-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="cbc28-p101">每个外接程序都通过一个 XML 清单进行描述，该文档为服务器提供有关外接程序的信息，为用户提供外接程序的描述性信息，并标识外接程序 UI HTML 文件的位置。您可以在本地文件夹或服务器上存储该清单，只要所测试的邮箱的 Exchange 服务器能够访问这个位置即可。我们假定您在本地文件夹中存储清单。有关如何创建清单文件的信息，请参阅 [Outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="cbc28-p101">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="cbc28-115">将加载项部署到 Web 服务器</span><span class="sxs-lookup"><span data-stu-id="cbc28-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="cbc28-p102">可以使用 HTML 和 JavaScript 创建外接程序。生成的源文件存储在 Web 服务器上，可供托管外接程序的 Exchange 服务器进行访问。在最初部署外接程序的源文件后，可以将 Web 服务器上存储的 HTML 文件或 JavaScript 文件替换为 HTML 文件的新版本，从而更新外接程序 UI 和行为。</span><span class="sxs-lookup"><span data-stu-id="cbc28-p102">You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="cbc28-119">安装加载项</span><span class="sxs-lookup"><span data-stu-id="cbc28-119">Install the add-in</span></span>

<span data-ttu-id="cbc28-120">准备好外接程序清单文件，并将外接程序 UI 部署到可供访问的 Web 服务器后，可以使用 Outlook 客户端为 Exchange 服务器上的邮箱旁加载外接程序，也可以通过运行远程 Windows PowerShell cmdlet 安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="cbc28-121">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="cbc28-121">Sideload the add-in</span></span>

<span data-ttu-id="cbc28-p103">如果邮箱位于 Exchange Online、Exchange 2013 或更高版本上，可以安装外接程序。至少必须拥有 Exchange Server 的**我的自定义应用程序**角色，才能旁加载外接程序。若要测试外接程序，或通过指定外接程序清单的 URL 或文件名来常规安装外接程序，应让 Exchange 管理员提供必要权限。</span><span class="sxs-lookup"><span data-stu-id="cbc28-p103">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="cbc28-p104">Exchange 管理员可以运行下列 PowerShell cmdlet，向一个用户分配必要权限。在下面的示例中，`wendyri` 是用户的电子邮件别名。</span><span class="sxs-lookup"><span data-stu-id="cbc28-p104">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="cbc28-127">如有必要，管理员可以运行下列 cmdlet，向多个用户分配类似的必要权限：</span><span class="sxs-lookup"><span data-stu-id="cbc28-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="cbc28-128">有关我的自定义应用角色的详细信息，请参阅[我的自定义应用角色](/exchange/my-custom-apps-role-exchange-2013-help)。</span><span class="sxs-lookup"><span data-stu-id="cbc28-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="cbc28-129">使用 Office 365 或 Visual Studio 开发外接程序会向你分配组织管理员角色，这便允许你按 EAC 中的文件或 URL 或者按 Powershell cmdlet 安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-129">Using Office 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="cbc28-130">使用远程 PowerShell 安装加载项</span><span class="sxs-lookup"><span data-stu-id="cbc28-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="cbc28-131">在 Exchange 服务器上创建远程 Windows PowerShell 会话后，可以运行 `New-App` cmdlet 和下列 PowerShell 命令，安装 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="cbc28-132">完全限定的 URL 是为外接程序准备的外接程序清单文件的位置。</span><span class="sxs-lookup"><span data-stu-id="cbc28-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="cbc28-133">可以运行下列附加 PowerShell cmdlet，管理邮箱的外接程序：</span><span class="sxs-lookup"><span data-stu-id="cbc28-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:</span></span>

-  <span data-ttu-id="cbc28-134">`Get-App` - 列出为邮箱启用的外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
-  <span data-ttu-id="cbc28-135">`Set-App` - 在邮箱中启用或禁用外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
-  <span data-ttu-id="cbc28-136">`Remove-App` - 从 Exchange 服务器中删除以前安装的外接程序。</span><span class="sxs-lookup"><span data-stu-id="cbc28-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="cbc28-137">客户端版本</span><span class="sxs-lookup"><span data-stu-id="cbc28-137">Client versions</span></span>

<span data-ttu-id="cbc28-138">若要决定测试什么版本的 Outlook 客户端，请综合考虑自己的开发需求。</span><span class="sxs-lookup"><span data-stu-id="cbc28-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="cbc28-p105">若要开发供私人使用或仅供组织成员使用的外接程序，请务必测试公司使用的 Outlook 版本。请注意，某些用户可能会使用 Outlook 网页版。因此，还请务必测试公司的标准浏览器版本。</span><span class="sxs-lookup"><span data-stu-id="cbc28-p105">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="cbc28-p106">如果开发的是要在 [AppSource](https://appsource.microsoft.com) 中列出的加载项，必须测试[商业市场认证策略 1120.3](/legal/marketplace/certification-policies#11203-functionality) 中指定的必需版本。这包括：</span><span class="sxs-lookup"><span data-stu-id="cbc28-p106">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:</span></span>
    - <span data-ttu-id="cbc28-143">最新版 Windows 版 Outlook 和前一个版本。</span><span class="sxs-lookup"><span data-stu-id="cbc28-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
    - <span data-ttu-id="cbc28-144">最新版 Mac 版 Outlook。</span><span class="sxs-lookup"><span data-stu-id="cbc28-144">The latest version of Outlook on Mac.</span></span>
    - <span data-ttu-id="cbc28-145">最新 iOS 版和 Android 版 Outlook（如果加载项[支持移动设备规格](add-mobile-support.md)）。</span><span class="sxs-lookup"><span data-stu-id="cbc28-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
    - <span data-ttu-id="cbc28-146">商业市场验证策略 1120.3 中指定的浏览器版本。</span><span class="sxs-lookup"><span data-stu-id="cbc28-146">The browser versions specified in the Commercial marketplace validation policy 1120.3.</span></span>

> [!NOTE]
> <span data-ttu-id="cbc28-147">如果由于[请求的 API 要求集](apis.md)不受客户端支持，导致外接程序不支持上述客户端之一，将从所需客户端列表中删除相应客户端。</span><span class="sxs-lookup"><span data-stu-id="cbc28-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="see-also"></a><span data-ttu-id="cbc28-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cbc28-148">See also</span></span>

- [<span data-ttu-id="cbc28-149">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="cbc28-149">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
