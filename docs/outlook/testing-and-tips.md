---
title: 部署和安装 Outlook 加载项以进行测试
description: 创建清单文件，将加载项 UI 文件部署到 Web 服务器，在邮箱中安装加载项，然后测试加载项。
ms.date: 05/20/2020
localization_priority: Priority
ms.openlocfilehash: 86fcf82ddbb9b7e786ae905ad0c0bd864e6fbc56
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348872"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="a10f5-103">部署和安装 Outlook 加载项以进行测试</span><span class="sxs-lookup"><span data-stu-id="a10f5-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="a10f5-104">作为开发 Outlook 外接程序的一个环节，您可能会发现自己在反复部署和安装外接程序以进行测试，这会涉及以下步骤。</span><span class="sxs-lookup"><span data-stu-id="a10f5-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps.</span></span>

1. <span data-ttu-id="a10f5-105">创建描述外接程序的清单文件。</span><span class="sxs-lookup"><span data-stu-id="a10f5-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="a10f5-106">将外接程序 UI 文件部署到 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="a10f5-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="a10f5-107">在邮箱中安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="a10f5-108">测试加载项，对 UI 或清单文件进行适当更改，并重复步骤 2 和步骤 3 来测试更改。</span><span class="sxs-lookup"><span data-stu-id="a10f5-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="a10f5-109">[已弃用自定义窗格](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)，因此请确保正在使用[受支持的加载项扩展点](outlook-add-ins-overview.md#extension-points)。</span><span class="sxs-lookup"><span data-stu-id="a10f5-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="a10f5-110">创建加载项清单文件</span><span class="sxs-lookup"><span data-stu-id="a10f5-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="a10f5-p101">每个外接程序都通过一个 XML 清单进行描述，该文档为服务器提供有关外接程序的信息，为用户提供外接程序的描述性信息，并标识外接程序 UI HTML 文件的位置。您可以在本地文件夹或服务器上存储该清单，只要所测试的邮箱的 Exchange 服务器能够访问这个位置即可。我们假定您在本地文件夹中存储清单。有关如何创建清单文件的信息，请参阅 [Outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="a10f5-p101">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="a10f5-115">将加载项部署到 Web 服务器</span><span class="sxs-lookup"><span data-stu-id="a10f5-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="a10f5-p102">可以使用 HTML 和 JavaScript 创建外接程序。生成的源文件存储在 Web 服务器上，可供托管外接程序的 Exchange 服务器进行访问。在最初部署外接程序的源文件后，可以将 Web 服务器上存储的 HTML 文件或 JavaScript 文件替换为 HTML 文件的新版本，从而更新外接程序 UI 和行为。</span><span class="sxs-lookup"><span data-stu-id="a10f5-p102">You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="a10f5-119">安装加载项</span><span class="sxs-lookup"><span data-stu-id="a10f5-119">Install the add-in</span></span>

<span data-ttu-id="a10f5-120">准备好外接程序清单文件，并将外接程序 UI 部署到可供访问的 Web 服务器后，可以使用 Outlook 客户端为 Exchange 服务器上的邮箱旁加载外接程序，也可以通过运行远程 Windows PowerShell cmdlet 安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="a10f5-121">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="a10f5-121">Sideload the add-in</span></span>

<span data-ttu-id="a10f5-p103">如果邮箱位于 Exchange Online、Exchange 2013 或更高版本上，可以安装外接程序。至少必须拥有 Exchange Server 的 **我的自定义应用程序** 角色，才能旁加载外接程序。若要测试外接程序，或通过指定外接程序清单的 URL 或文件名来常规安装外接程序，应让 Exchange 管理员提供必要权限。</span><span class="sxs-lookup"><span data-stu-id="a10f5-p103">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="a10f5-p104">Exchange 管理员可以运行下列 PowerShell cmdlet，向一个用户分配必要权限。在下面的示例中，`wendyri` 是用户的电子邮件别名。</span><span class="sxs-lookup"><span data-stu-id="a10f5-p104">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="a10f5-127">如有必要，管理员可以运行下列 cmdlet，向多个用户分配类似的必要权限。</span><span class="sxs-lookup"><span data-stu-id="a10f5-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions.</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="a10f5-128">有关我的自定义应用角色的详细信息，请参阅[我的自定义应用角色](/exchange/my-custom-apps-role-exchange-2013-help)。</span><span class="sxs-lookup"><span data-stu-id="a10f5-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="a10f5-129">在使用 Microsoft 365 或 Visual Studio 开发加载项时，会向你分配组织管理员角色，这便允许你按 EAC 中的文件或 URL 或者按 Powershell cmdlet 安装加载项。</span><span class="sxs-lookup"><span data-stu-id="a10f5-129">Using Microsoft 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="a10f5-130">使用远程 PowerShell 安装加载项</span><span class="sxs-lookup"><span data-stu-id="a10f5-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="a10f5-131">在 Exchange 服务器上创建远程 Windows PowerShell 会话后，可以运行 `New-App` cmdlet 和下列 PowerShell 命令，安装 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="a10f5-132">完全限定的 URL 是为外接程序准备的外接程序清单文件的位置。</span><span class="sxs-lookup"><span data-stu-id="a10f5-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="a10f5-133">可以运行下列附加 PowerShell cmdlet，管理邮箱的外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox.</span></span>

- <span data-ttu-id="a10f5-134">`Get-App` - 列出为邮箱启用的外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
- <span data-ttu-id="a10f5-135">`Set-App` - 在邮箱中启用或禁用外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
- <span data-ttu-id="a10f5-136">`Remove-App` - 从 Exchange 服务器中删除以前安装的外接程序。</span><span class="sxs-lookup"><span data-stu-id="a10f5-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="a10f5-137">客户端版本</span><span class="sxs-lookup"><span data-stu-id="a10f5-137">Client versions</span></span>

<span data-ttu-id="a10f5-138">若要决定测试什么版本的 Outlook 客户端，请综合考虑自己的开发需求。</span><span class="sxs-lookup"><span data-stu-id="a10f5-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="a10f5-p105">若要开发供私人使用或仅供组织成员使用的外接程序，请务必测试公司使用的 Outlook 版本。请注意，某些用户可能会使用 Outlook 网页版。因此，还请务必测试公司的标准浏览器版本。</span><span class="sxs-lookup"><span data-stu-id="a10f5-p105">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="a10f5-p106">如果开发的是要在 [AppSource](https://appsource.microsoft.com) 中列出的加载项，必须测试[商业市场认证策略 1120.3](/legal/marketplace/certification-policies#11203-functionality) 中指定的必需版本。这包括：</span><span class="sxs-lookup"><span data-stu-id="a10f5-p106">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:</span></span>
  - <span data-ttu-id="a10f5-143">最新版 Windows 版 Outlook 和前一个版本。</span><span class="sxs-lookup"><span data-stu-id="a10f5-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
  - <span data-ttu-id="a10f5-144">最新版 Mac 版 Outlook。</span><span class="sxs-lookup"><span data-stu-id="a10f5-144">The latest version of Outlook on Mac.</span></span>
  - <span data-ttu-id="a10f5-145">最新 iOS 版和 Android 版 Outlook（如果加载项[支持移动设备规格](add-mobile-support.md)）。</span><span class="sxs-lookup"><span data-stu-id="a10f5-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
  - <span data-ttu-id="a10f5-146">商业市场验证策略 1120.3 中指定的浏览器版本。</span><span class="sxs-lookup"><span data-stu-id="a10f5-146">The browser versions specified in the Commercial marketplace validation policy 1120.3.</span></span>

> [!NOTE]
> <span data-ttu-id="a10f5-147">如果由于[请求的 API 要求集](apis.md)不受客户端支持，导致外接程序不支持上述客户端之一，将从所需客户端列表中删除相应客户端。</span><span class="sxs-lookup"><span data-stu-id="a10f5-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="outlook-on-the-web-and-exchange-server-versions"></a><span data-ttu-id="a10f5-148">Outlook 网页版和 Exchange 服务器版本</span><span class="sxs-lookup"><span data-stu-id="a10f5-148">Outlook on the web and Exchange server versions</span></span>

<span data-ttu-id="a10f5-149">在访问 Outlook 网页版时，消费者和 Microsoft 365 帐户用户将看到新式 UI 版本，而不会再看到已弃用的经典版本。</span><span class="sxs-lookup"><span data-stu-id="a10f5-149">Consumer and Microsoft 365 account users see the modern UI version when they access Outlook on the web and no longer see the classic version which has been deprecated.</span></span> <span data-ttu-id="a10f5-150">但是，本地 Exchange 服务器将继续支持经典 Outlook 网页版。</span><span class="sxs-lookup"><span data-stu-id="a10f5-150">However, on-premises Exchange servers continue to support classic Outlook on the web.</span></span> <span data-ttu-id="a10f5-151">因此，在验证过程中，你的提交可能会收到一条警告，指出加载项与经典 Outlook 网页版不兼容。</span><span class="sxs-lookup"><span data-stu-id="a10f5-151">Therefore, during the validation process, your submission may receive a warning that the add-in is not compatible with classic Outlook on the web.</span></span> <span data-ttu-id="a10f5-152">在这种情况下，需考虑在本地 Exchange 环境中测试加载项。</span><span class="sxs-lookup"><span data-stu-id="a10f5-152">In that case, you should consider testing your add-in in an on-premises Exchange environment.</span></span> <span data-ttu-id="a10f5-153">此警告不会阻止你向 AppSource 提交加载项，但如果消费者是在本地 Exchange 环境中使用 Outlook 网页版，则可能无法获得最佳的体验。</span><span class="sxs-lookup"><span data-stu-id="a10f5-153">This warning won't block your submission to AppSource but your customers may experience a sub-optimal experience if they use Outlook on the web in an on-premises Exchange environment.</span></span>

<span data-ttu-id="a10f5-154">为缓解此问题，我们建议你在连接到你自己的专用本地 Exchange 环境的 Outlook 网页版中对加载项进行测试。</span><span class="sxs-lookup"><span data-stu-id="a10f5-154">To mitigate this, we recommend you test your add-in in Outlook on the web connected to your own private on-premises Exchange environment.</span></span> <span data-ttu-id="a10f5-155">有关详细信息，请参阅有关如何[建立 Exchange 2016 或 Exchange 2019 测试环境](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment)的指南以及有关如何管理[Exchange 服务器中的 Outlook 网页版](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true)的指南。</span><span class="sxs-lookup"><span data-stu-id="a10f5-155">For more information, see guidance on how to [Establish an Exchange 2016 or Exchange 2019 test environment](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment) and how to manage [Outlook on the web in Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true).</span></span>

<span data-ttu-id="a10f5-p109">或者，你也可以选择付费并使用托管和管理本地 Exchange 服务器的服务。有两个选项：</span><span class="sxs-lookup"><span data-stu-id="a10f5-p109">Alternatively, you can opt to pay for and use a service that hosts and manages on-premises Exchange servers. A couple of options are:</span></span>

- [<span data-ttu-id="a10f5-158">Rackspace</span><span class="sxs-lookup"><span data-stu-id="a10f5-158">Rackspace</span></span>](https://www.rackspace.com/email-hosting/exchange-server)
- [<span data-ttu-id="a10f5-159">Hostway</span><span class="sxs-lookup"><span data-stu-id="a10f5-159">Hostway</span></span>](https://hostway.com/products-services-2/hosted-microsoft-exchange/)

<span data-ttu-id="a10f5-160">此外，如果不想面向连接到本地 Exchange 的用户提供自己的加载项，可将加载项清单中的[要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support)设置为 1.6 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="a10f5-160">Furthermore, if you don't want your add-ins to be available for users who are connected to on-premises Exchange, you can set the [requirement set](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) in the add-in manifest to be 1.6 or higher.</span></span> <span data-ttu-id="a10f5-161">在经典 Outlook 网页版上，不会对此类加载项进行测试或验证。</span><span class="sxs-lookup"><span data-stu-id="a10f5-161">Such add-ins will not be tested or validated on the classic Outlook on the web UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="a10f5-162">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a10f5-162">See also</span></span>

- [<span data-ttu-id="a10f5-163">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="a10f5-163">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
