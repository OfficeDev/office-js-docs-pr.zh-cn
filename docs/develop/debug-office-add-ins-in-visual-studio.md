---
title: 在 Visual Studio 中调试 Office 加载项
description: 使用 Visual Studio 在 Windows 上的 Office 桌面客户端中调试 Office 加载项
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 018bfa24424514598d323c29d165e3e8ec066a8e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093656"
---
# <a name="debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="d5342-103">在 Visual Studio 中调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d5342-103">Debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="d5342-104">本文介绍如何使用 Visual Studio 2019 在 Windows 上的 Office 桌面客户端中调试 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5342-104">This article describes how to use Visual Studio 2019 to debug an Office Add-in in the Office desktop client on Windows.</span></span> <span data-ttu-id="d5342-105">如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。</span><span class="sxs-lookup"><span data-stu-id="d5342-105">If you're using another version of Visual Studio, the procedures might vary slightly.</span></span> 

> [!NOTE]
> <span data-ttu-id="d5342-106">无法使用 Visual Studio 在 Office 网页版或 Mac 版 Office 中调试加载项。</span><span class="sxs-lookup"><span data-stu-id="d5342-106">You can't use Visual Studio to debug add-ins in Office on the web or Mac.</span></span> <span data-ttu-id="d5342-107">若要了解如何在这些平台上进行调试，请参阅[在 Office 网页版中调试 Office 加载项](../testing/debug-add-ins-in-office-online.md)或[在 Mac 上调试 Office 加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="d5342-107">For information about debugging on these platforms, see [Debug Office Add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md) or [Debug Office Add-ins on a Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).</span></span>

## <a name="enable-debugging-for-add-in-commands-and-ui-less-code"></a><span data-ttu-id="d5342-108">对加载项命令和无 UI 的代码启用调试</span><span class="sxs-lookup"><span data-stu-id="d5342-108">Enable debugging for add-in commands and UI-less code</span></span>

<span data-ttu-id="d5342-109">当 Visual Studio 调试 Windows 上的 Office 时，加载项托管在 Microsoft Internet Explorer 或 Microsoft Edge 浏览器实例中。</span><span class="sxs-lookup"><span data-stu-id="d5342-109">When Visual Studio debugs Office on Windows, the add-in is hosted in either a Microsoft Internet Explorer or Microsoft Edge browser instance.</span></span> <span data-ttu-id="d5342-110">若要确定开发计算机上使用的浏览器，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="d5342-110">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>
> [!NOTE]
> <span data-ttu-id="d5342-111">以下过程不再需要 JS_Debug 环境变量。</span><span class="sxs-lookup"><span data-stu-id="d5342-111">The JS_Debug environment variable is no longer needed in the following procedure.</span></span> <span data-ttu-id="d5342-112">有关详细信息，请参阅 Microsoft 开发人员社区支持论坛中的 [Office Web 加载项中的调试行为](https://developercommunity.visualstudio.com/content/problem/740413/office-development-inconsistent-script-debugging-b.html)。</span><span class="sxs-lookup"><span data-stu-id="d5342-112">For more information, see [Debugging behaviors in Office web add-ins](https://developercommunity.visualstudio.com/content/problem/740413/office-development-inconsistent-script-debugging-b.html) on the Microsoft Developer Community support forum.</span></span>

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

## <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="d5342-113">查看生成和调试属性</span><span class="sxs-lookup"><span data-stu-id="d5342-113">Review the build and debug properties</span></span>

<span data-ttu-id="d5342-114">在开始调试之前，请检查每个项目的属性以确认 Visual Studio 将打开所需的主机应用程序，并已正确设置其他生成和调试属性。</span><span class="sxs-lookup"><span data-stu-id="d5342-114">Before you start debugging, review the properties of each project to confirm that Visual Studio will open the desired host application and that other build and debug properties are set appropriately.</span></span>

### <a name="add-in-project-properties"></a><span data-ttu-id="d5342-115">外接程序项目属性</span><span class="sxs-lookup"><span data-stu-id="d5342-115">Add-in project properties</span></span>

<span data-ttu-id="d5342-116">打开外接程序项目的“**属性**”窗口以查看项目属性：</span><span class="sxs-lookup"><span data-stu-id="d5342-116">Open the **Properties** window for the add-in project to review project properties:</span></span>

1. <span data-ttu-id="d5342-117">在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。</span><span class="sxs-lookup"><span data-stu-id="d5342-117">In **Solution Explorer**, choose the add-in project (*not* the web application project).</span></span>

2. <span data-ttu-id="d5342-118">在菜单栏中，依次选择“**视图**” > “**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="d5342-118">From the menu bar, choose **View** > **Properties Window**.</span></span>

<span data-ttu-id="d5342-119">下表介绍了外接程序项目的属性。</span><span class="sxs-lookup"><span data-stu-id="d5342-119">The following table describes the properties of the add-in project.</span></span>

|<span data-ttu-id="d5342-120">**属性**</span><span class="sxs-lookup"><span data-stu-id="d5342-120">**Property**</span></span>|<span data-ttu-id="d5342-121">**说明**</span><span class="sxs-lookup"><span data-stu-id="d5342-121">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="d5342-122">**启动操作**</span><span class="sxs-lookup"><span data-stu-id="d5342-122">**Start Action**</span></span>|<span data-ttu-id="d5342-123">指定外接程序的调试模式。</span><span class="sxs-lookup"><span data-stu-id="d5342-123">Specifies the debug mode for your add-in.</span></span> <span data-ttu-id="d5342-124">目前，Office 外接程序项目仅支持 **Office 桌面客户端**模式。</span><span class="sxs-lookup"><span data-stu-id="d5342-124">Currently only **Office Desktop Client** mode is supported for Office Add-in projects.</span></span>|
|<span data-ttu-id="d5342-125">**启动文档**</span><span class="sxs-lookup"><span data-stu-id="d5342-125">**Start Document**</span></span><br/><span data-ttu-id="d5342-126">（仅限 Excel、PowerPoint 和 Word 外接程序）</span><span class="sxs-lookup"><span data-stu-id="d5342-126">(Excel, PowerPoint, and Word add-ins only)</span></span>|<span data-ttu-id="d5342-127">指定要在启动项目时打开的文档。</span><span class="sxs-lookup"><span data-stu-id="d5342-127">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="d5342-128">**Web 项目**</span><span class="sxs-lookup"><span data-stu-id="d5342-128">**Web Project**</span></span>|<span data-ttu-id="d5342-129">指定与外接程序关联的 Web 项目的名称。</span><span class="sxs-lookup"><span data-stu-id="d5342-129">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="d5342-130">**电子邮件地址**</span><span class="sxs-lookup"><span data-stu-id="d5342-130">**Email Address**</span></span><br/><span data-ttu-id="d5342-131">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="d5342-131">(Outlook add-ins only)</span></span>|<span data-ttu-id="d5342-132">指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d5342-132">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in.</span></span>|
|<span data-ttu-id="d5342-133">**EWS Url**</span><span class="sxs-lookup"><span data-stu-id="d5342-133">**EWS Url**</span></span><br/><span data-ttu-id="d5342-134">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="d5342-134">(Outlook add-ins only)</span></span>|<span data-ttu-id="d5342-135">Exchange Web 服务 URL（例如：`https://www.contoso.com/ews/exchange.aspx`）。</span><span class="sxs-lookup"><span data-stu-id="d5342-135">Exchange Web service URL (For example: `https://www.contoso.com/ews/exchange.aspx`).</span></span> |
|<span data-ttu-id="d5342-136">**OWA Url**</span><span class="sxs-lookup"><span data-stu-id="d5342-136">**OWA Url**</span></span><br/><span data-ttu-id="d5342-137">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="d5342-137">(Outlook add-ins only)</span></span>|<span data-ttu-id="d5342-138">Outlook 网页版 URL（例如：`https://www.contoso.com/owa`）。</span><span class="sxs-lookup"><span data-stu-id="d5342-138">Outlook on the web URL (For example: `https://www.contoso.com/owa`).</span></span>|
|<span data-ttu-id="d5342-139">**使用多重身份验证**</span><span class="sxs-lookup"><span data-stu-id="d5342-139">**Use multi-factor auth**</span></span><br/><span data-ttu-id="d5342-140">（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="d5342-140">(Outlook add-ins only)</span></span>|<span data-ttu-id="d5342-141">布尔值，指示是否应使用多重身份验证。</span><span class="sxs-lookup"><span data-stu-id="d5342-141">Boolean value that indicates whether multi-factor authentication should be used.</span></span>|
|<span data-ttu-id="d5342-142">**用户名**</span><span class="sxs-lookup"><span data-stu-id="d5342-142">**User Name**</span></span><br/><span data-ttu-id="d5342-143">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="d5342-143">(Outlook add-ins only)</span></span>|<span data-ttu-id="d5342-144">指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的名称。</span><span class="sxs-lookup"><span data-stu-id="d5342-144">Specifies the name of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in.</span></span>|
|<span data-ttu-id="d5342-145">**项目文件**</span><span class="sxs-lookup"><span data-stu-id="d5342-145">**Project File**</span></span>|<span data-ttu-id="d5342-146">指定包含生成、配置和有关项目的其他信息的文件名称。</span><span class="sxs-lookup"><span data-stu-id="d5342-146">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="d5342-147">**项目文件夹**</span><span class="sxs-lookup"><span data-stu-id="d5342-147">**Project Folder**</span></span>|<span data-ttu-id="d5342-148">项目文件的位置。</span><span class="sxs-lookup"><span data-stu-id="d5342-148">The location of the project file.</span></span>|

> [!NOTE]
> <span data-ttu-id="d5342-149">对于 Outlook 外接程序，你可以选择在“**属性**”窗口中为一个或多个 *Outlook 外接程序*属性指定值，但这样做并不是必须的。</span><span class="sxs-lookup"><span data-stu-id="d5342-149">For an Outlook add-in, you may choose to specify values for one or more of the *Outlook add-in only* properties in the **Properties** window, but doing so is not required.</span></span>

### <a name="web-application-project-properties"></a><span data-ttu-id="d5342-150">Web 应用程序项目属性</span><span class="sxs-lookup"><span data-stu-id="d5342-150">Web application project properties</span></span>

<span data-ttu-id="d5342-151">打开 Web 应用程序项目的“**属性**”窗口以查看项目属性：</span><span class="sxs-lookup"><span data-stu-id="d5342-151">Open the **Properties** window for the web application project to review project properties:</span></span>

1. <span data-ttu-id="d5342-152">在 "**解决方案资源管理器**" 中，选择 "web 应用程序" 项目。</span><span class="sxs-lookup"><span data-stu-id="d5342-152">In **Solution Explorer**, choose the web application project.</span></span>

2. <span data-ttu-id="d5342-153">在菜单栏中，依次选择“**视图**” > “**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="d5342-153">From the menu bar, choose **View** > **Properties Window**.</span></span>

<span data-ttu-id="d5342-154">下表介绍了与 Office 外接程序项目最相关的 Web 应用程序项目的属性。</span><span class="sxs-lookup"><span data-stu-id="d5342-154">The following table describes the properties of the web application project that are most relevant to Office Add-in projects.</span></span>

|<span data-ttu-id="d5342-155">**属性**</span><span class="sxs-lookup"><span data-stu-id="d5342-155">**Property**</span></span>|<span data-ttu-id="d5342-156">**说明**</span><span class="sxs-lookup"><span data-stu-id="d5342-156">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="d5342-157">**SSL 已启用**</span><span class="sxs-lookup"><span data-stu-id="d5342-157">**SSL Enabled**</span></span>|<span data-ttu-id="d5342-158">指定是否在站点上启用 SSL。</span><span class="sxs-lookup"><span data-stu-id="d5342-158">Specifies whether SSL is enabled on the site.</span></span> <span data-ttu-id="d5342-159">对于 Office 外接程序项目，此属性应设置为 **True**。</span><span class="sxs-lookup"><span data-stu-id="d5342-159">This property should be set to **True** for Office Add-in projects.</span></span>|
|<span data-ttu-id="d5342-160">**SSL URL**</span><span class="sxs-lookup"><span data-stu-id="d5342-160">**SSL URL**</span></span>|<span data-ttu-id="d5342-161">指定站点的安全 HTTPS URL。</span><span class="sxs-lookup"><span data-stu-id="d5342-161">Specifies the secure HTTPS URL for the site.</span></span> <span data-ttu-id="d5342-162">只读。</span><span class="sxs-lookup"><span data-stu-id="d5342-162">Read-only.</span></span>|
|<span data-ttu-id="d5342-163">**URL**</span><span class="sxs-lookup"><span data-stu-id="d5342-163">**URL**</span></span>|<span data-ttu-id="d5342-164">指定站点的 HTTP URL。</span><span class="sxs-lookup"><span data-stu-id="d5342-164">Specifies the HTTP URL for the site.</span></span> <span data-ttu-id="d5342-165">只读。</span><span class="sxs-lookup"><span data-stu-id="d5342-165">Read-only.</span></span>|
|<span data-ttu-id="d5342-166">**项目文件**</span><span class="sxs-lookup"><span data-stu-id="d5342-166">**Project File**</span></span>|<span data-ttu-id="d5342-167">指定包含生成、配置和有关项目的其他信息的文件名称。</span><span class="sxs-lookup"><span data-stu-id="d5342-167">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="d5342-168">**项目文件夹**</span><span class="sxs-lookup"><span data-stu-id="d5342-168">**Project Folder**</span></span>|<span data-ttu-id="d5342-169">指定项目文件的位置。</span><span class="sxs-lookup"><span data-stu-id="d5342-169">Specifies the location of the project file.</span></span> <span data-ttu-id="d5342-170">只读。</span><span class="sxs-lookup"><span data-stu-id="d5342-170">Read-only.</span></span> <span data-ttu-id="d5342-171">Visual Studio 在运行时生成的清单文件将写入到此位置的 `bin\Debug\OfficeAppManifests` 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="d5342-171">The manifest file that Visual Studio generates at runtime is written to the `bin\Debug\OfficeAppManifests` folder in this location.</span></span>|

## <a name="use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="d5342-172">使用现有文档调试外接程序</span><span class="sxs-lookup"><span data-stu-id="d5342-172">Use an existing document to debug the add-in</span></span>

<span data-ttu-id="d5342-173">如果你有一个文档包含要在调试 Excel、PowerPoint 或 Word 外接程序时使用的测试数据，则可以将 Visual Studio 配置为在启动项目时打开该文档。</span><span class="sxs-lookup"><span data-stu-id="d5342-173">If you have a document that contains test data you want to use while debugging your Excel, PowerPoint, or Word add-in, Visual Studio can be configured to open that document when you start the project.</span></span> <span data-ttu-id="d5342-174">若要指定在调试外接程序时要使用的现有文档，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="d5342-174">To specify an existing document to use while debugging the add-in, complete the following steps.</span></span>

1. <span data-ttu-id="d5342-175">在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。</span><span class="sxs-lookup"><span data-stu-id="d5342-175">In **Solution Explorer**, choose the add-in project (*not* the web application project).</span></span>

2. <span data-ttu-id="d5342-176">从菜单栏中，选择“**项目**” > “**添加现有项**”。</span><span class="sxs-lookup"><span data-stu-id="d5342-176">From the menu bar, choose **Project** > **Add Existing Item**.</span></span>

3. <span data-ttu-id="d5342-177">在“**添加现有项**”对话框中，找到并选择要添加的文档。</span><span class="sxs-lookup"><span data-stu-id="d5342-177">In the **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>

4. <span data-ttu-id="d5342-178">选择“**添加**”按钮以将文档添加到项目中。</span><span class="sxs-lookup"><span data-stu-id="d5342-178">Choose the **Add** button to add the document to your project.</span></span>

5. <span data-ttu-id="d5342-179">在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。</span><span class="sxs-lookup"><span data-stu-id="d5342-179">In **Solution Explorer**, choose the add-in project (*not* the web application project).</span></span>

6. <span data-ttu-id="d5342-180">在菜单栏中，依次选择“**视图**” > “**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="d5342-180">From the menu bar, choose **View** > **Properties Window**.</span></span>

7. <span data-ttu-id="d5342-181">在“**属性**”窗口中，选择“**启动文档**”列表，然后选择添加到项目中的文档。</span><span class="sxs-lookup"><span data-stu-id="d5342-181">In the **Properties** window, choose the **Start Document** list, and then select the document that you added to the project.</span></span> <span data-ttu-id="d5342-182">该项目现在配置为在该文档中启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5342-182">The project is now configured to start the add-in in that document.</span></span>

## <a name="start-the-project"></a><span data-ttu-id="d5342-183">启动项目</span><span class="sxs-lookup"><span data-stu-id="d5342-183">Start the project</span></span>

<span data-ttu-id="d5342-184">从菜单栏中依次选择“**调试**” > “**开始调试**”，可启动项目。</span><span class="sxs-lookup"><span data-stu-id="d5342-184">Start the project by choosing **Debug** > **Start Debugging** from the menu bar.</span></span> <span data-ttu-id="d5342-185">Visual Studio 将自动生成解决方案并启动 Office 以托管外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5342-185">Visual Studio will automatically build the solution and start Office to host your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="d5342-186">启动 Outlook 外接程序项目时，系统会提示你输入登录凭据。</span><span class="sxs-lookup"><span data-stu-id="d5342-186">When you start an Outlook add-in project, you'll be prompted for login credentials.</span></span> <span data-ttu-id="d5342-187">如果要求您反复登录，或者如果您收到未经授权的错误，则可能会对 Microsoft 365 租户上的帐户禁用基本身份验证。</span><span class="sxs-lookup"><span data-stu-id="d5342-187">If you're asked to log in repeatedly or if you receive an error that you are unauthorized, then Basic Auth may be disabled for accounts on your Microsoft 365 tenant.</span></span> <span data-ttu-id="d5342-188">在这种情况下，请尝试使用 Microsoft 帐户。</span><span class="sxs-lookup"><span data-stu-id="d5342-188">In this case, try using a Microsoft account instead.</span></span> <span data-ttu-id="d5342-189">可能还需要在“Outlook Web 加载项”项目属性对话框中将属性“使用多重身份验证”设置为 True。</span><span class="sxs-lookup"><span data-stu-id="d5342-189">You may also need to set the property "Use multi-factor auth" to True in the Outlook Web Add-in project properties dialog.</span></span>

<span data-ttu-id="d5342-190">当 Visual Studio 生成项目时，它执行以下任务：</span><span class="sxs-lookup"><span data-stu-id="d5342-190">When Visual Studio builds the project it performs the following tasks:</span></span>

1. <span data-ttu-id="d5342-191">创建 XML 清单文件的副本并将其添加到 `_ProjectName_\bin\Debug\OfficeAppManifests` 目录。</span><span class="sxs-lookup"><span data-stu-id="d5342-191">Creates a copy of the XML manifest file and adds it to  `_ProjectName_\bin\Debug\OfficeAppManifests` directory.</span></span> <span data-ttu-id="d5342-192">启动 Visual Studio 并调试外接程序时，主机应用程序将使用此副本。</span><span class="sxs-lookup"><span data-stu-id="d5342-192">The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>

2. <span data-ttu-id="d5342-193">在计算机上创建一组允许外接程序在主机应用程序中显示的注册表项。</span><span class="sxs-lookup"><span data-stu-id="d5342-193">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>

3. <span data-ttu-id="d5342-194">生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (https://localhost))。</span><span class="sxs-lookup"><span data-stu-id="d5342-194">Builds the web application project, and then deploys it to the local IIS web server (https://localhost).</span></span>

4. <span data-ttu-id="d5342-195">如果这是你已部署到本地 IIS Web 服务器的第一个加载项项目，系统可能会提示你将自签名证书安装到当前用户的受信任的根证书存储中。</span><span class="sxs-lookup"><span data-stu-id="d5342-195">If this is the first add-in project that you have deployed to local IIS web server, you may be prompted to install a Self-Signed Certificate to the current user's Trusted Root Certificate store.</span></span> <span data-ttu-id="d5342-196">若要使 IIS Express 正确显示加载项内容，这是必需的操作。</span><span class="sxs-lookup"><span data-stu-id="d5342-196">This is required for IIS Express to display the content of your add-in correctly.</span></span>

> [!NOTE]
> <span data-ttu-id="d5342-197">在 Windows 10 上运行时，最新版本的 Office 可能会使用较新的 Web 控件来显示加载项内容。</span><span class="sxs-lookup"><span data-stu-id="d5342-197">The latest version of Office may use a newer web control to display the add-in contents when running on Windows 10.</span></span> <span data-ttu-id="d5342-198">如果是这种情况，Visual Studio 可能会提示你添加本地网络环回豁免。</span><span class="sxs-lookup"><span data-stu-id="d5342-198">If this is the case, Visual Studio may prompt you to add a local network loopback exemption.</span></span> <span data-ttu-id="d5342-199">在 Office 主机应用程序中，需要这样做才能使 Web 控件访问部署到本地 IIS Web 服务器的网站。</span><span class="sxs-lookup"><span data-stu-id="d5342-199">This is required for the web control, in the Office host application, to be able to access the website deployed to the local IIS web server.</span></span> <span data-ttu-id="d5342-200">还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="d5342-200">You can also change this setting anytime in Visual Studio under **Tools** > **Options** > **Office Tools (Web)** > **Web Add-In Debugging**.</span></span>

<span data-ttu-id="d5342-201">接下来，Visual Studio 会执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="d5342-201">Next, Visual Studio does the following:</span></span>

1. <span data-ttu-id="d5342-202">通过将 `~remoteAppUrl` 标记替换为起始页的完全限定地址（例如，`https://localhost:44302/Home.html`）来修改 XML 清单文件的 [SourceLocation](../reference/manifest/sourcelocation.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="d5342-202">Modifies the [SourceLocation](../reference/manifest/sourcelocation.md) element of the XML manifest file by replacing the `~remoteAppUrl` token with the fully qualified address of the start page (for example, `https://localhost:44302/Home.html`).</span></span>

2. <span data-ttu-id="d5342-203">在 IIS Express 中启动 Web 应用程序项目。</span><span class="sxs-lookup"><span data-stu-id="d5342-203">Starts the web application project in IIS Express.</span></span>

3. <span data-ttu-id="d5342-204">打开主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="d5342-204">Opens the host application.</span></span>

<span data-ttu-id="d5342-205">当您构建项目时，Visual Studio 不会在“输出”\*\*\*\* 窗口中显示验证错误。</span><span class="sxs-lookup"><span data-stu-id="d5342-205">Visual Studio doesn't show validation errors in the **OUTPUT** window when you build the project.</span></span> <span data-ttu-id="d5342-206">Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。</span><span class="sxs-lookup"><span data-stu-id="d5342-206">Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur.</span></span> <span data-ttu-id="d5342-207">通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。</span><span class="sxs-lookup"><span data-stu-id="d5342-207">Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor.</span></span> <span data-ttu-id="d5342-208">通过这些标志，你可以得知 Visual Studio 在你的代码中检测到的问题。</span><span class="sxs-lookup"><span data-stu-id="d5342-208">These marks notify you of problems that Visual Studio detected in your code.</span></span> <span data-ttu-id="d5342-209">有关如何启用或禁用验证的详细信息，请参阅[选项、文本编辑器、JavaScript、IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019)。</span><span class="sxs-lookup"><span data-stu-id="d5342-209">For more information about how to enable or disable validation, see [Options, Text Editor, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019).</span></span>

<span data-ttu-id="d5342-210">要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="d5342-210">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a><span data-ttu-id="d5342-211">调试 Excel、PowerPoint 或 Word 外接程序的代码</span><span class="sxs-lookup"><span data-stu-id="d5342-211">Debug the code for an Excel, PowerPoint, or Word add-in</span></span>

<span data-ttu-id="d5342-212">如果在[启动项目](#start-the-project)后，在主机应用程序（Excel、PowerPoint 或 Word）中显示的文档中看不到外接程序，请在主机应用程序中手动启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5342-212">If your add-in isn't visible within the document that's displayed in the host application (Excel, PowerPoint, or Word) after you've [started the project](#start-the-project), manually launch the add-in in the host application.</span></span> <span data-ttu-id="d5342-213">例如，通过选择“**主页**”选项卡功能区中的“**显示任务窗格**”按钮来启动任务窗格外接程序。在 Excel、PowerPoint 或 Word 中显示外接程序后，你可以通过执行以下操作来调试代码：</span><span class="sxs-lookup"><span data-stu-id="d5342-213">For example, launch your task pane add-in by choosing the **Show Taskpane** button in the ribbon of the **Home** tab. After your add-in is displayed in Excel, PowerPoint, or Word, you can debug your code by doing the following:</span></span>

1. <span data-ttu-id="d5342-214">在 Excel、PowerPoint 或 Word 中，选择“**插入**”选项卡，然后选择“**我的外接程序**”右侧的向下箭头。</span><span class="sxs-lookup"><span data-stu-id="d5342-214">In Excel, PowerPoint, or Word, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.</span></span>

    ![Windows 版 Excel 的“插入”功能区及突出显示的“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)

2. <span data-ttu-id="d5342-216">在可用外接程序列表中，找到“**开发人员外接程序**”部分并选择你的外接程序进行注册。</span><span class="sxs-lookup"><span data-stu-id="d5342-216">In the list of available add-ins, find the **Developer Add-ins** section and select the your add-in to register it.</span></span>

3. <span data-ttu-id="d5342-217">在 Visual Studio 中，在代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="d5342-217">In Visual Studio, set breakpoints in your code.</span></span>

4. <span data-ttu-id="d5342-218">在 Excel、PowerPoint 或 Word 中，与外接程序进行交互。</span><span class="sxs-lookup"><span data-stu-id="d5342-218">In Excel, PowerPoint, or Word, interact with your add-in.</span></span>

5. <span data-ttu-id="d5342-219">在 Visual Studio 中命中断点时，根据需要逐步执行代码。</span><span class="sxs-lookup"><span data-stu-id="d5342-219">As breakpoints are hit in Visual Studio, step through the code as needed.</span></span>

<span data-ttu-id="d5342-220">你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭主机应用程序并重新启动该项目。</span><span class="sxs-lookup"><span data-stu-id="d5342-220">You can change your code and review the effects of those changes in your add-in without having to close the host application and restart the project.</span></span> <span data-ttu-id="d5342-221">保存对代码的更改后，只需在主机应用程序中重新加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5342-221">After you save changes to your code, simply reload the add-in in the host application.</span></span> <span data-ttu-id="d5342-222">例如，通过选择任务窗格的右上角来激活[个性菜单](../design/task-pane-add-ins.md#personality-menu)，然后选择“**重新加载**”，便可重新加载任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5342-222">For example, reload a task pane add-in by choosing the top-right corner of the task pane to activate the [personality menu](../design/task-pane-add-ins.md#personality-menu) and then choose **Reload**.</span></span>

## <a name="debug-the-code-for-an-outlook-add-in"></a><span data-ttu-id="d5342-223">调试 Outlook 外接程序的代码</span><span class="sxs-lookup"><span data-stu-id="d5342-223">Debug the code for an Outlook add-in</span></span>

<span data-ttu-id="d5342-224">在你已[启动项目](#start-the-project)，且 Visual Studio 启动 Outlook 来托管外接程序后，打开电子邮件或约会项目。</span><span class="sxs-lookup"><span data-stu-id="d5342-224">After you've [started the project](#start-the-project) and Visual Studio launches Outlook to host your add-in, open an email message or appointment item.</span></span> 

<span data-ttu-id="d5342-p120">只要满足激活条件，Outlook 便会为项目激活外接程序。外接程序栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 外接程序显示为外接程序栏中的一个按钮。如果您的外接程序有外接程序命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该外接程序将不会显示在外接程序栏中。</span><span class="sxs-lookup"><span data-stu-id="d5342-p120">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="d5342-228">若要查看 Outlook 外接程序，请选择对应 Outlook 外接程序的按钮。</span><span class="sxs-lookup"><span data-stu-id="d5342-228">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span> <span data-ttu-id="d5342-229">在 Outlook 中显示外接程序后，你可以通过执行以下操作来调试代码：</span><span class="sxs-lookup"><span data-stu-id="d5342-229">After your add-in is displayed in Outlook, you can debug your code by doing the following:</span></span>

1. <span data-ttu-id="d5342-230">在 Visual Studio 中，在代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="d5342-230">In Visual Studio, set breakpoints in your code.</span></span>

2. <span data-ttu-id="d5342-231">在 Outlook 中，与外接程序进行交互。</span><span class="sxs-lookup"><span data-stu-id="d5342-231">In Outlook, interact with your add-in.</span></span>

3. <span data-ttu-id="d5342-232">在 Visual Studio 中命中断点时，根据需要逐步执行代码。</span><span class="sxs-lookup"><span data-stu-id="d5342-232">As breakpoints are hit in Visual Studio, step through the code as needed.</span></span>

<span data-ttu-id="d5342-233">你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭 Outlook 并重新启动该项目。</span><span class="sxs-lookup"><span data-stu-id="d5342-233">You can change your code and review the effects of those changes in your add-in without having to close Outlook and restart the project.</span></span> <span data-ttu-id="d5342-234">保存对代码的更改后，只需打开外接程序的快捷菜单（在 Outlook 中），然后选择“**重新加载**”。</span><span class="sxs-lookup"><span data-stu-id="d5342-234">After you save changes to your code, simply open the shortcut menu for the add-in (in Outlook), and then choose **Reload**.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d5342-235">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d5342-235">Next steps</span></span>

<span data-ttu-id="d5342-236">在外接程序正常工作后，请参阅[部署和发布 Office 外接程序](../publish/publish.md)，以了解可用于将外接程序分发给用户的方法。</span><span class="sxs-lookup"><span data-stu-id="d5342-236">After your add-in is working as desired, see [Deploy and publish your Office Add-in](../publish/publish.md) to learn about the ways you can distribute the add-in to users.</span></span>
