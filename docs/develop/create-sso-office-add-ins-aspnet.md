---
title: 创建使用单一登录的 ASP.NET Office 加载项
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: be7d6a8ab7f646c1ef9e77a2b459c41000c49f43
ms.sourcegitcommit: eea7f2b1679cf9a209d35880b906e311bdf1359c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/26/2018
ms.locfileid: "21241136"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="5b13d-102">创建使用单一登录的 ASP.NET Office 加载项（预览）</span><span class="sxs-lookup"><span data-stu-id="5b13d-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="5b13d-p101">如果用户已登录 Office，加载项可以使用相同的凭据，这样用户无需重新登录，即可访问多个应用。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="5b13d-105">本文将引导你完成在使用 ASP.NET、OWIN 和适用于 .NET 的 Microsoft 验证库 (MSAL) 生成的外接程序中启用单一登录 (SSO) 的过程。</span><span class="sxs-lookup"><span data-stu-id="5b13d-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="5b13d-106">有关与此类似的 Node.js 加载项文章，请参阅[创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)。</span><span class="sxs-lookup"><span data-stu-id="5b13d-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5b13d-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="5b13d-107">Prerequisites</span></span>

* <span data-ttu-id="5b13d-108">最新版 Visual Studio 2017 Preview。</span><span class="sxs-lookup"><span data-stu-id="5b13d-108">The latest available version of Visual Studio 2017 Preview.</span></span>

* <span data-ttu-id="5b13d-p102">Office 2016，版本 1708，内部版本 8424.nnnn 或更高版本（Office 365 订阅版本，有时称为“即点即用”）。可能需要成为 Office 预览体验成员才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/en-us/office-insider?tab=tab-1)。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p102">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”). You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="5b13d-112">设置初学者项目</span><span class="sxs-lookup"><span data-stu-id="5b13d-112">Set up the starter project</span></span>

1. <span data-ttu-id="5b13d-113">在 [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso) 处克隆或下载存储库。</span><span class="sxs-lookup"><span data-stu-id="5b13d-113">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="5b13d-p103">打开 **Before** 文件夹，并打开 Visual Studio 中的 .sln 文件。这是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5b13d-p104">在同一存储库中，还有此示例的已完成版本。这就像是完成本文中的过程后生成的加载项，不同之处在于已完成的项目有代码注释，但这对本文文本来说是多余的。若要使用已完成版本，只需打开 `sln` 文件，再按照本文中的说明操作即可，但要跳过**编写客户端代码**和**编写服务器端代码**这两部分。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="5b13d-p105">项目打开后，在 Visual Studio 中执行生成，这会安装 packages.config 文件中列出的包。此过程的完成耗时可能需要几秒到几分钟不等，具体视计算机本地包缓存中的包数量而定。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5b13d-122">将看到有关 Identity 命名空间的错误消息。</span><span class="sxs-lookup"><span data-stu-id="5b13d-122">You will get an error about the Identity namespace.</span></span> <span data-ttu-id="5b13d-123">这是由于将在下一步中修复的配置问题间接造成。</span><span class="sxs-lookup"><span data-stu-id="5b13d-123">This is a side effect of a configuration issue that you will fix with the next step.</span></span> <span data-ttu-id="5b13d-124">重要的是，包已安装。</span><span class="sxs-lookup"><span data-stu-id="5b13d-124">The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="5b13d-125">目前，SSO 所需的 MSAL 库 (Microsoft.Identity.Client) 版本（`1.1.1-alpha0393` 版本）没有在标准 NuGet 目录中列出，因此也没有在 package.config 中列出，必须单独进行安装。</span><span class="sxs-lookup"><span data-stu-id="5b13d-125">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.1-alpha0393`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span> 

   > 1. <span data-ttu-id="5b13d-126">在“工具”\*\*\*\* 菜单上，依次转到“NuGet 包管理器”\*\*\*\* > “包管理器控制台”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5b13d-126">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span> 

   > 2. <span data-ttu-id="5b13d-p107">在控制台上，运行以下命令。即使使用快速的 Internet 连接，也可能需要一分钟或更长时间才能完成。完成后你应该看到 **成功安装“Microsoft.Identity.Client 1.1.4-preview0002”......** 接近控制台输出的末尾。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p107">At the console, run the following command. It may take a minute or more to complete even with a fast Internet connection. When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.1-alpha0393' ...** near the end of the output in the console.</span></span>

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`

   > 3. <span data-ttu-id="5b13d-p108">在“解决方案资源管理器”\*\*\*\* 中，右键单击“引用”\*\*\*\*。验证是否列出了“Microsoft.Identity.Client”\*\*\*\*。如果没有列出或它的条目上有警告图标，请先删除此条目，再使用“Visual Studio 添加引用”向导，添加对 **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-alpha0002\lib\net45\Microsoft.Identity.Client.dll** 处程序集的引用。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p108">In **Solution Explorer**, right-click **References**. Verify that **Microsoft.Identity.Client** is listed. If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="5b13d-133">重新生成此项目。</span><span class="sxs-lookup"><span data-stu-id="5b13d-133">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="5b13d-134">向 Azure AD v2.0 终结点注册加载项</span><span class="sxs-lookup"><span data-stu-id="5b13d-134">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="5b13d-135">以下说明以通用方式书写，以便可以在多个地方使用。</span><span class="sxs-lookup"><span data-stu-id="5b13d-135">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="5b13d-136">对于本文而言，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5b13d-136">For this ariticle do the following:</span></span>
- <span data-ttu-id="5b13d-137">将占位符 **$ADD-IN-NAME$** 替换为 `Office-Add-in-ASPNET-SSO`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-137">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="5b13d-138">将占位符 **$FQDN-WITHOUT-PROTOCOL$** 替换为 `localhost:44355`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-138">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="5b13d-139">在 **“选择权限”** 对话框中指定权限时，选中以下权限框。</span><span class="sxs-lookup"><span data-stu-id="5b13d-139">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="5b13d-140">加载项本身真正需要的只是第一项权限，但服务器端代码使用的 MSAL 库需要有 `offline_access` 和 `openid`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-140">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="5b13d-141">Office 主机必须有 `profile` 权限，才能获取对加载项 Web 应用程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="5b13d-141">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="5b13d-142">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="5b13d-142">Files.Read.All</span></span>
    * <span data-ttu-id="5b13d-143">offline_access</span><span class="sxs-lookup"><span data-stu-id="5b13d-143">offline_access</span></span>
    * <span data-ttu-id="5b13d-144">openid</span><span class="sxs-lookup"><span data-stu-id="5b13d-144">openid</span></span>
    * <span data-ttu-id="5b13d-145">配置文件</span><span class="sxs-lookup"><span data-stu-id="5b13d-145">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="5b13d-146">向外接程序授予管理员许可</span><span class="sxs-lookup"><span data-stu-id="5b13d-146">Details are at: Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="5b13d-147">配置外接程序</span><span class="sxs-lookup"><span data-stu-id="5b13d-147">Configure the add-in</span></span>

1. <span data-ttu-id="5b13d-148">在以下字符串中，将占位符“{tenant_ID}”替换为你的 Office 365 租户ID。</span><span class="sxs-lookup"><span data-stu-id="5b13d-148">In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenant ID.</span></span> <span data-ttu-id="5b13d-149">使用[“找到你的 Office 365 租户 ID”](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b)中的一种方法获得它。</span><span class="sxs-lookup"><span data-stu-id="5b13d-149">Use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. <span data-ttu-id="5b13d-150">在 Visual Studio 中打开 web.config。你需要为 **appSettings** 部分中的某些键分配值。</span><span class="sxs-lookup"><span data-stu-id="5b13d-150">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

3. <span data-ttu-id="5b13d-p112">将在步骤 1 中构造的字符串用作名为“ida:Issuer”的键的值。请确保此值中没有空格。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p112">Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.</span></span>

4. <span data-ttu-id="5b13d-153">将下面的值分配给相应的键：</span><span class="sxs-lookup"><span data-stu-id="5b13d-153">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="5b13d-154">键</span><span class="sxs-lookup"><span data-stu-id="5b13d-154">Key</span></span>|<span data-ttu-id="5b13d-155">值</span><span class="sxs-lookup"><span data-stu-id="5b13d-155">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="5b13d-156">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="5b13d-156">ida:ClientID</span></span>|<span data-ttu-id="5b13d-157">注册外接程序时获取的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="5b13d-157">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="5b13d-158">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="5b13d-158">ida:Audience</span></span>|<span data-ttu-id="5b13d-159">注册外接程序时获取的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="5b13d-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="5b13d-160">ida:Password</span><span class="sxs-lookup"><span data-stu-id="5b13d-160">ida:Password</span></span>|<span data-ttu-id="5b13d-161">注册加载项时获取的密码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-161">TThe password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="5b13d-p113">下面的示例展示了四个键的更改后效果。*请注意，ClientID 和 Audience 是相同的*。也可以将一个键同时用于这两种用途，但如果继续单独使用两个键，web.config 标记的可重用性将更高，因为它们并非始终相同。此外，单独使用两个键也可以强化以下概念：相对于 Office 主机而言，加载项是 OAuth 资源；相对于 Microsoft Graph 而言，它同时又是 OAuth 客户端。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > <span data-ttu-id="5b13d-166">**appSettings** 部分中的其他设置保持不变。</span><span class="sxs-lookup"><span data-stu-id="5b13d-166">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="5b13d-167">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-167">Save and close the file.</span></span>

1. <span data-ttu-id="5b13d-168">在外接程序项目中，打开外接程序清单文件“Office-Add-in-ASPNET-SSO.xml”。</span><span class="sxs-lookup"><span data-stu-id="5b13d-168">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span>

1. <span data-ttu-id="5b13d-169">滚动到文件底部。</span><span class="sxs-lookup"><span data-stu-id="5b13d-169">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="5b13d-170">结束 `</VersionOverrides>` 标记的正上方有以下标记：</span><span class="sxs-lookup"><span data-stu-id="5b13d-170">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="5b13d-171">将标记中的*两处*占位符“{application_GUID here}”均替换成在注册加载项时复制的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="5b13d-171">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="5b13d-172">由于 ID 并不包含“{}”，因此请勿添加它们。</span><span class="sxs-lookup"><span data-stu-id="5b13d-172">The "{}" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="5b13d-173">这与在 web.config 中对 ClientID 和 Audience 使用的 ID 相同。</span><span class="sxs-lookup"><span data-stu-id="5b13d-173">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="5b13d-174">**Resource** 值是向注册的加载项添加 Web API 平台时设置的**应用 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="5b13d-174">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="5b13d-175">仅在通过 AppSource 销售加载项时，才使用 **Scopes** 部分生成许可对话框。</span><span class="sxs-lookup"><span data-stu-id="5b13d-175">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="5b13d-176">在 Visual Studio 中，打开“错误列表”\*\*\*\* 的“警告”\*\*\*\* 选项卡。</span><span class="sxs-lookup"><span data-stu-id="5b13d-176">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="5b13d-177">如果存在关于 `<WebApplicationInfo>` 不是 `<VersionOverrides>` 的有效子级的警告，则该 Visual Studio 2017 Preview 版本无法识别 SSO 标记。</span><span class="sxs-lookup"><span data-stu-id="5b13d-177">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not  recognize the SSO markup.</span></span> <span data-ttu-id="5b13d-178">作为解决方法，请对 Word、Excel 或 PowerPoint 外接程序执行以下操作。</span><span class="sxs-lookup"><span data-stu-id="5b13d-178">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="5b13d-179">（如果使用的是 Outlook 外接程序，请参阅下面的解决方法。）</span><span class="sxs-lookup"><span data-stu-id="5b13d-179">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="5b13d-180">**Word、Excel 和 Powerpoint 的解决方法**</span><span class="sxs-lookup"><span data-stu-id="5b13d-180">**Workaround for Word, Excel, and Powerpoint**</span></span>

        1. <span data-ttu-id="5b13d-181">在结束 `</VersionOverrides>` 标记正上方的清单中，注释掉 `<WebApplicationInfo>` 部分。</span><span class="sxs-lookup"><span data-stu-id="5b13d-181">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="5b13d-p116">按 F5 启动调试会话。此操作会在下列文件夹（相比 Visual Studio，在“文件资源管理器”\*\*\*\* 中访问此文件夹更方便）中创建清单副本： `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="5b13d-p116">Press F5 to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="5b13d-184">在清单副本中，删除 `<WebApplicationInfo>` 部分周围的注释语法。</span><span class="sxs-lookup"><span data-stu-id="5b13d-184">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="5b13d-185">保存此清单副本。</span><span class="sxs-lookup"><span data-stu-id="5b13d-185">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="5b13d-p117">现在，必须阻止 Visual Studio 在用户下次按 F5 时重写此清单副本。右键单击“解决方案资源管理器”\*\*\*\* 顶部的解决方案节点（而不是任何项目节点）。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="5b13d-188">选择上下文菜单中的“属性”\*\*\*\*，随后“解决方案属性页”\*\*\*\* 对话框打开。</span><span class="sxs-lookup"><span data-stu-id="5b13d-188">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="5b13d-189">展开“配置属性”\*\*\*\*，并选择“配置”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5b13d-189">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="5b13d-190">在 **Office-Add-in-ASPNET-SSO** 项目（*不是* **Office-Add-in-ASPNET-SSO-WebAPI** 项目）行中取消选择“生成”\*\*\*\* 和“部署”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5b13d-190">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="5b13d-191">按“确定”\*\*\*\* 关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="5b13d-191">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="5b13d-192">**Outlook 的解决方法**</span><span class="sxs-lookup"><span data-stu-id="5b13d-192">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="5b13d-193">在开发计算机上找到现有的 `MailAppVersionOverridesV1_1.xsd`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-193">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`.</span></span> <span data-ttu-id="5b13d-194">它应位于 `./Xml/Schemas/{lcid}` 下的 Visual Studio 安装目录中。</span><span class="sxs-lookup"><span data-stu-id="5b13d-194">This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`.</span></span> <span data-ttu-id="5b13d-195">例如，在英语（美国）的系统上进行 VS 2017 32 位的典型安装时，完整路径为 `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-195">For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="5b13d-196">将现有文件重命名为 `MailAppVersionOverridesV1_1.old`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-196">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="5b13d-197">将此修改后的文件版本复制到文件夹中：[修改后的 MailAppVersionOverrides 架构](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="5b13d-197">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="5b13d-198">在 Visual Studio 中保存并关闭该主清单文件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-198">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="5b13d-199">编写客户端代码</span><span class="sxs-lookup"><span data-stu-id="5b13d-199">Code the client side</span></span>

1. <span data-ttu-id="5b13d-p119">打开 **Scripts** 文件夹中的 Home.js 文件。其中已存在一些代码：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="5b13d-202">针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-202">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="5b13d-203">方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。`showResult`</span><span class="sxs-lookup"><span data-stu-id="5b13d-203">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="5b13d-204">方法，用于记录最终用户不应看到的控制台错误。`logErrors`</span><span class="sxs-lookup"><span data-stu-id="5b13d-204">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="5b13d-p120">在向 `Office.initialize` 分配函数下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="5b13d-207">加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="5b13d-207">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="5b13d-208">计数器变量 `timesGetOneDriveFilesHasRun` 和标志变量 `triedWithoutForceConsent` 用于确保用户不会重复循环失败的尝试来获取令牌。</span><span class="sxs-lookup"><span data-stu-id="5b13d-208">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="5b13d-p122">虽然 `getDataWithToken` 方法是在下一步中创建，但请注意，它会将 `forceConsent` 选项设置为 `false`。有关详细信息，请参阅下一步。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="5b13d-p123">在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="5b13d-p124">是 Office.js 中的新 API，可便于加载项要求 Office 主机应用（Excel、PowerPoint、Word 等）提供加载项访问令牌（对于已登录 Office 的用户）。反过来，Office 主机应用会向 Azure AD 2.0 终结点请求获取令牌。由于已在注册加载项时将 Office 主机预授权给加载项，因此 Azure AD 会发送访问令牌。`getAccessTokenAsync`</span><span class="sxs-lookup"><span data-stu-id="5b13d-p124">The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="5b13d-216">如果用户未登录 Office，Office 主机会提示用户登录。</span><span class="sxs-lookup"><span data-stu-id="5b13d-216">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="5b13d-217">options 参数将 `forceConsent` 设置为 `false`，因此用户不会在每次使用加载项时都看到提示，要求其许可向 Office 主机授予对加载项的访问权限。</span><span class="sxs-lookup"><span data-stu-id="5b13d-217">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="5b13d-218">用户首次运行加载项时，`getAccessTokenAsync` 调用会失败，但在后续步骤中添加的错误处理逻辑会自动重新调用（`forceConsent` 选项设置为 `true`），并提示用户许可，但仅限首次运行。</span><span class="sxs-lookup"><span data-stu-id="5b13d-218">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="5b13d-219">方法将在后续步骤中创建。`handleClientSideErrors`</span><span class="sxs-lookup"><span data-stu-id="5b13d-219">You will create the `handleClientSideErrors` method in a later step.</span></span>

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. <span data-ttu-id="5b13d-p126">用以下行替换 TODO1。可以在后续步骤中创建 `getData` 方法和服务器端“/api/values”路由。相对 URL 用于终结点，因为它必须与外接程序托管在相同的域中。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="5b13d-p127">在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="5b13d-p128">此方法调用指定 Web API 终结点，并向它传递访问令牌，这也是 Office 主机应用用于获取对加载项的访问权限的令牌。在服务器端，此访问令牌将用于“代表”流，以获取对 Microsoft Graph 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="5b13d-227">方法将在后续步骤中创建。`handleServerSideErrors`</span><span class="sxs-lookup"><span data-stu-id="5b13d-227">You will create the `handleServerSideErrors` method in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="5b13d-228">创建错误处理方法</span><span class="sxs-lookup"><span data-stu-id="5b13d-228">Create the error-handling methods</span></span>

1. <span data-ttu-id="5b13d-229">在 `getData` 方法下方，添加下列方法。</span><span class="sxs-lookup"><span data-stu-id="5b13d-229">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="5b13d-230">当 Office 主机无法获取对加载项 Web 服务的访问令牌时，此方法便会处理加载项客户端中的错误。</span><span class="sxs-lookup"><span data-stu-id="5b13d-230">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="5b13d-231">这些错误通过错误代码进行报告，因此下面的方法使用 `switch` 语句区分它们。</span><span class="sxs-lookup"><span data-stu-id="5b13d-231">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="5b13d-232">将 `TODO2` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-232">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="5b13d-233">如果用户未登录或用户取消（未响应）提供辅助身份验证因素的提示，错误 13001 发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-233">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="5b13d-234">无论属于上述哪种情况，代码都会重新运行 `getDataWithToken` 方法，并设置强制登录提示选项。</span><span class="sxs-lookup"><span data-stu-id="5b13d-234">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="5b13d-235">将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-235">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="5b13d-236">如果用户登录或许可被中止，错误 13002 发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-236">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="5b13d-237">建议用户重试一次，但只能重试一次。</span><span class="sxs-lookup"><span data-stu-id="5b13d-237">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="5b13d-238">将 `TODO4` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-238">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="5b13d-239">如果用户用于登录的帐户既不是工作帐户或学校帐户，也不是 Microsoft 帐户，错误 13003 发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-239">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Micrososoft Account.</span></span> <span data-ttu-id="5b13d-240">建议用户注销，再使用受支持的帐户类型重新登录。</span><span class="sxs-lookup"><span data-stu-id="5b13d-240">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="5b13d-241">此方法不处理错误 13004 和 13005，因为它们只在开发期间出现。</span><span class="sxs-lookup"><span data-stu-id="5b13d-241">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="5b13d-242">无法通过运行时代码进行修复，并且向最终用户报告这两个错误也没有意义。</span><span class="sxs-lookup"><span data-stu-id="5b13d-242">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="5b13d-p134">将 `TODO5` 替换为下列代码。如果 Office 主机中出现可能表明主机处于不稳定状态的未指定错误，就会发生错误 13006。建议用户重启 Office。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p134">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="5b13d-246">将 `TODO6` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-246">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="5b13d-247">如果 Office 主机与 AAD 之间的交互出现问题，导致主机无法获得对加载项 Web 服务/应用的访问令牌，错误 13007 发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-247">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="5b13d-248">这可能由于暂时网络问题所致。</span><span class="sxs-lookup"><span data-stu-id="5b13d-248">This may be a temporary network issue.</span></span> <span data-ttu-id="5b13d-249">建议用户稍后重试。</span><span class="sxs-lookup"><span data-stu-id="5b13d-249">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="5b13d-p136">将 `TODO7` 替换为下列代码。如果用户触发的操作未等到上一次调用完成就调用了 `getAccessTokenAsync`，就会发生错误 13008。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p136">Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="5b13d-252">将 `TODO8` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-252">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="5b13d-253">如果加载项不支持强制许可，但调用 `getAccessTokenAsync` 时 `forceConsent` 选项设置为 `true`，错误 13009 发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-253">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="5b13d-254">通常情况下，如果发生这种情况，代码应自动重新运行 `getAccessTokenAsync`，同时将许可选项设置为 `false`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-254">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="5b13d-255">不过，在某些情况下，调用将 `forceConsent` 设置为 `true` 的方法本身就是在自动响应调用将选项设置为 `false` 的方法时出现的错误。</span><span class="sxs-lookup"><span data-stu-id="5b13d-255">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="5b13d-256">此时，不得重试代码，而是应建议用户注销并重新登录。</span><span class="sxs-lookup"><span data-stu-id="5b13d-256">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="5b13d-257">将 `TODO9` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-257">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="5b13d-p138">在 `handleClientSideErrors` 方法下方，添加下列方法。此方法可处理加载项 Web 服务中发生的以下错误：无法执行代表流，或无法从 Microsoft Graph 获取数据。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p138">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO14: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. <span data-ttu-id="5b13d-260">将 `TODO10` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-260">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="5b13d-261">请注意，对于加载项 Web 服务传递给加载项客户端的大多数 `4xx` 错误，响应中都有 **ExceptionMessage** 属性，其中包含 AADSTS（Azure Active Directory 安全令牌服务）错误号和其他数据。</span><span class="sxs-lookup"><span data-stu-id="5b13d-261">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="5b13d-262">不过，如果 AAD 向加载项 Web 服务发送消息，请求提供其他身份验证因素，那么消息包含特殊的 **Claims** 属性，用于指定（使用代码编号）需要的其他因素。</span><span class="sxs-lookup"><span data-stu-id="5b13d-262">However, when AAD sends a message to the add-in's web service asking for an additonal authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="5b13d-263">由于创建并向客户端发送 HTTP Response 的 ASP.NET API 并不知道此 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。</span><span class="sxs-lookup"><span data-stu-id="5b13d-263">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="5b13d-264">将在后续步骤中创建的服务器端代码负责处理这个问题，具体方法是手动向 Response 对象添加 **Claims** 值。</span><span class="sxs-lookup"><span data-stu-id="5b13d-264">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="5b13d-265">因为此值位于 **Message** 属性中，所以代码也需要解析相应属性。</span><span class="sxs-lookup"><span data-stu-id="5b13d-265">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="5b13d-p140">将 `TODO11` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p140">Replace `TODO11` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-268">如果 Microsoft Graph 要求进行其他形式的身份验证，错误 50076 发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-268">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="5b13d-269">Office 主机应获取新令牌（使用 **Claims** 值作为 `authChallenge` 选项）。</span><span class="sxs-lookup"><span data-stu-id="5b13d-269">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="5b13d-270">这就指示 AAD 提示用户进行所有必需形式的身份验证。</span><span class="sxs-lookup"><span data-stu-id="5b13d-270">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. <span data-ttu-id="5b13d-p142">将 `TODO12` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p142">Replace `TODO12` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-273">错误 65001 表示未许可授予（或已撤消）一个或多个对 Microsoft Graph 的访问权限。</span><span class="sxs-lookup"><span data-stu-id="5b13d-273">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="5b13d-274">加载项应获取新令牌（`forceConsent` 选项设置为 `true`）。</span><span class="sxs-lookup"><span data-stu-id="5b13d-274">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. <span data-ttu-id="5b13d-p143">将 `TODO13` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p143">Replace `TODO13` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-p144">错误 70011 有多重含义。对于此加载项而言，最重要的含义是已请求获取的范围（权限）无效。因此，代码会检查是否有完整错误说明，而不仅仅是数字。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p144">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="5b13d-279">加载项应报告此错误。</span><span class="sxs-lookup"><span data-stu-id="5b13d-279">The add-in should report the error.</span></span>

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. <span data-ttu-id="5b13d-p145">将 `TODO14` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p145">Replace `TODO14` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-282">如果 `access_as_user` 范围（权限）不在访问令牌中，此令牌由加载项客户端发送到 AAD 以便在代表流中使用，那么在后续步骤中创建的服务器端代码将发送消息 `Missing access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-282">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="5b13d-283">加载项应报告此错误。</span><span class="sxs-lookup"><span data-stu-id="5b13d-283">The add-in should report the error.</span></span>

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. <span data-ttu-id="5b13d-p146">将 `TODO15` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p146">Replace `TODO15` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-p147">要在服务器端代码中使用的标识库（Microsoft 身份验证库 (MSAL)）应确保没有向 Microsoft Graph 发送任何到期或无效令牌；但如果这种情况确实发生，从 Microsoft Graph 返回到加载项 Web 服务的错误包含代码 `InvalidAuthenticationToken`。在后续步骤中创建的服务器端代码会将此消息中继到加载项客户端。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p147">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`. Server-side code you will create in a latter step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="5b13d-288">在这种情况下，加载项应重置计数器和标志变量，再重新调用按钮处理程序方法，以从头开始执行整个身份验证流程。</span><span class="sxs-lookup"><span data-stu-id="5b13d-288">In this case, the add-in should start the entire authentication process over by resetting the counter and flag varibles, and then re-calling the button handler method.</span></span>

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. <span data-ttu-id="5b13d-289">将 `TODO16` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-289">Replace `TODO16` with the following code.</span></span>

    ```javascript
    else {
        logError(result);
    }    
    ```

1. <span data-ttu-id="5b13d-290">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-290">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="5b13d-291">编写服务器端代码</span><span class="sxs-lookup"><span data-stu-id="5b13d-291">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="5b13d-292">配置 OWIN 中间件</span><span class="sxs-lookup"><span data-stu-id="5b13d-292">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="5b13d-293">在项目的根目录中打开 Startup.cs 文件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-293">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="5b13d-p148">将关键字 `partial` 添加到 Startup 类（如果其中尚不存在该关键字）的声明。具体应如下所示：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p148">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="5b13d-p149">将以下行添加至 `Configuration` 方法的正文。在后续步骤中创建 `ConfigureAuth` 方法。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p149">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="5b13d-298">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-298">Save and close the file.</span></span>

1. <span data-ttu-id="5b13d-299">右键单击“App_Start”\*\*\*\* 文件夹，并依次选择“添加”>“类”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5b13d-299">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="5b13d-300">在“添加新项”\*\*\*\* 对话框中，命名文件“Startup.Auth.cs”\*\*\*\*，再单击“添加”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5b13d-300">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="5b13d-301">将新文件中的命名空间名称缩短为 `Office_Add_in_ASPNET_SSO_WebAPI`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-301">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="5b13d-302">确保下列所有 `using` 语句都位于文件的顶部。</span><span class="sxs-lookup"><span data-stu-id="5b13d-302">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="5b13d-p150">将关键字 `partial` 添加到 `Startup` 类（如果其中尚不存在该关键字）的声明。具体应如下所示：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p150">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="5b13d-p151">将下列方法添加到 `Startup` 类。该方法指定 OWIN 中间件如何验证从客户端 Home.js 文件的 `getData` 方法传递给它的访问令牌。每次调用使用 `[Authorize]` 属性修饰的 Web API 终结点时都会触发授权过程。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p151">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="5b13d-308">将 TODO3 替换为以下代码行。</span><span class="sxs-lookup"><span data-stu-id="5b13d-308">Replace the TODO3 with the following.</span></span> <span data-ttu-id="5b13d-309">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-309">Note about this code:</span></span>

    * <span data-ttu-id="5b13d-310">代码指示 OWIN 以确保在来自 Office 主机（并通过客户端调用 `getData` 进行传递）的访问令牌中指定的受众和令牌颁发者必须与 web.config 中指定的值相匹配。</span><span class="sxs-lookup"><span data-stu-id="5b13d-310">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="5b13d-p153">将 `SaveSigninToken` 设置为 `true` 将导致 OWIN 从 Office 主机保存原始令牌。外接程序需要它来获取具有“代表”流的 Microsoft Graph 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p153">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.</span></span>
    * <span data-ttu-id="5b13d-p154">OWIN 中间件不验证范围。应包括 `access_as_user` 的访问令牌范围是在控制器中进行验证。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p154">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="5b13d-p155">将 TODO4 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p155">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-317">调用的是方法 `UseOAuthBearerAuthentication`，而不是更常见的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因为后者与 Azure AD V2 终结点不兼容。</span><span class="sxs-lookup"><span data-stu-id="5b13d-317">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="5b13d-318">传递给该方法的发现 URL 是 OWIN 中间件获得用于获取所需密钥说明的位置，以验证从 Office 主机接收到的访问令牌上的签名。</span><span class="sxs-lookup"><span data-stu-id="5b13d-318">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="5b13d-319">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="5b13d-319">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="5b13d-320">创建 /api/values 控制器</span><span class="sxs-lookup"><span data-stu-id="5b13d-320">Create the /api/values controller</span></span>

1. <span data-ttu-id="5b13d-321">打开文件 **Controllers\ValueController.cs**。</span><span class="sxs-lookup"><span data-stu-id="5b13d-321">Open the file **Controllers\ValueController.cs**.</span></span>

2. <span data-ttu-id="5b13d-322">请确保下列 `using` 语句位于文件顶部。</span><span class="sxs-lookup"><span data-stu-id="5b13d-322">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. <span data-ttu-id="5b13d-p156">在声明 `ValuesController` 的代码行的正上方，添加属性 `[Authorize]`。这可确保只要调用控制器方法时，加载项就会运行在上一过程中配置的授权过程。只有拥有对加载项的有效访问令牌，调用方才能调用控制器的方法。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p156">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5b13d-326">生产 ASP.NET MVC Web API 服务应在一个或多个自定义 [FilterAttribute](https://msdn.microsoft.com/en-us/library/system.web.http.filters(v=vs.108).aspx) 类中有代表流的自定义逻辑。</span><span class="sxs-lookup"><span data-stu-id="5b13d-326">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom [FilterAttribute](https://msdn.microsoft.com/en-us/library/system.web.http.filters(v=vs.108).aspx) classes.</span></span> <span data-ttu-id="5b13d-327">此说明性示例将逻辑放入主控制器中，以便能够轻松跟进授权和数据提取逻辑的整个流。</span><span class="sxs-lookup"><span data-stu-id="5b13d-327">This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed.</span></span> <span data-ttu-id="5b13d-328">这也可以让示例与 [Azure 示例](https://github.com/Azure-Samples/)中的授权示例模式保持一致。</span><span class="sxs-lookup"><span data-stu-id="5b13d-328">This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>    

4. <span data-ttu-id="5b13d-329">将下列方法添加到 `ValuesController`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-329">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="5b13d-330">请注意，返回值是 `Task<HttpResponseMessage>`（而不是 `Task<IEnumerable<string>>`），这对于 `GET api/values` 方法而言更为常见。</span><span class="sxs-lookup"><span data-stu-id="5b13d-330">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="5b13d-331">这是将自定义授权逻辑放入控制器中造成不良影响：此逻辑中的一些错误条件要求将 HTTP Response 对象发送到加载项客户端。</span><span class="sxs-lookup"><span data-stu-id="5b13d-331">This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span> 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. <span data-ttu-id="5b13d-332">将 `TODO1` 替换为以下代码行，以验证令牌中指定的范围是否包括 `access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-332">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > <span data-ttu-id="5b13d-p159">只能使用 `access_as_user` 范围授权 API 为 Office 加载项处理代表流。服务中的其他 API 应有自己的范围要求。这就限制了使用 Office 获得的令牌可以访问的内容。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p159">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.</span></span>

6. <span data-ttu-id="5b13d-p160">将 `TODO2` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p160">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="5b13d-337">它将从 Office 主机收到的原始访问令牌转换为，传递给另一个方法的 `UserAssertion` 对象。</span><span class="sxs-lookup"><span data-stu-id="5b13d-337">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="5b13d-p161">外接程序不再扮演 Office 主机和用户需要访问的资源（或受众）的角色。现在它本身就是一个需要访问 Microsoft Graph 的客户端。`ConfidentialClientApplication` 是 MSAL“客户端上下文”对象。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p161">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="5b13d-p162">构造函数的第三个参数是在“代表”流中实际不使用的重定向 URL，但使用正确的 URL 是一个很好的做法。第四和第五个参数可用于定义持久性存储，该存储使得外接程序能在不同的会话之间重用未过期的令牌。此示例不实现任何持久性存储。`ConfidentialClientApplication`</span><span class="sxs-lookup"><span data-stu-id="5b13d-p162">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="5b13d-344">MSAL 要求 `openid`、`offline_access` 作用域能够发挥作用，但如果代码过多地发出请求，则会抛出错误。</span><span class="sxs-lookup"><span data-stu-id="5b13d-344">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="5b13d-345">如果代码请求获取 `profile`，也会抛出错误，这真正仅适用于 Office 主机应用程序获取对加载项 Web 应用程序的令牌时。</span><span class="sxs-lookup"><span data-stu-id="5b13d-345">It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application.</span></span> <span data-ttu-id="5b13d-346">因此，只会显式请求获取 `Files.Read.All`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-346">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. <span data-ttu-id="5b13d-p164">将 `TODO3` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p164">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-p165">方法将首先查找内存中的 MSAL 缓存，获取匹配的访问令牌。仅当不存在任何令牌时，该方法才会通过 Azure AD V2 终结点启动“代表”流。`ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync`</span><span class="sxs-lookup"><span data-stu-id="5b13d-p165">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="5b13d-351">如果 MS Graph 资源要求进行多重身份验证，但用户尚未提供，AAD 就会抛出包含 Claims 属性的异常。</span><span class="sxs-lookup"><span data-stu-id="5b13d-351">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="5b13d-p166">必须将 Claims 属性值传递到客户端，接着客户端会将它传递到 Office 主机，然后主机会将它添加到新令牌请求中。AAD 将提示用户进行所有必需形式的身份验证。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p166">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="5b13d-354">任何不属于类型 `MsalServiceException` 的异常都是有意不捕获的，这样才能作为 `500 Server Error` 消息传播到客户端。</span><span class="sxs-lookup"><span data-stu-id="5b13d-354">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. <span data-ttu-id="5b13d-p167">将 `TODO3a` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p167">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-p168">如果 MS Graph 资源要求进行多重身份验证，但用户尚未提供，AAD 就会返回包含错误 AADSTS50076 和 **Claims** 属性的“400 错误请求”。MSAL 会抛出包含此信息的 **MsalUiRequiredException**（继承自 **MsalServiceException**）。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p168">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="5b13d-p169">必须将 **Claims** 属性值传递到客户端，接着客户端应将它传递到 Office 主机，然后主机会将它添加到新令牌请求中。AAD 会提示用户进行所有必需形式的身份验证。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p169">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="5b13d-361">由于创建异常 HTTP Response 的 API 并不知道 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。</span><span class="sxs-lookup"><span data-stu-id="5b13d-361">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object.</span></span> <span data-ttu-id="5b13d-362">必须手动创建消息来添加它。</span><span class="sxs-lookup"><span data-stu-id="5b13d-362">We have to manually create a message that includes it.</span></span> <span data-ttu-id="5b13d-363">不过，自定义 **Message** 属性会阻止创建 **ExceptionMessage** 属性，因此向客户端发送错误 ID `AADSTS50076` 的唯一方法是，将它添加到自定义 **Message** 中。</span><span class="sxs-lookup"><span data-stu-id="5b13d-363">A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**.</span></span> <span data-ttu-id="5b13d-364">客户端中的 JavaScript 需要发现响应是否包含 **Message** 或 **ExceptionMessage**，这样才能了解要读取的内容。</span><span class="sxs-lookup"><span data-stu-id="5b13d-364">JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="5b13d-365">自定义消息被格式化为 JSON，以便客户端 JavaScript 能够使用已知的 `JSON` 对象方法分析它。</span><span class="sxs-lookup"><span data-stu-id="5b13d-365">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="5b13d-366">方法将在后续步骤中创建。`SendErrorToClient`</span><span class="sxs-lookup"><span data-stu-id="5b13d-366">You will create the `SendErrorToClient` method in a later step.</span></span> <span data-ttu-id="5b13d-367">它的第二个参数是 **Exception** 对象。</span><span class="sxs-lookup"><span data-stu-id="5b13d-367">It's second parameter is an **Exception** object.</span></span> <span data-ttu-id="5b13d-368">在此示例中，代码传递 `null`，因为添加 **Exception** 对象会阻止在生成的 HTTP Response 中添加 **Message** 属性。</span><span class="sxs-lookup"><span data-stu-id="5b13d-368">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. <span data-ttu-id="5b13d-p172">将 `TODO3b` 和 `TODO3c` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p172">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-371">如果 AAD 调用包含至少一个范围（权限）未获用户和租户管理员的许可（或许可被撤消），</span><span class="sxs-lookup"><span data-stu-id="5b13d-371">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked).</span></span> <span data-ttu-id="5b13d-372">AAD 返回“400 错误请求”和错误 `AADSTS65001`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-372">AAD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="5b13d-373">MSAL 抛出包含此信息的 **MsalUiRequiredException**。</span><span class="sxs-lookup"><span data-stu-id="5b13d-373">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="5b13d-374">客户端应通过选项 `{ forceConsent: true }` 重新调用 `getAccessTokenAsync`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-374">The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="5b13d-375">如果 AAD 调用包含至少一个 AAD 无法识别的范围，AAD 返回“400 错误请求”和错误 `AADSTS70011`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-375">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="5b13d-376">MSAL 抛出包含此信息的 **MsalUiRequiredException**。</span><span class="sxs-lookup"><span data-stu-id="5b13d-376">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="5b13d-377">客户端应通知用户。</span><span class="sxs-lookup"><span data-stu-id="5b13d-377">The client should inform the user.</span></span>
    *  <span data-ttu-id="5b13d-378">包含完整说明，因为 70011 也会在其他情况下返回，只有在表示存在无效范围时，才需要在此加载项中处理它。</span><span class="sxs-lookup"><span data-stu-id="5b13d-378">The entire description is included beause 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span> 
    *  <span data-ttu-id="5b13d-p175">**MsalUiRequiredException** 对象传递给 `SendErrorToClient`。这样可确保 HTTP 响应中有包含错误消息的 **ExceptionMessage** 属性。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p175">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="5b13d-381">由于没有自定义消息，因此会对第三个参数传递 `null`。</span><span class="sxs-lookup"><span data-stu-id="5b13d-381">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. <span data-ttu-id="5b13d-382">将 `TODO3d` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-382">Replace `TODO3d` with the following code.</span></span> <span data-ttu-id="5b13d-383">请注意，代码会重新抛出异常，而不是在包含 **HttpStatusCode.Forbidden** (401) 的自定义 HTTP Response 内中继它。</span><span class="sxs-lookup"><span data-stu-id="5b13d-383">Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401).</span></span> <span data-ttu-id="5b13d-384">结果就是，ASP.NET 发送自己的 HTTP Response，其中包含“500 服务器错误”状态。</span><span class="sxs-lookup"><span data-stu-id="5b13d-384">The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. <span data-ttu-id="5b13d-p177">将 `TODO4` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p177">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="5b13d-p178">和 `ODataHelper` 类在 **Helpers** 文件夹的文件中定义。`OneDriveItem` 类在 **Models** 文件夹的一个文件中定义。 这些类的详细讨论内容与授权或 SSO 无关，因此不在本文的讨论范围内。`GraphApiHelper`</span><span class="sxs-lookup"><span data-stu-id="5b13d-p178">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="5b13d-390">通过只请求 Microsoft Graph 提供实际所需数据，可以提升性能，因此代码使用 ` $select` 查询参数来指定仅需要 name 属性，并使用 `$top` 参数来指定仅需要前 3 个文件夹或文件名。</span><span class="sxs-lookup"><span data-stu-id="5b13d-390">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="5b13d-391">如果发送到 Microsoft Graph 的令牌无效，Microsoft Graph 会发送“401 未授权”错误和“InvalidAuthenticationToken”代码。</span><span class="sxs-lookup"><span data-stu-id="5b13d-391">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken".</span></span> <span data-ttu-id="5b13d-392">然后，ASP.NET 抛出 **RuntimeBinderException**。</span><span class="sxs-lookup"><span data-stu-id="5b13d-392">ASP.NET then throws a **RuntimeBinderException**.</span></span> <span data-ttu-id="5b13d-393">这也是当令牌到期时发生的情况，尽管 MSAL 应阻止这种情况发生。</span><span class="sxs-lookup"><span data-stu-id="5b13d-393">This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. <span data-ttu-id="5b13d-p180">将 `TODO5` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-p180">Replace `TODO5` with the following. Note about this code:</span></span> 

    * <span data-ttu-id="5b13d-p181">尽管上述代码仅请求获取 OneDrive 项的 *name* 属性，但 Microsoft Graph 始终包括 OneDrive 项的 *eTag* 属性。为减少发送到客户端的有效负载，下面的代码重新构造了仅包含项名称的结果。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p181">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="5b13d-398">包含三个 OneDrive 文件和文件夹的列表作为“200 OK”HTTP Response 发送到客户端。</span><span class="sxs-lookup"><span data-stu-id="5b13d-398">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. <span data-ttu-id="5b13d-399">在 Get 方法下方，添加下列方法。</span><span class="sxs-lookup"><span data-stu-id="5b13d-399">Below the Get method, add the following method.</span></span> <span data-ttu-id="5b13d-400">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5b13d-400">About this code note:</span></span>  

    * <span data-ttu-id="5b13d-401">此方法将服务器端异常信息中继到客户端。</span><span class="sxs-lookup"><span data-stu-id="5b13d-401">The method relays to the client information about a server-side exception.</span></span> 
    * <span data-ttu-id="5b13d-402">如果将原始异常传递到此方法，那么 HttpError 构造函数会在 **ExceptionMessage** 属性中添加来自 Exception 对象的信息。</span><span class="sxs-lookup"><span data-stu-id="5b13d-402">If the original exception is passed to the method, then the HttpError constuctor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="5b13d-403">如果对异常传递了 `null`，那么 HttpError 构造函数会在 **Message** 属性中添加 message 参数，且 **ExceptionMessage** 属性不存在。</span><span class="sxs-lookup"><span data-stu-id="5b13d-403">If `null` is passed for the exception, then the HttpError constuctor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## <a name="run-the-add-in"></a><span data-ttu-id="5b13d-404">运行加载项</span><span class="sxs-lookup"><span data-stu-id="5b13d-404">Run the add-in</span></span>

1. <span data-ttu-id="5b13d-405">请确保 OneDrive 中有一些文件，以便可以验证结果。</span><span class="sxs-lookup"><span data-stu-id="5b13d-405">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="5b13d-p183">在 Visual Studio 中，按 F5。PowerPoint 将打开，“主页”\*\*\*\* 功能区上会有一个“SSO ASP.NET”\*\*\*\* 组。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p183">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="5b13d-408">按此组中的“显示加载项”\*\*\*\* 按钮，在任务窗格中查看此加载项的 UI。</span><span class="sxs-lookup"><span data-stu-id="5b13d-408">Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.</span></span>

1. <span data-ttu-id="5b13d-p184">按“从 OneDrive 获取我的文件”\*\*\*\* 按钮。如果尚未登录 Office，便会看到登录提示。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p184">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="5b13d-411">如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已在 PowerPoint 中更改过，也不例外。</span><span class="sxs-lookup"><span data-stu-id="5b13d-411">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="5b13d-412">在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。</span><span class="sxs-lookup"><span data-stu-id="5b13d-412">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="5b13d-413">为了防止发生这种情况，请务必先*关闭其他所有 Office 应用*，再按“从 OneDrive 获取我的文件”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5b13d-413">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="5b13d-p186">登录后，便会在按钮下方看到 OneDrive 文件和文件夹列表。此过程可能需要超过 15 秒才能完成，特别是首次使用时。</span><span class="sxs-lookup"><span data-stu-id="5b13d-p186">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
