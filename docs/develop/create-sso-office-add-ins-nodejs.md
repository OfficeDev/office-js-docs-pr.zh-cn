---
title: 创建使用单一登录的 Node.js Office 加载项
description: ''
ms.date: 12/7/2018
ms.openlocfilehash: 5a3a4d398842119dc8c0d935f83a233313bb35c4
ms.sourcegitcommit: f130dfa423bc536804fa4a90e1183d85f1bef730
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/12/2018
ms.locfileid: "27243483"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="7c845-102">创建使用单一登录的 Node.js Office 加载项（预览）</span><span class="sxs-lookup"><span data-stu-id="7c845-102">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="7c845-p101">用户可以登录 Office，Office Web 加载项能够利用此登录进程，授权用户访问加载项和 Microsoft Graph，而无需要求用户再登录一次。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="7c845-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="7c845-105">本文将逐步介绍如何在使用 Node.js 和 Express 生成的加载项中启用单一登录 (SSO) 。</span><span class="sxs-lookup"><span data-stu-id="7c845-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> 

> [!NOTE]
> <span data-ttu-id="7c845-106">有关与此类似的 ASP.NET 加载项文章，请参阅[创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)。</span><span class="sxs-lookup"><span data-stu-id="7c845-106">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="7c845-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="7c845-107">Prerequisites</span></span>

* <span data-ttu-id="7c845-108">[节点和 npm](https://nodejs.org/en/) 版本 6.9.4 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7c845-108">[Node and npm](https://nodejs.org/en/), version 6.9.4 or later</span></span>

* <span data-ttu-id="7c845-109">[Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）</span><span class="sxs-lookup"><span data-stu-id="7c845-109">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="7c845-110">TypeScript 版本 2.2.2 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7c845-110">TypeScript version 2.2.2 or later</span></span>

* <span data-ttu-id="7c845-111">Office 2016 版本 1708（生成号 8424.nnnn）或更高版本（Office 365 订阅版本，有时亦称为“即点即用”）</span><span class="sxs-lookup"><span data-stu-id="7c845-111">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”)</span></span>

  <span data-ttu-id="7c845-p102">可能必须成为 Office 预览体验成员，才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。</span><span class="sxs-lookup"><span data-stu-id="7c845-p102">You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="7c845-114">创建起始项目</span><span class="sxs-lookup"><span data-stu-id="7c845-114">Set up the starter project</span></span>

1. <span data-ttu-id="7c845-115">克隆或下载 [Office 外接程序 NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso) 中的存储库。</span><span class="sxs-lookup"><span data-stu-id="7c845-115">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="7c845-116">示例有三个版本：</span><span class="sxs-lookup"><span data-stu-id="7c845-116">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="7c845-p103">**Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。</span><span class="sxs-lookup"><span data-stu-id="7c845-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span> 
    > * <span data-ttu-id="7c845-p104">如果完成了本文中的过程，该示例的**已完成**版本会与所生成的外接程序类似，只不过完成的项目具有对本文文本冗余的代码注释。若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Completed”，并跳过**编写客户端代码**和**编写服务器端代码**部分。</span><span class="sxs-lookup"><span data-stu-id="7c845-p104">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="7c845-122">“已完成的多租户”\*\*\*\* 版本是支持多租户的已完成示例。</span><span class="sxs-lookup"><span data-stu-id="7c845-122">The **Completed Multitenant** version is a completed sample that supports multitenancy.</span></span> <span data-ttu-id="7c845-123">如果要使用 SSO 从不同域支持 Microsoft 帐户，则浏览此示例。</span><span class="sxs-lookup"><span data-stu-id="7c845-123">Explore this sample if you intend to support Microsoft accounts from different domains with SSO.</span></span>
    >
    > <span data-ttu-id="7c845-124">_不论使用何种版本，都需要信任本地主机的证书。请参阅存储库自述文件中的“重要”说明。_</span><span class="sxs-lookup"><span data-stu-id="7c845-124">_Regardless of which version you use, you will need to trust a certificate for the localhost. See the "IMPORTANT" note in the Readme of the repo._</span></span>

2. <span data-ttu-id="7c845-125">在“Before”\*\*\*\* 文件夹中打开 Git bash 控制台。</span><span class="sxs-lookup"><span data-stu-id="7c845-125">Open a Git bash console in the **Before** folder.</span></span>

3. <span data-ttu-id="7c845-126">在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。</span><span class="sxs-lookup"><span data-stu-id="7c845-126">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

4. <span data-ttu-id="7c845-127">在控制台中输入 `npm run build `，以生成项目。</span><span class="sxs-lookup"><span data-stu-id="7c845-127">Enter `npm run build ` in the console to build the project.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="7c845-p106">可能会看到一些生成错误，提示某些变量已声明但未使用。请忽略这些错误。之所以会看到这些错误是因为，示例项目的“之前”版本缺少某代码，将在后续步骤中添加。</span><span class="sxs-lookup"><span data-stu-id="7c845-p106">You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="7c845-131">向 Azure AD v2.0 终结点注册外接程序</span><span class="sxs-lookup"><span data-stu-id="7c845-131">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="7c845-132">通常编写以下指令，以便可以在多个位置使用它们。</span><span class="sxs-lookup"><span data-stu-id="7c845-132">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="7c845-133">对于此文章，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="7c845-133">For this article do the following:</span></span>
- <span data-ttu-id="7c845-134">将占位符“$ADD-IN-NAME$”\*\*\*\* 替换为 `“Office-Add-in-NodeJS-SSO`。</span><span class="sxs-lookup"><span data-stu-id="7c845-134">Replace the placeholder **$ADD-IN-NAME$** with `“Office-Add-in-NodeJS-SSO`.</span></span>
- <span data-ttu-id="7c845-135">将占位符“$FQDN-WITHOUT-PROTOCOL$”\*\*\*\* 替换为 `localhost:3000`。</span><span class="sxs-lookup"><span data-stu-id="7c845-135">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.</span></span>
- <span data-ttu-id="7c845-136">在“选择权限”\*\*\*\* 对话框中指定权限时，请选中以下权限对应的框。</span><span class="sxs-lookup"><span data-stu-id="7c845-136">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="7c845-137">外接程序本身真正需要的只是第一项权限，但 Office 主机必须有 `profile` 权限，才能获取访问外接程序 Web 应用程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-137">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="7c845-138">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="7c845-138">Files.Read.All</span></span>
    * <span data-ttu-id="7c845-139">配置文件</span><span class="sxs-lookup"><span data-stu-id="7c845-139">profile</span></span>

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="7c845-140">同意管理员访问外接程序</span><span class="sxs-lookup"><span data-stu-id="7c845-140">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="7c845-141">配置外接程序</span><span class="sxs-lookup"><span data-stu-id="7c845-141">Configure the add-in</span></span>

1. <span data-ttu-id="7c845-p109">在代码编辑器中打开 src\server.ts 文件。顶部附近存在对 `AuthModule` 类的构造函数的调用。该构造函数中存在一些需要为其分配值的字符串参数。</span><span class="sxs-lookup"><span data-stu-id="7c845-p109">In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.</span></span>

2. <span data-ttu-id="7c845-145">对于 `client_id` 属性，将占位符 `{client GUID}` 替换为注册外接程序时保存的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="7c845-145">For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in.</span></span> <span data-ttu-id="7c845-146">完成后，应该有一个括在单引号中的 GUID。</span><span class="sxs-lookup"><span data-stu-id="7c845-146">When you are done, there should just be a GUID in single quotation marks.</span></span> <span data-ttu-id="7c845-147">而不应存在任何“{}”字符。</span><span class="sxs-lookup"><span data-stu-id="7c845-147">There should not be any "{}" characters.</span></span>

3. <span data-ttu-id="7c845-148">对于 `client_secret` 属性，将占位符 `{client secret}` 替换为注册外接程序时保存的应用程序机密。</span><span class="sxs-lookup"><span data-stu-id="7c845-148">For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.</span></span>

4. <span data-ttu-id="7c845-p111">对于 `audience` 属性，将占位符 `{audience GUID}` 替换为注册外接程序时保存的应用程序 ID。（即分配给 `client_id` 属性的同一值）。</span><span class="sxs-lookup"><span data-stu-id="7c845-p111">For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)</span></span>
  
3. <span data-ttu-id="7c845-151">在分配给 `issuer` 属性的字符串中，将看到占位符 *{O365 tenant GUID}*。</span><span class="sxs-lookup"><span data-stu-id="7c845-151">In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*.</span></span> <span data-ttu-id="7c845-152">将此占位符替换为 Office 365 租户 ID。</span><span class="sxs-lookup"><span data-stu-id="7c845-152">Replace this with the Office 365 tenancy ID.</span></span> <span data-ttu-id="7c845-153">使用[查找 Office 365 租户 ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) 中的一种方法来获取 ID。</span><span class="sxs-lookup"><span data-stu-id="7c845-153">Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span> <span data-ttu-id="7c845-154">完成后，`issuer` 属性值应如下所示：</span><span class="sxs-lookup"><span data-stu-id="7c845-154">When you are done, the `issuer` property value should look something like this:</span></span>

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. <span data-ttu-id="7c845-155">保持 `AuthModule` 构造函数中的其他参数不变。</span><span class="sxs-lookup"><span data-stu-id="7c845-155">Leave the other parameters in the `AuthModule` constructor unchanged.</span></span> <span data-ttu-id="7c845-156">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-156">Save and close the file.</span></span>

1. <span data-ttu-id="7c845-157">在项目的根目录中，打开外接程序清单文件“Office-Add-in-NodeJS-SSO.xml”。</span><span class="sxs-lookup"><span data-stu-id="7c845-157">In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.</span></span>

1. <span data-ttu-id="7c845-158">滚动到文件底部。</span><span class="sxs-lookup"><span data-stu-id="7c845-158">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="7c845-159">在结束 `</VersionOverrides>` 标记的正上方，你会发现以下标记：</span><span class="sxs-lookup"><span data-stu-id="7c845-159">Just above the end `</VersionOverrides>` tag, you will find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="7c845-160">将标记中的*两处*占位符“{application_GUID here}”均替换为在注册外接程序时复制的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="7c845-160">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="7c845-161">（由于 ID 并不包含“{}”，因此请勿添加它们。）这与在 web.config 中对 ClientID 和 Audience 所使用的 ID 相同。</span><span class="sxs-lookup"><span data-stu-id="7c845-161">(The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="7c845-162">“Resource”\*\*\*\* 值是向注册的外接程序添加 Web API 平台时设置的“应用程序 ID URI”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-162">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="7c845-163">仅在通过 AppSource 销售加载项时，才使用 **Scopes** 部分生成许可对话框。</span><span class="sxs-lookup"><span data-stu-id="7c845-163">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="7c845-164">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-164">Save and close the file.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="7c845-165">编写客户端代码</span><span class="sxs-lookup"><span data-stu-id="7c845-165">Code the client side</span></span>

1. <span data-ttu-id="7c845-p115">打开 **public** 文件夹中的 program.js 文件。其中已存在一些代码：</span><span class="sxs-lookup"><span data-stu-id="7c845-p115">Open the program.js file in the **public** folder. It already has some code in it:</span></span>

    * <span data-ttu-id="7c845-168">针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。</span><span class="sxs-lookup"><span data-stu-id="7c845-168">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="7c845-169">`showResult` 方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。</span><span class="sxs-lookup"><span data-stu-id="7c845-169">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="7c845-170">`logErrors` 方法，用于记录最终用户不应看到的控制台错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-170">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

11. <span data-ttu-id="7c845-p116">在向 `Office.initialize` 分配函数下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p116">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="7c845-173">加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-173">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="7c845-174">计数器变量 `timesGetOneDriveFilesHasRun` 以及标志变量 `triedWithoutForceConsent` 和 `timesMSGraphErrorReceived` 用于确保用户不会重复循环失败的尝试来获取令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-174">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="7c845-p118">虽然 `getDataWithToken` 方法是在下一步中创建，但请注意，它会将 `forceConsent` 选项设置为 `false`。有关详细信息，请参阅下一步。</span><span class="sxs-lookup"><span data-stu-id="7c845-p118">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="7c845-p119">在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p119">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="7c845-179">[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 是 Office.js 中新增的 API，支持外接程序向 Office 主机应用程序（Excel、PowerPoint、Word 等）请求获取对外接程序的访问令牌（对于已登录 Office 的用户）。</span><span class="sxs-lookup"><span data-stu-id="7c845-179">The [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="7c845-180">反过来，Office 主机应用程序会向 Azure AD 2.0 终结点请求获取令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-180">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="7c845-181">由于已在注册加载项时将 Office 主机预授权给加载项，因此 Azure AD 将会发送令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-181">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="7c845-182">如果用户未登录 Office，Office 主机会提示用户登录。</span><span class="sxs-lookup"><span data-stu-id="7c845-182">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="7c845-183">options 参数将 `forceConsent` 设置为 `false`，因此用户不会在每次使用加载项时都看到提示，要求其许可向 Office 主机授予对加载项的访问权限。</span><span class="sxs-lookup"><span data-stu-id="7c845-183">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="7c845-184">用户首次运行加载项时，`getAccessTokenAsync` 调用会失败，但在后续步骤中添加的错误处理逻辑会自动重新调用（`forceConsent` 选项设置为 `true`），并提示用户许可，但仅限首次运行。</span><span class="sxs-lookup"><span data-stu-id="7c845-184">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="7c845-185">`handleClientSideErrors` 方法将在后续步骤中创建。</span><span class="sxs-lookup"><span data-stu-id="7c845-185">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="7c845-p122">用以下行替换 TODO1。可以在后续步骤中创建 `getData` 方法和服务器端“/api/values”路由。相对 URL 用于终结点，因为它必须与外接程序托管在相同的域中。</span><span class="sxs-lookup"><span data-stu-id="7c845-p122">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="7c845-p123">在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p123">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="7c845-p124">此方法调用指定 Web API 终结点，并向它传递访问令牌，这也是 Office 主机应用用于获取对加载项的访问权限的令牌。在服务器端，此访问令牌将用于“代表”流，以获取对 Microsoft Graph 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-p124">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="7c845-193">`handleServerSideErrors` 方法将在后续步骤中创建。</span><span class="sxs-lookup"><span data-stu-id="7c845-193">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="7c845-194">创建错误处理方法</span><span class="sxs-lookup"><span data-stu-id="7c845-194">Create the error-handling methods</span></span>

1. <span data-ttu-id="7c845-195">在 `getData` 方法下方，添加下列方法。</span><span class="sxs-lookup"><span data-stu-id="7c845-195">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="7c845-196">当 Office 主机无法获取对加载项 Web 服务的访问令牌时，此方法便会处理加载项客户端中的错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-196">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="7c845-197">这些错误通过错误代码进行报告，因此下面的方法使用 `switch` 语句区分它们。</span><span class="sxs-lookup"><span data-stu-id="7c845-197">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Microsoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user triggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="7c845-198">将 `TODO2` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-198">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="7c845-199">如果用户未登录或用户取消（未响应）提供辅助身份验证因素的提示，错误 13001 发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-199">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="7c845-200">无论属于上述哪种情况，代码都会重新运行 `getDataWithToken` 方法，并设置强制登录提示选项。</span><span class="sxs-lookup"><span data-stu-id="7c845-200">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="7c845-201">将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-201">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="7c845-202">如果用户登录或许可被中止，错误 13002 发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-202">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="7c845-203">建议用户重试一次，但只能重试一次。</span><span class="sxs-lookup"><span data-stu-id="7c845-203">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="7c845-204">将 `TODO4` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-204">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="7c845-205">如果用户用于登录的帐户既不是工作帐户或学校帐户，也不是 Microsoft 帐户，错误 13003 发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-205">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft account.</span></span> <span data-ttu-id="7c845-206">建议用户注销，再使用受支持的帐户类型重新登录。</span><span class="sxs-lookup"><span data-stu-id="7c845-206">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="7c845-207">此方法不处理错误 13004 和 13005，因为它们只在开发期间出现。</span><span class="sxs-lookup"><span data-stu-id="7c845-207">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="7c845-208">无法通过运行时代码进行修复，并且向最终用户报告这两个错误也没有意义。</span><span class="sxs-lookup"><span data-stu-id="7c845-208">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="7c845-p130">将 `TODO5` 替换为下列代码。如果 Office 主机中出现可能表明主机处于不稳定状态的未指定错误，就会发生错误 13006。建议用户重启 Office。</span><span class="sxs-lookup"><span data-stu-id="7c845-p130">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="7c845-212">将 `TODO6` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-212">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="7c845-213">如果 Office 主机与 AAD 之间的交互出现问题，导致主机无法获得对加载项 Web 服务/应用的访问令牌，错误 13007 发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-213">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="7c845-214">这可能由于暂时网络问题所致。</span><span class="sxs-lookup"><span data-stu-id="7c845-214">This may be a temporary network issue.</span></span> <span data-ttu-id="7c845-215">建议用户稍后重试。</span><span class="sxs-lookup"><span data-stu-id="7c845-215">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="7c845-216">将 `TODO7` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-216">Replace `TODO7` with the following code.</span></span> <span data-ttu-id="7c845-217">如果用户触发的操作未等到上一次调用完成就调用了 `getAccessTokenAsync`，错误 13008 发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-217">Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="7c845-218">将 `TODO8` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-218">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="7c845-219">如果加载项不支持强制许可，但调用 `getAccessTokenAsync` 时 `forceConsent` 选项设置为 `true`，错误 13009 发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-219">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="7c845-220">通常情况下，如果发生这种情况，代码应自动重新运行 `getAccessTokenAsync`，同时将许可选项设置为 `false`。</span><span class="sxs-lookup"><span data-stu-id="7c845-220">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="7c845-221">不过，在某些情况下，调用将 `forceConsent` 设置为 `true` 的方法本身就是在自动响应调用将选项设置为 `false` 的方法时出现的错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-221">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="7c845-222">此时，不得重试代码，而是应建议用户注销并重新登录。</span><span class="sxs-lookup"><span data-stu-id="7c845-222">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="7c845-223">将 `TODO9` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-223">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. <span data-ttu-id="7c845-p134">在 `handleClientSideErrors` 方法下方，添加下列方法。此方法可处理加载项 Web 服务中发生的以下错误：无法执行代表流，或无法从 Microsoft Graph 获取数据。</span><span class="sxs-lookup"><span data-stu-id="7c845-p134">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="7c845-p135">将 `TODO10` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p135">Replace `TODO10` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="7c845-p136">一些 Azure Active Directory 配置要求用户，必须提供其他一个或多个身份验证因素，才能访问一些 Microsoft Graph 目标（例如 OneDrive），即使用户仅使用密码就能登录 Office，也不例外。在这种情况下，AAD 将发送包含错误 50076 的响应（具有 `Claims` 属性）。</span><span class="sxs-lookup"><span data-stu-id="7c845-p136">There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property.</span></span> 
    * <span data-ttu-id="7c845-230">Office 主机应获取新令牌（使用 **Claims** 值作为 `authChallenge` 选项）。</span><span class="sxs-lookup"><span data-stu-id="7c845-230">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="7c845-231">这就指示 AAD 提示用户进行所有必需形式的身份验证。</span><span class="sxs-lookup"><span data-stu-id="7c845-231">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. <span data-ttu-id="7c845-p138">*在上一步添加的代码的最后一个右大括号正下方*，将 `TODO11` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p138">Replace `TODO11` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="7c845-234">错误 65001 表示未许可授予（或已撤消）一个或多个对 Microsoft Graph 的访问权限。</span><span class="sxs-lookup"><span data-stu-id="7c845-234">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="7c845-235">加载项应获取新令牌（`forceConsent` 选项设置为 `true`）。</span><span class="sxs-lookup"><span data-stu-id="7c845-235">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="7c845-p139">*在上一步添加的代码的最后一个右大括号正下方*，将 `TODO12` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p139">Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="7c845-238">错误 70011 表示已请求获取的范围（权限）无效。</span><span class="sxs-lookup"><span data-stu-id="7c845-238">Error 70011 means that an invalid scope (permission) has been requested.</span></span> <span data-ttu-id="7c845-239">加载项应报告此错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-239">The add-in should report the error.</span></span>
    * <span data-ttu-id="7c845-240">代码使用 AAD 错误号记录其他任何错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-240">The code logs any other error with an AAD error number.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="7c845-p141">*在上一步添加的代码的最后一个右大括号正下方*，将 `TODO13` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p141">Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="7c845-243">如果 `access_as_user` 范围（权限）不在访问令牌中，此令牌由加载项客户端发送到 AAD 以便在代表流中使用，那么在后续步骤中创建的服务器端代码将发送以 `... expected access_as_user` 结尾的消息。</span><span class="sxs-lookup"><span data-stu-id="7c845-243">Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="7c845-244">加载项应报告此错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-244">The add-in should report the error.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="7c845-p142">*在上一步添加的代码的最后一个右大括号正下方*，将 `TODO14` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p142">Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="7c845-247">不太可能将到期或无效令牌发送到 Microsoft Graph，但如果这种情况确实发生，在后续步骤中创建的服务器端代码将以字符串 `Microsoft Graph error` 结尾。</span><span class="sxs-lookup"><span data-stu-id="7c845-247">It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.</span></span>
    * <span data-ttu-id="7c845-248">在这种情况下，加载项应重置 `timesGetOneDriveFilesHasRun` 计数器和 `timesGetOneDriveFilesHasRun` 标志变量，再重新调用按钮处理程序方法，以从头开始执行整个身份验证流程。</span><span class="sxs-lookup"><span data-stu-id="7c845-248">In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method.</span></span> <span data-ttu-id="7c845-249">但它只能执行此操作一次。</span><span class="sxs-lookup"><span data-stu-id="7c845-249">But it should do this only once.</span></span> <span data-ttu-id="7c845-250">如果再次发生，它应只记录此错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-250">If it happens again, it should just log the error.</span></span>
    * <span data-ttu-id="7c845-251">如果连续两次出现这种情况，代码会记录此错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-251">The code logs the error if it happens twice in succession.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }        
    }
    ```

1. <span data-ttu-id="7c845-252">*在上一步添加的代码的最后一个右大括号正下方*，将 `TODO15` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-252">Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*.</span></span>

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a><span data-ttu-id="7c845-253">编写服务器端代码</span><span class="sxs-lookup"><span data-stu-id="7c845-253">Code the server side</span></span>

<span data-ttu-id="7c845-254">有两个需要修改的服务器端文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-254">There are two server-side files that need to be modified.</span></span> 
- <span data-ttu-id="7c845-p144">src\auth.js 提供授权 helper 函数。它已具有在各种授权流中使用的泛型成员。我们需要为其添加可实现“代表”流的函数。</span><span class="sxs-lookup"><span data-stu-id="7c845-p144">The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.</span></span>
- <span data-ttu-id="7c845-p145">src\server.js文件具有运行服务器和 Express 中间件所需的基本成员。我们需要为其添加服务于主页和 Web API 的函数，以获取 Microsoft Graph 数据。</span><span class="sxs-lookup"><span data-stu-id="7c845-p145">The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.</span></span>

### <a name="create-a-method-to-exchange-tokens"></a><span data-ttu-id="7c845-260">创建交换令牌的方法</span><span class="sxs-lookup"><span data-stu-id="7c845-260">Create a method to exchange tokens</span></span>

1. <span data-ttu-id="7c845-p146">打开 \src\auth.ts 文件。将下面的方法添加到 `AuthModule` 类。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p146">Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="7c845-p147">`jwt` 参数是对应用的访问令牌。在“代表”流中，它与 AAD 进行交换，以获取对资源的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-p147">The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.</span></span>
    * <span data-ttu-id="7c845-266">虽然 scopes 参数具有默认值，但在此示例中，它将被调用代码覆盖。</span><span class="sxs-lookup"><span data-stu-id="7c845-266">The scopes parameter has a default value, but in this sample it will be overridden by the calling code.</span></span>
    * <span data-ttu-id="7c845-267">resource 参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="7c845-267">The resource parameter is optional.</span></span> <span data-ttu-id="7c845-268">它不应在[安全令牌服务 (STS)](https://docs.microsoft.com/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) 是 AAD V 2.0 终结点时使用。</span><span class="sxs-lookup"><span data-stu-id="7c845-268">It should not be used when the [Secure Token Service (STS)](https://docs.microsoft.com/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) is the AAD V 2.0 endpoint.</span></span> <span data-ttu-id="7c845-269">V 2.0 终结点从作用域推断资源，如果在 HTTP 请求中发送资源，则它将返回错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-269">The resource parameter is optional. It should not be used when the STS is the AAD V 2.0 endpoint. The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request.</span></span> 
    * <span data-ttu-id="7c845-270">`catch` 信息块中抛出异常*不会*导致立即向客户端发送“500 内部服务器错误”。</span><span class="sxs-lookup"><span data-stu-id="7c845-270">Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="7c845-271">server.js 文件中的调用代码会捕获此异常，并将它变成发送到客户端的错误消息。</span><span class="sxs-lookup"><span data-stu-id="7c845-271">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. <span data-ttu-id="7c845-p150">将 `TODO3` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p150">Replace `TODO3` with the following code. About this code, note:</span></span>
    * <span data-ttu-id="7c845-p151">支持“代表”流的 STS 需要 HTTP 请求正文中的某些属性/值对。此代码构造一个可成为请求正文的对象。</span><span class="sxs-lookup"><span data-stu-id="7c845-p151">An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request.</span></span> 
    * <span data-ttu-id="7c845-276">仅当资源传递到方法时，才将 resource 属性添加到正文。</span><span class="sxs-lookup"><span data-stu-id="7c845-276">A resource property is added to the body if, and only if, a resource was passed to the method.</span></span>

        ```typescript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            } 
        ```

3. <span data-ttu-id="7c845-277">将 `TODO4` 替换为以下代码，用于将 HTTP 请求发送到 STS 的令牌终结点。</span><span class="sxs-lookup"><span data-stu-id="7c845-277">Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.</span></span>

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. <span data-ttu-id="7c845-278">将 `TODO5` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-278">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="7c845-279">请注意，抛出异常*不会*导致立即向客户端发送“500 内部服务器错误”。</span><span class="sxs-lookup"><span data-stu-id="7c845-279">Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="7c845-280">server.js 文件中的调用代码会捕获此异常，并将它变成发送到客户端的错误消息。</span><span class="sxs-lookup"><span data-stu-id="7c845-280">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. <span data-ttu-id="7c845-p153">将 `TODO6` 替换为以下代码。请注意，代码会返回并保留对资源的访问令牌及其到期时间。调用代码可以重用对资源的未到期访问令牌，避免了对 STS 执行不必要的调用。下一部分将介绍如何执行此操作。</span><span class="sxs-lookup"><span data-stu-id="7c845-p153">Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.</span></span>

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. <span data-ttu-id="7c845-285">保存但不关闭文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-285">Save the file, but don't close it.</span></span>

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a><span data-ttu-id="7c845-286">使用“代表”流创建一个获取资源访问权限的方法</span><span class="sxs-lookup"><span data-stu-id="7c845-286">Create a method to get access to the resource using the "on behalf of" flow</span></span>

1. <span data-ttu-id="7c845-p154">还是在 src/auth.ts 中，将下面的方法添加到 `AuthModule` 类。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p154">Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="7c845-289">上面关于 `exchangeForToken` 方法参数的注释也适用于此方法的参数。</span><span class="sxs-lookup"><span data-stu-id="7c845-289">The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.</span></span>
    * <span data-ttu-id="7c845-p155">方法先检查对资源（尚未到期且不会在下一分钟到期）的访问令牌是否有永久性存储。仅在需要的情况下，它才会调用在上一部分中创建的 `exchangeForToken` 方法。</span><span class="sxs-lookup"><span data-stu-id="7c845-p155">The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.</span></span>

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. <span data-ttu-id="7c845-292">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-292">Save and close the file.</span></span>

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a><span data-ttu-id="7c845-293">创建服务于外接程序主页和数据的终结点</span><span class="sxs-lookup"><span data-stu-id="7c845-293">Create the endpoints that will serve the add-in's home page and data</span></span>

1. <span data-ttu-id="7c845-294">打开 src\server.ts 文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-294">Open the src\server.ts file.</span></span> 

2. <span data-ttu-id="7c845-p156">将以下方法添加到文件底部。此方法将为外接程序的主页提供服务。外接程序清单指定主页 URL。</span><span class="sxs-lookup"><span data-stu-id="7c845-p156">Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.</span></span>

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. <span data-ttu-id="7c845-p157">将以下方法添加到文件底部。此方法将处理对 `values` API 的任何请求。</span><span class="sxs-lookup"><span data-stu-id="7c845-p157">Add the following method to bottom of the file. This method will handle any requests for the `values` API.</span></span>
    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. <span data-ttu-id="7c845-300">将 `TODO7` 替换为以下代码行，可验证从 Office 主机应用程序收到的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7c845-300">Replace `TODO7` with the following code which validates the access token received from the Office host application.</span></span> <span data-ttu-id="7c845-301">`verifyJWT` 方法在 src\auth.ts 文件中进行定义。</span><span class="sxs-lookup"><span data-stu-id="7c845-301">The `verifyJWT` method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="7c845-302">它始终验证受众和颁发者。</span><span class="sxs-lookup"><span data-stu-id="7c845-302">It always validates the audience and the issuer.</span></span> <span data-ttu-id="7c845-303">此可选参数可用于指定是否还要它验证访问令牌中的作用域是否为 `access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="7c845-303">We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`.</span></span> <span data-ttu-id="7c845-304">这是用户和 Office 主机通过“代表”流获取对 Microsoft Graph 的访问令牌时，唯一需要拥有的对外接程序的权限。</span><span class="sxs-lookup"><span data-stu-id="7c845-304">This is the only permission to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf" flow.</span></span> 

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > <span data-ttu-id="7c845-305">只可使用 `access_as_user` 作用域授权 API 为 Office 外接程序处理代表流。服务中的其他 API 应有自己的作用域要求。</span><span class="sxs-lookup"><span data-stu-id="7c845-305">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="7c845-306">这就限制了使用 Office 获得的令牌可以访问的内容。</span><span class="sxs-lookup"><span data-stu-id="7c845-306">This limits what can be accessed with the tokens that Office acquires.</span></span>

5. <span data-ttu-id="7c845-p160">将 `TODO8` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p160">Replace `TODO8` with the following code. Note the following about this code:</span></span>

    * <span data-ttu-id="7c845-309">`acquireTokenOnBehalfOf` 调用中不包括 resource 参数，因为 `AuthModule` 对象 (`auth`) 是使用不支持 resource 属性的 AAD V2.0 终结点进行构造。</span><span class="sxs-lookup"><span data-stu-id="7c845-309">The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.</span></span>
    * <span data-ttu-id="7c845-310">调用的第二个参数指定了加载项获取 OneDrive 上用户文件和文件夹列表时所需的权限。</span><span class="sxs-lookup"><span data-stu-id="7c845-310">The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive.</span></span> <span data-ttu-id="7c845-311">（之所以不需要 `profile` 权限是因为，只有当 Office 主机获取对加载项的访问令牌时，才需要此权限，用此令牌交换对 Microsoft Graph 的访问令牌时并不需要。）</span><span class="sxs-lookup"><span data-stu-id="7c845-311">(The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)</span></span>

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. <span data-ttu-id="7c845-p162">将 `TODO9` 替换为以下代码行。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p162">Replace `TODO9` with the following line. Note the following about this code:</span></span>

    * <span data-ttu-id="7c845-314">MSGraphHelper 类是在 Src\msgraph helper.ts 中定义。</span><span class="sxs-lookup"><span data-stu-id="7c845-314">The MSGraphHelper class is defined in src\msgraph-helper.ts.</span></span> 
    * <span data-ttu-id="7c845-315">通过指定只需要 name 属性和前 3 项，可以最大限度地减少必须返回的数据。</span><span class="sxs-lookup"><span data-stu-id="7c845-315">We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.</span></span>

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

7. <span data-ttu-id="7c845-316">将 `TODO10` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-316">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="7c845-317">请注意，此代码处理 Microsoft Graph 返回的“401 未授权”错误，此错误表示令牌到期或无效。</span><span class="sxs-lookup"><span data-stu-id="7c845-317">Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token.</span></span> <span data-ttu-id="7c845-318">由于令牌暂留逻辑应该会阻止，因此这种情况不太可能会发生。</span><span class="sxs-lookup"><span data-stu-id="7c845-318">It is very unlikely that this would ever happen since the token persisting logic should prevent it.</span></span> <span data-ttu-id="7c845-319">（请参阅上面的**使用“代表”流创建方法以获取对资源的访问权限**部分。）如果这种情况确实发生，此代码会将错误中继到客户端，并在错误名称中显示“Microsoft Graph 错误”。</span><span class="sxs-lookup"><span data-stu-id="7c845-319">(See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name.</span></span> <span data-ttu-id="7c845-320">（请参阅在之前步骤中在 program.js 文件内创建的 `handleClientSideErrors` 方法。）在后续步骤中添加到 ODataHelper.js 文件的代码有助于处理 Microsoft Graph 返回的错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-320">(See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.</span></span>

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. <span data-ttu-id="7c845-p164">将 `TODO11` 替换为以下代码。请注意，Microsoft Graph 对每项返回某 OData 元数据和 **eTag** 属性，即使 `name` 是所请求的唯一属性，也不例外。代码仅向客户端发送项名称。</span><span class="sxs-lookup"><span data-stu-id="7c845-p164">Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.</span></span>

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. <span data-ttu-id="7c845-324">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-324">Save and close the file.</span></span>

### <a name="add-response-handling-to-the-odatahelper"></a><span data-ttu-id="7c845-325">向 ODataHelper 添加响应处理</span><span class="sxs-lookup"><span data-stu-id="7c845-325">Add response handling to the ODataHelper</span></span>

1. <span data-ttu-id="7c845-326">打开文件 src\odata-helper.ts。</span><span class="sxs-lookup"><span data-stu-id="7c845-326">Open the file src\odata-helper.ts.</span></span> <span data-ttu-id="7c845-327">文件几乎已完成。</span><span class="sxs-lookup"><span data-stu-id="7c845-327">The file is almost complete.</span></span> <span data-ttu-id="7c845-328">缺少的是，请求“结束”事件处理程序的回调主体。</span><span class="sxs-lookup"><span data-stu-id="7c845-328">What's missing is the body of the callback to the handler for the request "end" event.</span></span> <span data-ttu-id="7c845-329">将 `TODO` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="7c845-329">Replace the `TODO` with the following code.</span></span> <span data-ttu-id="7c845-330">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-330">About this code note:</span></span>

    * <span data-ttu-id="7c845-331">OData 终结点返回的响应可能是错误（如 401）。如果终结点需要访问令牌，但令牌无效或到期，就会生成 401 错误。</span><span class="sxs-lookup"><span data-stu-id="7c845-331">The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired.</span></span> <span data-ttu-id="7c845-332">不过，错误消息仍是*消息*，而不是 `https.get` 调用中的错误，因此不会触发 `https.get` 末尾的 `on('error', reject)` 代码行。</span><span class="sxs-lookup"><span data-stu-id="7c845-332">But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered.</span></span> <span data-ttu-id="7c845-333">所以，代码区分成功 (200) 消息和错误消息，并向调用方发送 JSON 对象，其中包含请求获取的 OData 或错误消息。</span><span class="sxs-lookup"><span data-stu-id="7c845-333">So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.</span></span>

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  <span data-ttu-id="7c845-p167">将 `TODO1` 替换为下列代码。请注意，此代码假设数据是以 JSON 形式返回。</span><span class="sxs-lookup"><span data-stu-id="7c845-p167">Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.</span></span>

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  <span data-ttu-id="7c845-p168">将 `TODO2` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7c845-p168">Replace `TODO2` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="7c845-338">OData 源返回的错误响应将始终包含 statusCode，通常是 statusMessage。</span><span class="sxs-lookup"><span data-stu-id="7c845-338">An error response from an OData source will always have a statusCode and usually a statusMessage.</span></span> <span data-ttu-id="7c845-339">一些 OData 源还向主体添加错误属性，以提供更多信息，如内部或更具体的代码和消息。</span><span class="sxs-lookup"><span data-stu-id="7c845-339">Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.</span></span>
    * <span data-ttu-id="7c845-340">Promise 对象已解析，未被拒绝。</span><span class="sxs-lookup"><span data-stu-id="7c845-340">The Promise object is resolved, not rejected.</span></span> <span data-ttu-id="7c845-341">Web 服务在服务器间调用 OData 终结点时，`https.get` 运行。</span><span class="sxs-lookup"><span data-stu-id="7c845-341">The `https.get` runs when a web service calls an OData endpoint server-to-server.</span></span> <span data-ttu-id="7c845-342">但这种调用出现的上下文是，客户端在 Web 服务中调用 Web API。</span><span class="sxs-lookup"><span data-stu-id="7c845-342">But that call comes in the context of a call from a client to a web API in the web service.</span></span> <span data-ttu-id="7c845-343">如果此“内部”请求被拒绝，客户端向 Web 服务发送的“外部”请求永不会完成。</span><span class="sxs-lookup"><span data-stu-id="7c845-343">The "outer" request from the client to the web service never completes if this "inner" request is rejected.</span></span> <span data-ttu-id="7c845-344">此外，如果 `http.get` 的调用方需要将 OData 终结点返回的错误中继到客户端，必须解析具有自定义 `Error` 对象的请求。</span><span class="sxs-lookup"><span data-stu-id="7c845-344">Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.</span></span>

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;
    
    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. <span data-ttu-id="7c845-345">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="7c845-345">Save and close the file.</span></span>

## <a name="deploy-the-add-in"></a><span data-ttu-id="7c845-346">部署外接程序</span><span class="sxs-lookup"><span data-stu-id="7c845-346">Deploy the add-in</span></span>

<span data-ttu-id="7c845-347">现在，你需要让 Office 知道在哪里可以找到该外接程序。</span><span class="sxs-lookup"><span data-stu-id="7c845-347">Now you need to let Office know where to find the add-in.</span></span>

1. <span data-ttu-id="7c845-348">创建网络共享，或[将文件夹共享到网络](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11))。</span><span class="sxs-lookup"><span data-stu-id="7c845-348">Create a network share, or [share a folder to the network](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span></span>

2. <span data-ttu-id="7c845-349">将 Office-Add-in-NodeJS-SSO.xml 清单文件从项目根目录复制到共享文件夹。</span><span class="sxs-lookup"><span data-stu-id="7c845-349">Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.</span></span>

3. <span data-ttu-id="7c845-350">启动 PowerPoint 并打开文档。</span><span class="sxs-lookup"><span data-stu-id="7c845-350">Launch PowerPoint and open a document.</span></span>

4. <span data-ttu-id="7c845-351">选择“文件”\*\*\*\* 选项卡，然后选择“选项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-351">Choose the **File** tab, and then choose **Options**.</span></span>

5. <span data-ttu-id="7c845-352">选择**信任中心**，然后选择**信任中心设置**按钮。</span><span class="sxs-lookup"><span data-stu-id="7c845-352">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

6. <span data-ttu-id="7c845-353">选择“受信任的外接程序目录”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-353">Choose **Trusted Add-ins Catalogs**.</span></span>

7. <span data-ttu-id="7c845-354">在“目录 URL”\*\*\*\* 字段中，输入包含 Office-Add-in-NodeJS-SSO.xml 的文件夹共享的网络路径，然后选择“添加目录”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-354">In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.</span></span>

8. <span data-ttu-id="7c845-355">选中“显示在菜单中”\*\*\*\* 复选框，然后选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-355">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

9. <span data-ttu-id="7c845-p171">随后会出现一条消息，告知你下次启动 Microsoft Office 时将应用你的设置。关闭 PowerPoint。</span><span class="sxs-lookup"><span data-stu-id="7c845-p171">A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.</span></span>

## <a name="build-and-run-the-project"></a><span data-ttu-id="7c845-358">生成和运行项目</span><span class="sxs-lookup"><span data-stu-id="7c845-358">Build and run the project</span></span>

<span data-ttu-id="7c845-p172">根据是否使用 Visual Studio Code，有两种生成和运行项目的方法。对于这两种方法，当更改代码时，该项目将生成和自动生成并重新运行。</span><span class="sxs-lookup"><span data-stu-id="7c845-p172">There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.</span></span>

1. <span data-ttu-id="7c845-361">如果使用的不是 Visual Studio Code：</span><span class="sxs-lookup"><span data-stu-id="7c845-361">If you are not using Visual Studio Code:</span></span> 
 1. <span data-ttu-id="7c845-362">打开节点终端，然后导航到该项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="7c845-362">Open a node terminal and navigate to the root folder of the project.</span></span>
 2. <span data-ttu-id="7c845-363">在终端中，输入 **npm run build**。</span><span class="sxs-lookup"><span data-stu-id="7c845-363">In the terminal, enter **npm run build**.</span></span> 
 3. <span data-ttu-id="7c845-364">打开第二个节点终端，然后导航到该项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="7c845-364">Open a second node terminal and navigate to the root folder of the project.</span></span>
 4. <span data-ttu-id="7c845-365">在终端中，输入 **npm run start**。</span><span class="sxs-lookup"><span data-stu-id="7c845-365">In the terminal, enter **npm run start**.</span></span>

2. <span data-ttu-id="7c845-366">如果使用的是 VS Code：</span><span class="sxs-lookup"><span data-stu-id="7c845-366">If you are using VS Code:</span></span>
 1. <span data-ttu-id="7c845-367">通过 VS Code 打开项目。</span><span class="sxs-lookup"><span data-stu-id="7c845-367">Open the project in VS Code.</span></span>
 2. <span data-ttu-id="7c845-368">按 CTRL-SHIFT-B 生成项目。</span><span class="sxs-lookup"><span data-stu-id="7c845-368">Press CTRL-SHIFT-B to build the project.</span></span>
 3. <span data-ttu-id="7c845-369">按 F5 键在调试会话中运行该项目。</span><span class="sxs-lookup"><span data-stu-id="7c845-369">Press F5 to run the project in a debugging session.</span></span>


## <a name="add-the-add-in-to-an-office-document"></a><span data-ttu-id="7c845-370">将外接程序添加到 Office 文档</span><span class="sxs-lookup"><span data-stu-id="7c845-370">Add the add-in to an Office document</span></span>

1. <span data-ttu-id="7c845-371">重启 PowerPoint，然后打开或创建演示文稿。</span><span class="sxs-lookup"><span data-stu-id="7c845-371">Restart PowerPoint and open or create a presentation.</span></span>

1. <span data-ttu-id="7c845-372">如果功能区上未显示“开发工具”\*\*\*\* 选项卡，请按照以下步骤操作来启用它：</span><span class="sxs-lookup"><span data-stu-id="7c845-372">If the **Developer** tab is not visible on the ribbon, enable it with the following steps:</span></span>
 1. <span data-ttu-id="7c845-373">依次导航到“文件”\*\*\*\* | “选项”\*\*\*\* | “自定义功能区”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-373">Navigate to **File** | **Options** | **Customize Ribbon**.</span></span>
 2. <span data-ttu-id="7c845-374">在“自定义功能区”\*\*\*\* 页面右侧的控件名称树形结构中，点击相应复选框以启用“开发工具”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-374">Click the check box to enable **Developer** in the tree of control names on the right of the **Customize Ribbon** page.</span></span>
 3. <span data-ttu-id="7c845-375">按“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-375">Press **OK**.</span></span>

2. <span data-ttu-id="7c845-376">在 PowerPoint 中的“开发工具”\*\*\*\* 选项卡上，选择“我的外接程序”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-376">On the **Developer** tab in PowerPoint, choose **My Add-ins**.</span></span>

3. <span data-ttu-id="7c845-377">选择“共享文件夹”\*\*\*\* 选项卡。</span><span class="sxs-lookup"><span data-stu-id="7c845-377">Select the **SHARED FOLDER** tab.</span></span>

4. <span data-ttu-id="7c845-378">选择“SSO NodeJS 示例”\*\*\*\*，然后选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-378">Choose **SSO NodeJS Sample**, and then select **OK**.</span></span>

5. <span data-ttu-id="7c845-379">“主页”\*\*\*\* 功能区上有一个名为“**SSO NodeJS**”的新组，包含标记为“显示外接程序”\*\*\*\* 的按钮和一个图标。</span><span class="sxs-lookup"><span data-stu-id="7c845-379">On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="7c845-380">测试加载项</span><span class="sxs-lookup"><span data-stu-id="7c845-380">Test the add-in</span></span>

1. <span data-ttu-id="7c845-381">请确保 OneDrive 中有一些文件，以便可以验证结果。</span><span class="sxs-lookup"><span data-stu-id="7c845-381">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

2. <span data-ttu-id="7c845-382">单击“显示加载项”\*\*\*\* 按钮，打开此加载项。</span><span class="sxs-lookup"><span data-stu-id="7c845-382">Click **Show Add-in** button to open the add-in.</span></span>

2. <span data-ttu-id="7c845-p173">此时，加载项打开并显示欢迎页。单击“从 OneDrive 获取我的文件”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="7c845-p173">The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.</span></span>

2. <span data-ttu-id="7c845-p174">如果你已登录 Office，则 OneDrive 上的文件和文件夹列表将显示在该按钮的下方。首次操作需要的时间可能会超过 15 秒。</span><span class="sxs-lookup"><span data-stu-id="7c845-p174">If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.</span></span>

3. <span data-ttu-id="7c845-387">如果没有登录 Office，弹出窗口将打开并提示进行登录。</span><span class="sxs-lookup"><span data-stu-id="7c845-387">If you are not signed into Office, a popup will open and prompt you to sign in.</span></span> <span data-ttu-id="7c845-388">完成登录后，文件和文件夹列表将在几秒钟后显示。</span><span class="sxs-lookup"><span data-stu-id="7c845-388">After you have completed the sign-in, the list of your files and folders will appear after a few seconds.</span></span> <span data-ttu-id="7c845-389">*请勿再次按下此按钮。*</span><span class="sxs-lookup"><span data-stu-id="7c845-389">*You should not press the button a second time.*</span></span>

> [!NOTE]
> <span data-ttu-id="7c845-390">如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已在 PowerPoint 中更改过，也不例外。</span><span class="sxs-lookup"><span data-stu-id="7c845-390">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="7c845-391">在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。</span><span class="sxs-lookup"><span data-stu-id="7c845-391">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="7c845-392">为了防止发生这种情况，请务必先*关闭其他所有 Office 应用程序*，然后再按“从 OneDrive 获取我的文件”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="7c845-392">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
