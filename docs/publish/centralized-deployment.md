---
title: 通过 Office 365 管理中心进行集中部署来发布 Office 加载项
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 422fae01328a76c0d815fcf007b9970c3eceba36
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871849"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="bb5b4-102">通过 Office 365 管理中心进行集中部署来发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="bb5b4-102">Publish Office Add-ins using Centralized Deployment via the Office 365 admin center</span></span>

<span data-ttu-id="bb5b4-p101">通过 Office 365 管理中心，管理员可以轻松地为组织内的用户和组部署 Office 加载项。通过管理中心部署加载项后，用户可立即在其 Office 应用程序中使用此加载项，而无需进行客户端配置。可以通过集中部署来部署内部加载项以及 ISV 提供的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p101">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="bb5b4-106">Office 365 管理中心当前支持以下方案：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-106">The Office 365 admin center currently supports the following scenarios:</span></span>

- <span data-ttu-id="bb5b4-107">为个人、组或组织集中部署新的和更新的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-107">Centralized Deployment of new and updated add-ins to individuals, groups, or an organization.</span></span>
- <span data-ttu-id="bb5b4-108">可以部署到多个平台，其中包括 Windows 和 Office Online，即将推出对 Mac 的支持。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-108">Deployment to multiple platforms, including Windows and Office Online, with Mac coming soon.</span></span>
- <span data-ttu-id="bb5b4-109">到英语语言租户和全球范围租户的部署。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-109">Deployment to English language and worldwide tenants.</span></span>
- <span data-ttu-id="bb5b4-110">部署云托管的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-110">Deployment of cloud-hosted add-ins.</span></span>
- <span data-ttu-id="bb5b4-111">部署托管在防火墙内的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-111">Deployment of add-ins that are hosted within a firewall.</span></span>
- <span data-ttu-id="bb5b4-112">部署 AppSource 加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-112">Deployment of AppSource add-ins.</span></span>
- <span data-ttu-id="bb5b4-113">当用户启动 Office 应用时自动为用户安装加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-113">Automatic installation of an add-in for users when they launch the Office application.</span></span>
- <span data-ttu-id="bb5b4-114">当管理员禁用或删除加载项，或者将用户从 Azure Active Directory 或从已部署加载项的组中删除时，则自动为用户删除该加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-114">Automatic removal of an add-in for users if the admin turns off or deletes the add-in, or if users are removed from Azure Active Directory or from a group to which the add-in has been deployed.</span></span>

<span data-ttu-id="bb5b4-115">如果组织满足使用集中部署的所有要求，则建议 Office 365 管理员通过集中部署在组织内部署 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-115">Centralized Deployment is the recommended way for an Office 365 admin to deploy Office Add-ins within an organization, provided that the organization meets all requirements for using Centralized Deployment.</span></span> <span data-ttu-id="bb5b4-116">有关如何确定组织是否可以使用集中部署的信息，请参阅[确定加载项集中部署是否适用于你的 Office 365 组织](https://support.office.com/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-b4527d49-4073-4b43-8274-31b7a3166f92)。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-116">For information about how to determine if your organization can use Centralized Deployment, see [Determine if Centralized Deployment of add-ins works for your Office 365 organization](https://support.office.com/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-b4527d49-4073-4b43-8274-31b7a3166f92).</span></span>

> [!NOTE]
> <span data-ttu-id="bb5b4-p103">在没有连接到 Office 365 的本地环境中，或若要部署 SharePoint 加载项或定目标到 Office 2013 的 Office 加载项，请使用 [SharePoint 加载项目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。 若要部署 COM/VSTO 加载项，请使用 ClickOnce 或 Windows Installer，如[部署 Office 解决方案](/visualstudio/vsto/deploying-an-office-solution)中所述。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p103">In an on-premises environment with no connection to Office 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint add-in catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).</span></span>

## <a name="recommended-approach-for-deploying-office-add-ins"></a><span data-ttu-id="bb5b4-119">部署 Office 加载项的推荐方法</span><span class="sxs-lookup"><span data-stu-id="bb5b4-119">Recommended approach for deploying Office Add-ins</span></span>

<span data-ttu-id="bb5b4-p104">请考虑分阶段部署 Office 加载项，以确保部署顺利进行。建议使用以下计划：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p104">Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:</span></span>

1. <span data-ttu-id="bb5b4-p105">为一小部分的企业利益干系人和 IT 部门成员部署加载项。 如果部署成功，则转到步骤 2。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p105">Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.</span></span>

2. <span data-ttu-id="bb5b4-p106">为企业内更多的将使用加载项的个人部署加载项。 如果部署成功，则转到步骤 3。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p106">Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.</span></span>

3. <span data-ttu-id="bb5b4-126">为所有将使用加载项的个人部署加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-126">Deploy the add-in to the full set of individuals who will be using the add-in.</span></span>

<span data-ttu-id="bb5b4-127">根据目标受众的规模，可能需要在此过程中添加步骤或删除步骤。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-127">Depending on the size of the target audience, you may want to add steps to or remove steps from this procedure.</span></span>

## <a name="publish-an-office-add-in-via-centralized-deployment"></a><span data-ttu-id="bb5b4-128">通过集中部署发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="bb5b4-128">Publish an Office Add-in via Centralized Deployment</span></span>

<span data-ttu-id="bb5b4-129">在开始之前，请按照[确定加载项集中部署是否适用于你的 Office 365 组织](https://support.office.com/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-b4527d49-4073-4b43-8274-31b7a3166f92)中所述确认组织是否满足使用集中部署的所有要求。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-129">Before you begin, confirm that your organization meets all requirements for using Centralized Deployment, as described in [Determine if Centralized Deployment of add-ins works for your Office 365 organization](https://support.office.com/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-b4527d49-4073-4b43-8274-31b7a3166f92).</span></span>

<span data-ttu-id="bb5b4-130">如果组织满足所有要求，请完成以下步骤以通过集中部署发布 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-130">If your organization meets all requirements, complete the following steps to publish an Office Add-in via Centralized Deployment:</span></span>

1. <span data-ttu-id="bb5b4-131">使用工作或学校帐户登录 Office 365。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-131">Sign in to Office 365 with your work or school account.</span></span>
2. <span data-ttu-id="bb5b4-132">选择左上角的应用启动器图标，然后选择“**管理员**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-132">Select the app launcher icon in the upper-left and choose **Admin**.</span></span>
3. <span data-ttu-id="bb5b4-133">在导航菜单中，按“**显示更多内容**”，然后选择“**设置**” > “**服务和加载项**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-133">In the navigation menu, press **Show more**, then choose **Settings** > **Services & add-ins**.</span></span>
4. <span data-ttu-id="bb5b4-134">如果在页面顶部看到公布新的 Office 365 管理中心的消息，请选择该消息以转至“管理中心预览版”（请参阅[关于 Office 365 管理中心](https://support.office.com/en-ie/article/About-the-Office-365-admin-center-758befc4-0888-4009-9f14-0d147402fd23)）。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-134">If you see a message on the top of the page announcing the new Office 365 admin center, choose the message to go to the Admin Center Preview (see [About the Office 365 admin center](https://support.office.com/en-ie/article/About-the-Office-365-admin-center-758befc4-0888-4009-9f14-0d147402fd23)).</span></span>
5. <span data-ttu-id="bb5b4-135">在页面顶部选择“**部署加载项**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-135">Choose **Deploy Add-In** at the top of the page.</span></span>
6. <span data-ttu-id="bb5b4-136">查看要求后，请选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-136">Choose **Next** after reviewing the requirements.</span></span>
7. <span data-ttu-id="bb5b4-137">在“**集中部署**”页面上，选择以下选项之一：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-137">Choose one of the following options on the **Centralized Deployment** page:</span></span>

    - <span data-ttu-id="bb5b4-138">**我想从 Office 应用商店添加加载项。**</span><span class="sxs-lookup"><span data-stu-id="bb5b4-138">**I want to add an Add-In from the Office Store.**</span></span>
    - <span data-ttu-id="bb5b4-139">**我在此设备上具有清单文件 (.xml)。**</span><span class="sxs-lookup"><span data-stu-id="bb5b4-139">**I have the manifest file (.xml) on this device.**</span></span> <span data-ttu-id="bb5b4-140">对于此选项，请选择“浏览”\*\*\*\* 以找到想要使用的清单文件 (.xml)。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-140">For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.</span></span>
    - <span data-ttu-id="bb5b4-p108">**我具有清单文件的 URL。** 对于此选项，请在提供的字段中键入清单的 URL。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p108">**I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.</span></span>

    ![Office 365 管理中心中的新加载项对话框](../images/new-add-in.png)

8. <span data-ttu-id="bb5b4-144">如果选择了此选项以从 Office 应用商店添加某个加载项，请选择该加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-144">If you selected the option to add an add-in from the Office Store, select the add-in.</span></span> <span data-ttu-id="bb5b4-145">可以通过“**为你推荐**”、“**评级**”或“**名称**”类别，查看可用的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-145">You can view available add-ins via categories of **Suggested for you**, **Rating**, or **Name**.</span></span> <span data-ttu-id="bb5b4-146">仅能从 Office 应用商店添加免费加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-146">You may only add free add-ins from Office Store.</span></span> <span data-ttu-id="bb5b4-147">目前不支持添加付费加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-147">Adding paid add-ins isn't currently supported.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bb5b4-148">使用 Office 应用商店选项，无需干预，用户即可自动获得加载项的更新和增强功能。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-148">With the Office Store option, updates and enhancements to the add-in are automatically available to users without your intervention.</span></span>

    ![在 Office 365 管理中心中选择“加载项”对话框](../images/select-an-add-in.png)

9. <span data-ttu-id="bb5b4-150">查看加载项的详细信息后，请选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-150">Choose **Next** after reviewing the add-in details.</span></span>

    ![Office 365 管理中心中的 Power BI 磁贴加载项页面](../images/power-bi-tiles.png)

10. <span data-ttu-id="bb5b4-152">在“**编辑有权访问的人员**”页面上，选择“**任何人**”或“**特定用户/组**”或“**仅自己**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-152">On the **Edit who has access** page, choose **Everyone**, **Specific Users/Groups**, or **Only me**.</span></span> <span data-ttu-id="bb5b4-153">使用“搜索”框查找要向其部署加载项的用户和组。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-153">Use the search box to find the users and groups to whom you want to deploy the add-in.</span></span>

    ![编辑 Office 365 管理中心内的“谁有权访问”页面](../images/power-bi-tiles-edit.png)

    > [!NOTE]
    > <span data-ttu-id="bb5b4-155">用于加载项的[单一登录 (SSO)](/office/dev/add-ins/develop/sso-in-office-add-ins) 系统目前处于预览状态，不应用于生产加载项。部署使用 SSO 的加载项时，分配的用户和组也将与共享相同 Azure App ID 的加载项共享。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-155">A [single sign-on (SSO)](/office/dev/add-ins/develop/sso-in-office-add-ins) system for add-ins is currently in preview and should not be used for production add-ins. When an add-in using SSO is deployed, the users and groups assigned are also shared with add-ins that share the same Azure App ID.</span></span> <span data-ttu-id="bb5b4-156">对用户分配进行的任何更改也会应用于这些加载项。相关加载项显示在此页面上。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-156">Any changes to user assignments are also applied to those add-ins. The related add-ins are shown on this page.</span></span> <span data-ttu-id="bb5b4-157">仅对于 SSO 加载项，此页面将显示加载项所需的 Microsoft Graph 权限的列表。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-157">For SSO add-ins only, this page displays the list of Microsoft Graph permissions that the add-in requires.</span></span>

11. <span data-ttu-id="bb5b4-158">完成后，选择“**保存**”以保存清单。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-158">When finished, choose **Save** to save the manifest.</span></span> <span data-ttu-id="bb5b4-159">此过程可能最多用时 3 分钟。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-159">This process may take up to three minutes.</span></span> <span data-ttu-id="bb5b4-160">然后，按“**下一步**”完成演练。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-160">Then, finish the walkthrough by pressing **Next**.</span></span> <span data-ttu-id="bb5b4-161">现在，可以看到此加载项与其他应用一起显示在 Office 365 中。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-161">You now see your add-in along with other apps in Office 365.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bb5b4-162">管理员选择“保存”\*\*\*\* 后，即表示向所有用户授予许可。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-162">When an administrator chooses **Save**, consent is given for all users.</span></span>

    ![Office 365 管理中心内的应用列表](../images/citations.png)

> [!TIP]
> <span data-ttu-id="bb5b4-164">为组织中的用户和/或组部署新加载项时，请考虑向他们发送一封电子邮件，说明加载项的应用场景和使用方式，并添加相关帮助内容、FAQ 或其他支持资源的链接。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-164">When you deploy a new add-in to users and/or groups in your organization, consider sending them an email that describes when and how to use the add-in, and includes links to relevant Help content, FAQs, or other support resources.</span></span>

## <a name="considerations-when-granting-access-to-an-add-in"></a><span data-ttu-id="bb5b4-165">授予加载项的访问权限时的注意事项</span><span class="sxs-lookup"><span data-stu-id="bb5b4-165">Considerations when granting access to an add-in</span></span>

<span data-ttu-id="bb5b4-p113">管理员可以将加载项分配给组织中的每个人或组织内的特定用户和/或组。 以下列表描述了每个选项的含义：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p113">Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:</span></span>

- <span data-ttu-id="bb5b4-p114">**每个人**：顾名思义，此选项为租户中的每位用户分配加载项。请谨慎使用此选项，且仅应用于真正在组织中通用的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p114">**Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.</span></span>

- <span data-ttu-id="bb5b4-p115">**用户**：如果将加载项分配给单个用户，则每次要将其分配给其他用户时，都需要更新此加载项的集中部署设置。 同样，每次要删除某个用户对该加载项的访问权限时，都需要更新该加载项的集中部署设置。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p115">**Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.</span></span>

- <span data-ttu-id="bb5b4-p116">**组**：如果将加载项分配给组，则会自动为被添加到此组的用户分配此加载项。 同样，当将某个用户从组中删除时，此用户将自动失去对此加载项的访问权限。 在上述任一情况下，均无需从 Office 365 管理处执行任何额外操作。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p116">**Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in. Likewise, when a user is removed from a group, the user automatically loses access to the add-in. In either case, no additional action is required from the Office 365 admin.</span></span>

<span data-ttu-id="bb5b4-p117">一般情况下，为了便于维护，我们建议尽可能使用组来分配加载项。 但是，在想要将加载项的访问权限限制在极少数用户的情况下，将加载项分配给特定用户的做法可能更为实用。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p117">In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.</span></span> 

## <a name="add-in-states"></a><span data-ttu-id="bb5b4-177">加载项状态</span><span class="sxs-lookup"><span data-stu-id="bb5b4-177">Add-in states</span></span>

<span data-ttu-id="bb5b4-178">下表介绍了加载项的不同状态。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-178">The following table describes the different states of an add-in.</span></span>

|<span data-ttu-id="bb5b4-179">状态</span><span class="sxs-lookup"><span data-stu-id="bb5b4-179">State</span></span>|<span data-ttu-id="bb5b4-180">此状态如何出现</span><span class="sxs-lookup"><span data-stu-id="bb5b4-180">How the state occurs</span></span>|<span data-ttu-id="bb5b4-181">影响</span><span class="sxs-lookup"><span data-stu-id="bb5b4-181">Impact</span></span>|
|-----|--------------------|------|
|<span data-ttu-id="bb5b4-182">**活动**</span><span class="sxs-lookup"><span data-stu-id="bb5b4-182">**Active**</span></span>|<span data-ttu-id="bb5b4-183">管理员已上传加载项并已将其分配给用户和/或组。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-183">Admin uploaded the add-in and assigned it to users and/or groups.</span></span>|<span data-ttu-id="bb5b4-184">已为其分配加载项的用户和/或组，可在相关的 Office 客户端中找到它。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-184">Users and/or groups assigned to the add-in see it in the relevant Office clients.</span></span>|
|<span data-ttu-id="bb5b4-185">**已禁用**</span><span class="sxs-lookup"><span data-stu-id="bb5b4-185">**Turned off**</span></span>|<span data-ttu-id="bb5b4-186">管理员已禁用加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-186">Admin turned off the add-in.</span></span>|<span data-ttu-id="bb5b4-p118">已为其分配加载项的用户和/或组不再能够访问它。 如果加载项状态从“已禁用”\*\*\*\* 更改为“活动”\*\*\*\*，则用户和组将重新获得对它的访问权限。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p118">Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.</span></span>|
|<span data-ttu-id="bb5b4-189">**已删除**</span><span class="sxs-lookup"><span data-stu-id="bb5b4-189">**Deleted**</span></span>|<span data-ttu-id="bb5b4-190">管理员已删除加载项。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-190">Admin deleted the add-in.</span></span>|<span data-ttu-id="bb5b4-191">已为其分配加载项的用户和/或组不再能够访问它。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-191">Users and/or groups assigned the add-in no longer have access to it.</span></span>|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a><span data-ttu-id="bb5b4-192">更新通过集中部署发布的 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="bb5b4-192">Updating Office Add-ins that are published via Centralized Deployment</span></span>

<span data-ttu-id="bb5b4-p119">如果通过集中部署发布 Office 加载项，则在该加载项的 Web 应用程序中实现对该 Web 应用程序所做的更改后，将自动向所有用户提供相应的更改。 对加载项的 [XML 清单文件](../develop/add-in-manifests.md)所做的更改（例如，更新加载项的图标、文本或加载项命令）以以下方式实现：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p119">After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users as soon as those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md), for example, to update the add-in's icon, text, or add-in commands, happen as follows:</span></span>

- <span data-ttu-id="bb5b4-p120">**业务线加载项**：如果管理员在通过 Office 365 管理中心实施集中部署时显式上传了清单文件，则管理员必须上传包含所需更改的新清单文件。 上传更新后的清单文件后，加载项就会在下次相关 Office 应用启动时更新。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p120">**Line-of-business add-in**: If an admin explicitly uploaded a manifest file when implementing Centralized Deployment via the Office 365 admin center, the admin must upload a new manifest file that contains the desired changes. After the updated manifest file has been uploaded, the next time the relevant Office applications start, the add-in will update.</span></span>

- <span data-ttu-id="bb5b4-197">**Office 应用商店加载项**：如果管理员在通过 Office 365 管理中心实施集中部署时从 Office 应用商店选择了加载项，并且 Office 应用商店更新了此加载项，则此加载项稍后将通过集中部署更新。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-197">**Office Store add-in**: If an admin selected an add-in from the Office Store when implementing Centralized Deployment via the Office 365 admin center, and the add-in updates in the Office Store, the add-in will update later via Centralized Deployment.</span></span> <span data-ttu-id="bb5b4-198">加载项会在下次相关 Office 应用启动时更新。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-198">The next time the relevant Office applications start, the add-in will update.</span></span>

## <a name="end-user-experience-with-add-ins"></a><span data-ttu-id="bb5b4-199">加载项最终用户体验</span><span class="sxs-lookup"><span data-stu-id="bb5b4-199">End user experience with add-ins</span></span>

<span data-ttu-id="bb5b4-200">通过集中部署发布加载项后，最终用户可以在加载项支持的任何平台上开始使用它。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-200">After an add-in has been published via Centralized Deployment, end users may start using it on any platform that the add-in supports.</span></span> 

<span data-ttu-id="bb5b4-p122">如果外接程序支持外接程序命令，则这些命令将出现在为其部署外接程序的所有用户的 Office 应用程序功能区上。 在以下的示例中，**搜索引文**命令将显示在**引文**加载项的功能区上。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-p122">If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.</span></span> 

![屏幕截图显示了 Office 功能区的一部分，其中突出显示了引文加载项中的“搜索引文”命令](../images/search-citation.png)

<span data-ttu-id="bb5b4-204">如果加载项不支持加载项命令，用户可以通过执行以下操作将其添加到 Office 应用程序中：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-204">If the add-in does not support add-in commands, users can add it to their Office application by doing the following:</span></span>

1. <span data-ttu-id="bb5b4-205">在 Word 2016 或更高版本、Excel 2016 或更高版本，或 PowerPoint 2016 或更高版本，选择“**插入**” > “**我的加载项**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-205">In Word 2016 or later, Excel 2016 or later, or PowerPoint 2016 or later, choose **Insert** > **My Add-ins**.</span></span>
2. <span data-ttu-id="bb5b4-206">在加载项窗口中选择“**管理托管**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-206">Choose the **Admin Managed** tab in the add-in window.</span></span>
3. <span data-ttu-id="bb5b4-207">选择加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-207">Choose the add-in, and then choose **Add**.</span></span>

    ![屏幕截图显示 Office 应用程序的“Office 加载项”页的“管理托管”选项卡。 引文加载项显示在此选项卡上。](../images/office-add-ins-admin-managed.png)

<span data-ttu-id="bb5b4-210">但是，对于 Outlook 2016 或更高版本，用户可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="bb5b4-210">However, for Outlook 2016 or later, users can do the following:</span></span>

1. <span data-ttu-id="bb5b4-211">在 Outlook 中，选择“**开始**” > “**应用商店**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-211">In Outlook, choose **Home** > **Store**.</span></span>
2. <span data-ttu-id="bb5b4-212">选择“加载项”选项卡下的“**管理员管理**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-212">Choose the **Admin-managed** item under the add-in tab.</span></span>
3. <span data-ttu-id="bb5b4-213">选择加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="bb5b4-213">Choose the add-in, and then choose **Add**.</span></span>

    ![屏幕截图显示了 Outlook 应用程序的“应用商店”页面的管理员管理区域。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a><span data-ttu-id="bb5b4-215">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bb5b4-215">See also</span></span>

- [<span data-ttu-id="bb5b4-216">确定加载项的集中式部署是否适用于你的 Office 365 组织</span><span class="sxs-lookup"><span data-stu-id="bb5b4-216">Determine if Centralized Deployment of add-ins works for your Office 365 organization</span></span>](https://support.office.com/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-b4527d49-4073-4b43-8274-31b7a3166f92)
