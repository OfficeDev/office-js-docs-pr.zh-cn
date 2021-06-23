---
title: Office集中部署发布加载项Microsoft 365 管理中心
description: 了解如何使用集中部署来部署内部外接程序以及 ISV 提供的外接程序。
ms.date: 03/22/2021
localization_priority: Normal
ms.openlocfilehash: 3107fc58601683f5356594f2f79ffc5293ea266f
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076649"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a><span data-ttu-id="71eaa-103">Office集中部署发布加载项Microsoft 365 管理中心</span><span class="sxs-lookup"><span data-stu-id="71eaa-103">Publish Office Add-ins using Centralized Deployment via the Microsoft 365 admin center</span></span>

<span data-ttu-id="71eaa-104">通过Microsoft 365 管理中心，管理员能够轻松地将Office加载项部署到其组织内的用户和组。</span><span class="sxs-lookup"><span data-stu-id="71eaa-104">The Microsoft 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization.</span></span> <span data-ttu-id="71eaa-105">通过管理中心部署加载项后，用户可立即在其 Office 应用程序中使用此加载项，而无需进行客户端配置。</span><span class="sxs-lookup"><span data-stu-id="71eaa-105">Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required.</span></span> <span data-ttu-id="71eaa-106">可以通过集中部署来部署内部加载项以及 ISV 提供的加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-106">You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="71eaa-107">当前Microsoft 365 管理中心支持以下方案。</span><span class="sxs-lookup"><span data-stu-id="71eaa-107">The Microsoft 365 admin center currently supports the following scenarios.</span></span>

- <span data-ttu-id="71eaa-108">为个人、组或组织集中部署新的和更新的加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-108">Centralized Deployment of new and updated add-ins to individuals, groups, or an organization.</span></span>
- <span data-ttu-id="71eaa-109">部署到多个客户端平台，Windows、Mac 和 Web。</span><span class="sxs-lookup"><span data-stu-id="71eaa-109">Deployment to multiple client platforms, including Windows, Mac, and the web.</span></span> <span data-ttu-id="71eaa-110">对于 Outlook，也支持部署到 iOS 和 Android。</span><span class="sxs-lookup"><span data-stu-id="71eaa-110">For Outlook, deployment to iOS and Android is also supported.</span></span> <span data-ttu-id="71eaa-111"> (但是，尽管支持在 iPad 上安装 Excel、Outlook、Word 和 PowerPoint 加载项，但不支持集中部署到 iPad。) </span><span class="sxs-lookup"><span data-stu-id="71eaa-111">(However, while user installation of Excel, Outlook, Word, and PowerPoint add-ins on iPad is supported, Centralized Deployment to iPad is **not** supported.)</span></span>
- <span data-ttu-id="71eaa-112">到英语语言租户和全球范围租户的部署。</span><span class="sxs-lookup"><span data-stu-id="71eaa-112">Deployment to English language and worldwide tenants.</span></span>
- <span data-ttu-id="71eaa-113">部署云托管的加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-113">Deployment of cloud-hosted add-ins.</span></span>
- <span data-ttu-id="71eaa-114">部署托管在防火墙内的加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-114">Deployment of add-ins that are hosted within a firewall.</span></span>
- <span data-ttu-id="71eaa-115">部署 AppSource 加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-115">Deployment of AppSource add-ins.</span></span>
- <span data-ttu-id="71eaa-116">当用户启动 Office 应用时自动为用户安装加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-116">Automatic installation of an add-in for users when they launch the Office application.</span></span>
- <span data-ttu-id="71eaa-117">当管理员禁用或删除加载项，或者将用户从 Azure Active Directory 或从已部署加载项的组中删除时，则自动为用户删除该加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-117">Automatic removal of an add-in for users if the admin turns off or deletes the add-in, or if users are removed from Azure Active Directory or from a group to which the add-in has been deployed.</span></span>

<span data-ttu-id="71eaa-118">建议管理员在Microsoft 365部署 Office 外接程序，只要组织满足使用集中部署的所有要求。</span><span class="sxs-lookup"><span data-stu-id="71eaa-118">Centralized Deployment is the recommended way for a Microsoft 365 admin to deploy Office Add-ins within an organization, provided that the organization meets all requirements for using Centralized Deployment.</span></span> <span data-ttu-id="71eaa-119">若要了解如何确定组织能否使用集中部署，请参阅确定加载项集中部署是否适用于Microsoft 365[部署](/office365/admin/manage/centralized-deployment-of-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="71eaa-119">For information about how to determine if your organization can use Centralized Deployment, see [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/office365/admin/manage/centralized-deployment-of-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="71eaa-120">在未连接到 Microsoft 365 或部署面向 Office 2013 的 SharePoint 加载项或 Office 加载项的本地环境中，请使用[SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="71eaa-120">In an on-premises environment with no connection to Microsoft 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint app catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span> <span data-ttu-id="71eaa-121">若要部署 COM/VSTO 加载项，请使用 ClickOnce 或 Windows Installer，如[部署 Office 解决方案](/visualstudio/vsto/deploying-an-office-solution)中所述。</span><span class="sxs-lookup"><span data-stu-id="71eaa-121">To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).</span></span>

## <a name="recommended-approach-for-deploying-office-add-ins"></a><span data-ttu-id="71eaa-122">部署 Office 加载项的推荐方法</span><span class="sxs-lookup"><span data-stu-id="71eaa-122">Recommended approach for deploying Office Add-ins</span></span>

<span data-ttu-id="71eaa-p105">请考虑分阶段部署 Office 加载项，以确保部署顺利进行。建议使用以下计划：</span><span class="sxs-lookup"><span data-stu-id="71eaa-p105">Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:</span></span>

1. <span data-ttu-id="71eaa-p106">为一小部分的企业利益干系人和 IT 部门成员部署加载项。 如果部署成功，则转到步骤 2。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p106">Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.</span></span>

2. <span data-ttu-id="71eaa-p107">为企业内更多的将使用加载项的个人部署加载项。 如果部署成功，则转到步骤 3。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p107">Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.</span></span>

3. <span data-ttu-id="71eaa-129">为所有将使用加载项的个人部署加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-129">Deploy the add-in to the full set of individuals who will be using the add-in.</span></span>

<span data-ttu-id="71eaa-130">根据目标受众的规模，可能需要在此过程中添加步骤或删除步骤。</span><span class="sxs-lookup"><span data-stu-id="71eaa-130">Depending on the size of the target audience, you may want to add steps to or remove steps from this procedure.</span></span>

## <a name="publish-an-office-add-in-via-centralized-deployment"></a><span data-ttu-id="71eaa-131">通过集中部署发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="71eaa-131">Publish an Office Add-in via Centralized Deployment</span></span>

<span data-ttu-id="71eaa-132">在开始之前，请确认您的组织满足使用集中部署的所有要求，如确定外接程序的集中部署是否适用于 Microsoft 365[组织中所述](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="71eaa-132">Before you begin, confirm that your organization meets all requirements for using Centralized Deployment, as described in [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).</span></span>

<span data-ttu-id="71eaa-133">如果组织满足所有要求，请完成以下步骤以通过集中部署发布 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="71eaa-133">If your organization meets all requirements, complete the following steps to publish an Office Add-in via Centralized Deployment:</span></span>

1. <span data-ttu-id="71eaa-134">Sign in to Microsoft 365 with your work or education account.</span><span class="sxs-lookup"><span data-stu-id="71eaa-134">Sign in to Microsoft 365 with your work or education account.</span></span>
1. <span data-ttu-id="71eaa-135">选择左上角的应用启动器图标，然后选择“**管理员**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-135">Select the app launcher icon in the upper-left and choose **Admin**.</span></span>
1. <span data-ttu-id="71eaa-136">在导航菜单中，选择"**显示更多"，** 然后选择 **"设置**  >  **集成应用"。**</span><span class="sxs-lookup"><span data-stu-id="71eaa-136">In the navigation menu, select **Show more**, then choose **Settings** > **Integrated apps**.</span></span>
1. <span data-ttu-id="71eaa-137">在页面顶部，选择"**外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="71eaa-137">At the top of the page, choose **Add-ins**.</span></span>
1. <span data-ttu-id="71eaa-138">如果在页面顶部看到一条宣布推出新 Microsoft 365 管理中心 的消息，请选择该消息以转到管理中心预览版 (请参阅关于 Microsoft 365 管理中心[) 。](/microsoft-365/admin/admin-overview/about-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="71eaa-138">If you see a message on the top of the page announcing the new Microsoft 365 admin center, choose the message to go to the Admin Center Preview (see [About the Microsoft 365 admin center](/microsoft-365/admin/admin-overview/about-the-admin-center)).</span></span>
1. <span data-ttu-id="71eaa-139">在页面顶部选择“**部署加载项**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-139">Choose **Deploy Add-In** at the top of the page.</span></span>
1. <span data-ttu-id="71eaa-140">查看要求后，请选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-140">Choose **Next** after reviewing the requirements.</span></span>
1. <span data-ttu-id="71eaa-141">在“**集中部署**”页面上，选择以下选项之一：</span><span class="sxs-lookup"><span data-stu-id="71eaa-141">Choose one of the following options on the **Centralized Deployment** page:</span></span>

    - <span data-ttu-id="71eaa-142">**我想从 Office 应用商店添加加载项。**</span><span class="sxs-lookup"><span data-stu-id="71eaa-142">**I want to add an Add-In from the Office Store.**</span></span>
    - <span data-ttu-id="71eaa-143">**我在此设备上具有清单文件 (.xml)。**</span><span class="sxs-lookup"><span data-stu-id="71eaa-143">**I have the manifest file (.xml) on this device.**</span></span> <span data-ttu-id="71eaa-144">对于此选项，请选择“浏览”以找到想要使用的清单文件 (.xml)。</span><span class="sxs-lookup"><span data-stu-id="71eaa-144">For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.</span></span>
    - <span data-ttu-id="71eaa-p109">**我具有清单文件的 URL。** 对于此选项，请在提供的字段中键入清单的 URL。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p109">**I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.</span></span>

    !["新建Add-In"对话框Microsoft 365 管理中心。](../images/new-add-in.png)

8. <span data-ttu-id="71eaa-148">如果选择了此选项以从 Office 应用商店添加某个加载项，请选择该加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-148">If you selected the option to add an add-in from the Office Store, select the add-in.</span></span> <span data-ttu-id="71eaa-149">可以通过“**为你推荐**”、“**评级**”或“**名称**”类别，查看可用的加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-149">You can view available add-ins via categories of **Suggested for you**, **Rating**, or **Name**.</span></span> <span data-ttu-id="71eaa-150">仅能从 Office 应用商店添加免费加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-150">You may only add free add-ins from Office Store.</span></span> <span data-ttu-id="71eaa-151">目前不支持添加付费加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-151">Adding paid add-ins isn't currently supported.</span></span>

    > [!NOTE]
    > <span data-ttu-id="71eaa-152">使用 Office 应用商店选项，无需干预，用户即可自动获得加载项的更新和增强功能。</span><span class="sxs-lookup"><span data-stu-id="71eaa-152">With the Office Store option, updates and enhancements to the add-in are automatically available to users without your intervention.</span></span>

    ![Select an add-In dialog in Microsoft 365 管理中心.](../images/select-an-add-in.png)

9. <span data-ttu-id="71eaa-154">查看 **外接程序** 详细信息、隐私策略和许可条款后，选择"继续"。</span><span class="sxs-lookup"><span data-stu-id="71eaa-154">Choose **Continue** after reviewing the add-in details, Privacy Policy, and License Terms.</span></span>

    ![已选择外接程序页面中的Microsoft 365 管理中心。](../images/selected-add-in-admin-center.png)

10. <span data-ttu-id="71eaa-156">在"**分配用户"** 页上，选择"**任何人\*\*\*\*"、"特定用户/组**"或"**仅我"。**</span><span class="sxs-lookup"><span data-stu-id="71eaa-156">On the **Assign Users** page, choose **Everyone**, **Specific Users/Groups**, or **Only me**.</span></span> <span data-ttu-id="71eaa-157">使用“搜索”框查找要向其部署加载项的用户和组。</span><span class="sxs-lookup"><span data-stu-id="71eaa-157">Use the search box to find the users and groups to whom you want to deploy the add-in.</span></span> <span data-ttu-id="71eaa-158">For Outlook add-ins， you can also choose the deployment method **Fixed**， **Available**， or **Optional**.</span><span class="sxs-lookup"><span data-stu-id="71eaa-158">For Outlook add-ins, you can also choose the deployment method **Fixed**, **Available**, or **Optional**.</span></span>

    ![管理谁在部署中具有访问权限Microsoft 365 管理中心。](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > <span data-ttu-id="71eaa-160">使用 [SSO ](../develop/sso-in-office-add-ins.md) (单一登录) 将提示管理员同意外接程序清单中列出的范围。</span><span class="sxs-lookup"><span data-stu-id="71eaa-160">Add-ins that utilize [single sign-on (SSO)](../develop/sso-in-office-add-ins.md) will prompt the admin to consent to the scopes listed in the add-in manifest.</span></span>  <span data-ttu-id="71eaa-161">如果跨多个外接程序使用同一个支持服务 (则同一 Azure 应用 ID 用于不同外接程序) 中的 SSO，将提示每个外接程序的范围，以征得每个部署的同意。</span><span class="sxs-lookup"><span data-stu-id="71eaa-161">If the same backing service is used across multiple add-ins (the same Azure App ID is used with SSO in different add-ins), the scopes for each add-in will be prompted for consent with each deployment.</span></span> <span data-ttu-id="71eaa-162">此页面还将显示外接程序所需的权限列表。</span><span class="sxs-lookup"><span data-stu-id="71eaa-162">This page will also display the list of permissions that the add-in requires.</span></span>

11. <span data-ttu-id="71eaa-163">完成后，选择"部署 **"。**</span><span class="sxs-lookup"><span data-stu-id="71eaa-163">When finished, choose **Deploy**.</span></span> <span data-ttu-id="71eaa-164">此过程可能最多用时 3 分钟。</span><span class="sxs-lookup"><span data-stu-id="71eaa-164">This process may take up to three minutes.</span></span> <span data-ttu-id="71eaa-165">然后，按“**下一步**”完成演练。</span><span class="sxs-lookup"><span data-stu-id="71eaa-165">Then, finish the walkthrough by pressing **Next**.</span></span> <span data-ttu-id="71eaa-166">现在，你将看到加载项以及其他Office应用程序。</span><span class="sxs-lookup"><span data-stu-id="71eaa-166">You now see your add-in along with other Office apps.</span></span>

    > [!NOTE]
    > <span data-ttu-id="71eaa-167">当管理员选择"部署 **"** 时，将给予所有用户同意。</span><span class="sxs-lookup"><span data-stu-id="71eaa-167">When an administrator chooses **Deploy**, consent is given for all users.</span></span>

    ![应用程序中的应用Microsoft 365 管理中心。](../images/citations.png)

> [!TIP]
> <span data-ttu-id="71eaa-169">为组织中的用户和/或组部署新加载项时，请考虑向他们发送一封电子邮件，说明加载项的应用场景和使用方式，并添加相关帮助内容、FAQ 或其他支持资源的链接。</span><span class="sxs-lookup"><span data-stu-id="71eaa-169">When you deploy a new add-in to users and/or groups in your organization, consider sending them an email that describes when and how to use the add-in, and includes links to relevant Help content, FAQs, or other support resources.</span></span>

## <a name="considerations-when-granting-access-to-an-add-in"></a><span data-ttu-id="71eaa-170">授予加载项的访问权限时的注意事项</span><span class="sxs-lookup"><span data-stu-id="71eaa-170">Considerations when granting access to an add-in</span></span>

<span data-ttu-id="71eaa-p114">管理员可以将加载项分配给组织中的每个人或组织内的特定用户和/或组。 以下列表描述了每个选项的含义：</span><span class="sxs-lookup"><span data-stu-id="71eaa-p114">Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:</span></span>

- <span data-ttu-id="71eaa-p115">**每个人**：顾名思义，此选项为租户中的每位用户分配加载项。请谨慎使用此选项，且仅应用于真正在组织中通用的加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p115">**Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.</span></span>

- <span data-ttu-id="71eaa-p116">**用户**：如果将加载项分配给单个用户，则每次要将其分配给其他用户时，都需要更新此加载项的集中部署设置。 同样，每次要删除某个用户对该加载项的访问权限时，都需要更新该加载项的集中部署设置。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p116">**Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.</span></span>

- <span data-ttu-id="71eaa-177">**组**：如果将加载项分配给组，则会自动为被添加到此组的用户分配此加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-177">**Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in.</span></span> <span data-ttu-id="71eaa-178">同样，当将某个用户从组中删除时，此用户将自动失去对此加载项的访问权限。</span><span class="sxs-lookup"><span data-stu-id="71eaa-178">Likewise, when a user is removed from a group, the user automatically loses access to the add-in.</span></span> <span data-ttu-id="71eaa-179">在任一情况下，不需要管理员执行Microsoft 365操作。</span><span class="sxs-lookup"><span data-stu-id="71eaa-179">In either case, no additional action is required from the Microsoft 365 admin.</span></span>

<span data-ttu-id="71eaa-p118">一般情况下，为了便于维护，我们建议尽可能使用组来分配加载项。 但是，在想要将加载项的访问权限限制在极少数用户的情况下，将加载项分配给特定用户的做法可能更为实用。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p118">In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.</span></span>

## <a name="add-in-states"></a><span data-ttu-id="71eaa-182">加载项状态</span><span class="sxs-lookup"><span data-stu-id="71eaa-182">Add-in states</span></span>

<span data-ttu-id="71eaa-183">下表介绍了加载项的不同状态。</span><span class="sxs-lookup"><span data-stu-id="71eaa-183">The following table describes the different states of an add-in.</span></span>

|<span data-ttu-id="71eaa-184">状态</span><span class="sxs-lookup"><span data-stu-id="71eaa-184">State</span></span>|<span data-ttu-id="71eaa-185">此状态如何出现</span><span class="sxs-lookup"><span data-stu-id="71eaa-185">How the state occurs</span></span>|<span data-ttu-id="71eaa-186">影响</span><span class="sxs-lookup"><span data-stu-id="71eaa-186">Impact</span></span>|
|-----|--------------------|------|
|<span data-ttu-id="71eaa-187">**活动**</span><span class="sxs-lookup"><span data-stu-id="71eaa-187">**Active**</span></span>|<span data-ttu-id="71eaa-188">管理员已上传加载项并已将其分配给用户和/或组。</span><span class="sxs-lookup"><span data-stu-id="71eaa-188">Admin uploaded the add-in and assigned it to users and/or groups.</span></span>|<span data-ttu-id="71eaa-189">已为其分配加载项的用户和/或组，可在相关的 Office 客户端中找到它。</span><span class="sxs-lookup"><span data-stu-id="71eaa-189">Users and/or groups assigned to the add-in see it in the relevant Office clients.</span></span>|
|<span data-ttu-id="71eaa-190">**已禁用**</span><span class="sxs-lookup"><span data-stu-id="71eaa-190">**Turned off**</span></span>|<span data-ttu-id="71eaa-191">管理员已禁用加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-191">Admin turned off the add-in.</span></span>|<span data-ttu-id="71eaa-p119">已为其分配加载项的用户和/或组不再能够访问它。 如果加载项状态从“已禁用”\*\*\*\* 更改为“活动”\*\*\*\*，则用户和组将重新获得对它的访问权限。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p119">Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.</span></span>|
|<span data-ttu-id="71eaa-194">**已删除**</span><span class="sxs-lookup"><span data-stu-id="71eaa-194">**Deleted**</span></span>|<span data-ttu-id="71eaa-195">管理员已删除加载项。</span><span class="sxs-lookup"><span data-stu-id="71eaa-195">Admin deleted the add-in.</span></span>|<span data-ttu-id="71eaa-196">已为其分配加载项的用户和/或组不再能够访问它。</span><span class="sxs-lookup"><span data-stu-id="71eaa-196">Users and/or groups assigned the add-in no longer have access to it.</span></span>|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a><span data-ttu-id="71eaa-197">更新通过集中部署发布的 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="71eaa-197">Updating Office Add-ins that are published via Centralized Deployment</span></span>

<span data-ttu-id="71eaa-198">在Office集中部署发布外接程序后，在 Web 应用程序中实现对外接程序 Web 应用程序的任何更改后，所有用户将自动获得这些更改。</span><span class="sxs-lookup"><span data-stu-id="71eaa-198">After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users after those changes are implemented in the web application.</span></span> <span data-ttu-id="71eaa-199">对外接程序 [的 XML](../develop/add-in-manifests.md) 清单文件所做的更改（例如，更新外接程序的图标、文本或外接程序命令）发生如下：</span><span class="sxs-lookup"><span data-stu-id="71eaa-199">Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md) to, for example, update the add-in's icon, text, or add-in commands, happen as follows:</span></span>

- <span data-ttu-id="71eaa-200">业务线外接程序：如果管理员在通过 Microsoft 365 管理中心 实现集中部署时，从设备或通过指向 URL) 显式上传了清单文件 (，则管理员必须上载包含所需更改的新清单文件。</span><span class="sxs-lookup"><span data-stu-id="71eaa-200">**Line-of-business add-in**: If an admin explicitly uploaded a manifest file (either from their device or by pointing to a URL) when implementing Centralized Deployment via the Microsoft 365 admin center, the admin must upload a new manifest file that contains the desired changes.</span></span> <span data-ttu-id="71eaa-201">上传更新后的清单文件后，加载项就会在下次相关 Office 应用启动时更新。</span><span class="sxs-lookup"><span data-stu-id="71eaa-201">After the updated manifest file has been uploaded, the next time the relevant Office applications start, the add-in will update.</span></span>

  > [!NOTE]
  > <span data-ttu-id="71eaa-202">管理员无需删除 LOB 加载项来进行更新。</span><span class="sxs-lookup"><span data-stu-id="71eaa-202">An admin doesn't need to remove a LOB add-in to make an update.</span></span> <span data-ttu-id="71eaa-203">在"外接程序"部分，管理员只需按右下角的"更新外接程序"按钮，即可选择 LOB 外接程序并调用此功能。</span><span class="sxs-lookup"><span data-stu-id="71eaa-203">In the Add-ins section, the admin can simply choose the LOB add-in and invoke this functionality by pressing the **Update add-in** button present in the bottom right corner.</span></span>
  >
  > ![Screenshot shows the Update add-in dialog in Microsoft 365 管理中心.](../images/update-add-in-admin-center.png)

- <span data-ttu-id="71eaa-205">**Office** 应用商店外接程序：如果管理员在通过 Microsoft 365 管理中心 实现集中部署时从 Office 应用商店选择了外接程序，并且 Office 应用商店中的外接程序更新，则外接程序稍后将通过集中部署更新。</span><span class="sxs-lookup"><span data-stu-id="71eaa-205">**Office Store add-in**: If an admin selected an add-in from the Office Store when implementing Centralized Deployment via the Microsoft 365 admin center, and the add-in updates in the Office Store, the add-in will update later via Centralized Deployment.</span></span> <span data-ttu-id="71eaa-206">需要最多 24 小时，应用商店外接程序更新才能针对所有最终用户流动。</span><span class="sxs-lookup"><span data-stu-id="71eaa-206">It can take up to 24 hours for the Store add-in updates to flow for all end users.</span></span> <span data-ttu-id="71eaa-207">在此持续时间后，下一次Office用户重新启动相关应用程序时，外接程序将更新。</span><span class="sxs-lookup"><span data-stu-id="71eaa-207">After this duration, the next time the relevant Office applications restart for these users, the add-in will update.</span></span> <span data-ttu-id="71eaa-208">用户还可以触发手动刷新，以通过选择插入选项卡加载项管理托管选项卡点击刷新获取最新的应用商店  >    >    >  **外接程序版本**。</span><span class="sxs-lookup"><span data-stu-id="71eaa-208">Users can also trigger a Manual Refresh to get the latest Store add-in version by selecting **Insert Tab** > **Add-ins** > **Admin Managed Tab** > **Hit Refresh**.</span></span>

## <a name="end-user-experience-with-add-ins"></a><span data-ttu-id="71eaa-209">加载项最终用户体验</span><span class="sxs-lookup"><span data-stu-id="71eaa-209">End user experience with add-ins</span></span>

<span data-ttu-id="71eaa-210">通过集中部署发布加载项后，最终用户可以在加载项支持的任何平台上开始使用它。</span><span class="sxs-lookup"><span data-stu-id="71eaa-210">After an add-in has been published via Centralized Deployment, end users may start using it on any platform that the add-in supports.</span></span>

<span data-ttu-id="71eaa-p124">如果外接程序支持外接程序命令，则这些命令将出现在为其部署外接程序的所有用户的 Office 应用程序功能区上。 在以下的示例中，**搜索引文** 命令将显示在 **引文** 加载项的功能区上。</span><span class="sxs-lookup"><span data-stu-id="71eaa-p124">If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.</span></span>

![Screenshot shows a section of the Office 应用 ribbon with the Search Citation command highlighted in the Citations add-in.](../images/search-citation.png)

<span data-ttu-id="71eaa-214">如果加载项不支持加载项命令，用户可以通过执行以下操作将其添加到 Office 应用程序中：</span><span class="sxs-lookup"><span data-stu-id="71eaa-214">If the add-in does not support add-in commands, users can add it to their Office application by doing the following:</span></span>

1. <span data-ttu-id="71eaa-215">在 Word 2016 或更高版本、Excel 2016 或更高版本，或 PowerPoint 2016 或更高版本，选择“**插入**” > “**我的加载项**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-215">In Word 2016 or later, Excel 2016 or later, or PowerPoint 2016 or later, choose **Insert** > **My Add-ins**.</span></span>
2. <span data-ttu-id="71eaa-216">在加载项窗口中选择“**管理托管**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="71eaa-216">Choose the **Admin Managed** tab in the add-in window.</span></span>
3. <span data-ttu-id="71eaa-217">选择加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-217">Choose the add-in, and then choose **Add**.</span></span>

    ![屏幕截图显示 Office 应用程序的“Office 加载项”页的“管理托管”选项卡。 引文加载项显示在此选项卡上。](../images/office-add-ins-admin-managed.png)

<span data-ttu-id="71eaa-220">但是，对于 Outlook 2016 或更高版本，用户可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="71eaa-220">However, for Outlook 2016 or later, users can do the following:</span></span>

1. <span data-ttu-id="71eaa-221">在 Outlook 中，选择“**开始**” > “**应用商店**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-221">In Outlook, choose **Home** > **Store**.</span></span>
2. <span data-ttu-id="71eaa-222">选择“加载项”选项卡下的“**管理员管理**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="71eaa-222">Choose the **Admin-managed** item under the add-in tab.</span></span>
3. <span data-ttu-id="71eaa-223">选择加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="71eaa-223">Choose the add-in, and then choose **Add**.</span></span>

    ![屏幕截图显示了 Outlook 应用程序的“应用商店”页面的管理员管理区域。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a><span data-ttu-id="71eaa-225">另请参阅</span><span class="sxs-lookup"><span data-stu-id="71eaa-225">See also</span></span>

- [<span data-ttu-id="71eaa-226">确定加载项的集中部署是否适用于你的 Microsoft 365 组织</span><span class="sxs-lookup"><span data-stu-id="71eaa-226">Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization</span></span>](/office365/admin/manage/centralized-deployment-of-add-ins)
