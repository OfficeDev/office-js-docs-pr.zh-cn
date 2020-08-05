---
title: Outlook 加载项的隐私、权限和安全性
description: 了解如何管理 Outlook 加载项中的隐私、权限和安全性。
ms.date: 08/03/2020
localization_priority: Priority
ms.openlocfilehash: 9807cbb2346d6fc067f3894c9f5d265f83dccdc3
ms.sourcegitcommit: a3b743598025466bad19177e0ba9ca94ea66d490
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/04/2020
ms.locfileid: "46547533"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a><span data-ttu-id="7dd8f-103">Outlook 外接程序的隐私、权限和安全性</span><span class="sxs-lookup"><span data-stu-id="7dd8f-103">Privacy, permissions, and security for Outlook add-ins</span></span>

<span data-ttu-id="7dd8f-104">最终用户、开发人员和管理员可以使用 Outlook 外接程序的安全模型的分层权限级别来控制隐私和性能。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-104">End users, developers, and administrators can use the tiered permission levels of the security model for Outlook add-ins to control privacy and performance.</span></span>

<span data-ttu-id="7dd8f-105">本文介绍了 Outlook 加载项可以请求的可能权限，并从以下几个角度审视安全模型：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-105">This article describes the possible permissions that Outlook add-ins can request, and examines the security model from the following perspectives:</span></span>

- <span data-ttu-id="7dd8f-106">**AppSource**：加载项完整性</span><span class="sxs-lookup"><span data-stu-id="7dd8f-106">**AppSource**: add-in integrity</span></span>
    
- <span data-ttu-id="7dd8f-107">**最终用户**：隐私和性能问题</span><span class="sxs-lookup"><span data-stu-id="7dd8f-107">**End-users**: privacy and performance concerns</span></span>
    
- <span data-ttu-id="7dd8f-108">**开发人员**：权限选择和资源使用限制。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-108">**Developers**: permissions choices and resource usage limits</span></span>
    
- <span data-ttu-id="7dd8f-109">**管理员**：设置性能阈值的权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-109">**Administrators**: privileges to set performance thresholds</span></span>
    

## <a name="permissions-model"></a><span data-ttu-id="7dd8f-110">权限模型</span><span class="sxs-lookup"><span data-stu-id="7dd8f-110">Permissions model</span></span>

<span data-ttu-id="7dd8f-p101">客户对外接程序安全的理解可能会影响外接程序采用情况，因此 Outlook 外接程序安全依赖于一个多层权限模型。Outlook 外接程序可能会公开其所需的权限级别，从而确定外接程序可以对客户邮箱数据采取的可能访问和操作。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-p101">Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.</span></span> 

<span data-ttu-id="7dd8f-113">清单架构版本 1.1 包含四个级别的权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-113">Manifest schema version 1.1 includes four levels of permissions.</span></span> 


<span data-ttu-id="7dd8f-114">**表 1.外接程序权限级别**</span><span class="sxs-lookup"><span data-stu-id="7dd8f-114">**Table 1. Add-in permission levels**</span></span>

|<span data-ttu-id="7dd8f-115">**权限级别**</span><span class="sxs-lookup"><span data-stu-id="7dd8f-115">**Permission level**</span></span>|<span data-ttu-id="7dd8f-116">**Outlook 外接程序清单中的值**</span><span class="sxs-lookup"><span data-stu-id="7dd8f-116">**Value in Outlook add-in manifest**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7dd8f-117">受限</span><span class="sxs-lookup"><span data-stu-id="7dd8f-117">Restricted</span></span>|<span data-ttu-id="7dd8f-118">受限</span><span class="sxs-lookup"><span data-stu-id="7dd8f-118">Restricted</span></span>|
|<span data-ttu-id="7dd8f-119">读取项目</span><span class="sxs-lookup"><span data-stu-id="7dd8f-119">Read item</span></span>|<span data-ttu-id="7dd8f-120">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7dd8f-120">ReadItem</span></span>|
|<span data-ttu-id="7dd8f-121">读/写项目</span><span class="sxs-lookup"><span data-stu-id="7dd8f-121">Read/write item</span></span>|<span data-ttu-id="7dd8f-122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7dd8f-122">ReadWriteItem</span></span>|
|<span data-ttu-id="7dd8f-123">读/写邮箱</span><span class="sxs-lookup"><span data-stu-id="7dd8f-123">Read/write mailbox</span></span>|<span data-ttu-id="7dd8f-124">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="7dd8f-124">ReadWriteMailbox</span></span>|

<span data-ttu-id="7dd8f-125">四个级别的权限具有累积性：**读/写邮箱**权限包括**读/写项**权限、**读取项**权限和**受限**权限；**读/写项**权限包括**读取项**权限和**受限**权限；**读取项**权限包括**受限**权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-125">The four levels of permissions are cumulative: the **read/write mailbox** permission includes the permissions of **read/write item**, **read item** and **restricted**, **read/write item** includes **read item** and **restricted**, and the **read item** permission includes **restricted**.</span></span> 

<span data-ttu-id="7dd8f-126">下图显示了四个级别的权限并说明了每一层提供给最终用户、开发人员和管理员的功能。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-126">The following figure shows the four levels of permissions and describes the capabilities offered to the end user, developer, and administrator by each tier.</span></span> <span data-ttu-id="7dd8f-127">有关这些权限的详细信息，请参阅 [最终用户：隐私和性能问题](#end-users-privacy-and-performance-concerns)、[开发人员：权限选择和资源使用限制](#developers-permission-choices-and-resource-usage-limits) 和[了解 Outlook 加载项权限](understanding-outlook-add-in-permissions.md)。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-127">For more information about these permissions, see [End users: privacy and performance concerns](#end-users-privacy-and-performance-concerns), [Developers: permission choices and resource usage limits](#developers-permission-choices-and-resource-usage-limits), and [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span> 


<span data-ttu-id="7dd8f-128">**将四层权限模型与最终用户、开发人员和管理员关联**</span><span class="sxs-lookup"><span data-stu-id="7dd8f-128">**Relating the four-tier permission model to the end user, developer, and administrator**</span></span>

![邮件应用程序架构 v1.1 的 4 层权限模型](../images/add-in-permission-tiers.png)


## <a name="appsource-add-in-integrity"></a><span data-ttu-id="7dd8f-130">AppSource：加载项完整性</span><span class="sxs-lookup"><span data-stu-id="7dd8f-130">AppSource: add-in integrity</span></span>

<span data-ttu-id="7dd8f-131">[AppSource](https://appsource.microsoft.com) 托管可由最终用户和管理员安装的加载项。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-131">[AppSource](https://appsource.microsoft.com) hosts add-ins that can be installed by end users and administrators.</span></span> <span data-ttu-id="7dd8f-132">AppSource 强制执行以下措施来维护这些 Outlook 加载项的完整性：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-132">AppSource enforces the following measures to maintain the integrity of these Outlook add-ins:</span></span>

- <span data-ttu-id="7dd8f-133">要求加载项的主机服务器始终使用安全套接字层 (SSL) 进行通信。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-133">Requires the host server of an add-in to always use Secure Socket Layer (SSL) to communicate.</span></span>
    
- <span data-ttu-id="7dd8f-134">要求开发人员在提交加载项时提供身份证明、合约协议和适合的隐私策略。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-134">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span> 
    
- <span data-ttu-id="7dd8f-135">以只读模式存档加载项。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-135">Archives add-ins in read-only mode.</span></span>
    
- <span data-ttu-id="7dd8f-136">支持针对可用加载项的用户审阅系统以推广自我管理的社区。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-136">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>
    

## <a name="end-users-privacy-and-performance-concerns"></a><span data-ttu-id="7dd8f-137">最终用户：隐私和性能问题</span><span class="sxs-lookup"><span data-stu-id="7dd8f-137">End users: privacy and performance concerns</span></span>

<span data-ttu-id="7dd8f-138">安全模型通过下列方式解决最终用户的安全、隐私和性能问题：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-138">The security model addresses security, privacy, and performance concerns of end users in the following ways:</span></span>

- <span data-ttu-id="7dd8f-139">受 Outlook 信息权限管理 (IRM) 保护的最终用户邮件不与 Outlook 外接程序交互。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-139">End user's messages that are protected by Outlook's Information Rights Management (IRM) do not interact with Outlook add-ins.</span></span>
    
  > [!IMPORTANT]
  > <span data-ttu-id="7dd8f-140">现在，Windows 版 Outlook 从内部版本 13120.1000 开始可以在受 IRM 保护的项目上激活加载项。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-140">Starting with Outlook build 13120.1000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="7dd8f-141">有关处于预览阶段的此功能的详细信息，请参阅[在受信息权限管理 (IRM) 保护的项目上激活加载项](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-141">For more information about this feature in preview, see [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="7dd8f-142">从 AppSource 安装加载项之前，最终用户能够查看加载项可以对其数据进行的访问和采取的操作，且必须明确确认后才能继续操作。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-142">Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed.</span></span> <span data-ttu-id="7dd8f-143">未经用户或管理员手动验证，Outlook 外接程序不会自动推送到客户端计算机。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-143">No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.</span></span>
    
- <span data-ttu-id="7dd8f-p106">授予“受限”权限可允许 Outlook 外接程序仅具有对当前项目的有限访问权限。授予“读取项目”权限可允许 Outlook 外接程序仅访问当前项目上的个人识别信息，例如发件人和收件人姓名以及电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-p106">Granting the **restricted** permission allows the Outlook add-in to have limited access on only the current item. Granting the **read item** permission allows the Outlook add-in to access personal identifiable information, such as sender and recipient names and email addresses, on only the current item,.</span></span>
    
- <span data-ttu-id="7dd8f-p107">最终用户仅能为他/她自己安装低信任度的 Outlook 外接程序。对组织产生影响的 Outlook 外接程序由管理员安装。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-p107">An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.</span></span>
    
- <span data-ttu-id="7dd8f-148">最终用户可以安装支持上下文相关方案的低信任度 Outlook 外接程序，这不仅对用户具有吸引力，同时还可以最大限度地降低用户的安全风险。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-148">End users can install Outlook add-ins that enable context-sensitive scenarios that are compelling to users while minimizing the users' security risks.</span></span>
    
- <span data-ttu-id="7dd8f-149">已安装 Outlook 外接程序的清单文件在用户电子邮件帐户中受到保护。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-149">Manifest files of installed Outlook add-ins are secured in the user's email account.</span></span>
    
- <span data-ttu-id="7dd8f-150">通过托管 Office 外接程序的服务器传送的数据始终根据安全套接字层 (SSL) 协议进行加密。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-150">Data communicated with servers hosting Office Add-ins is always encrypted according to the Secure Socket Layer (SSL) protocol.</span></span>
    
- <span data-ttu-id="7dd8f-151">仅适用于 Outlook 富客户端：Outlook 富客户端监视已安装 Outlook 外接程序的性能，实施管治控制，以及禁用在以下方面超过限制的 Outlook 外接程序：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-151">Applicable to only the Outlook rich clients: The Outlook rich clients monitor the performance of installed Outlook add-ins, exercise governance control, and disable those Outlook add-ins that exceed limits in the following areas:</span></span>
    
  - <span data-ttu-id="7dd8f-152">激活响应时间</span><span class="sxs-lookup"><span data-stu-id="7dd8f-152">Response time to activate</span></span>
    
  - <span data-ttu-id="7dd8f-153">激活或重新激活失败次数</span><span class="sxs-lookup"><span data-stu-id="7dd8f-153">Number of failures to activate or reactivate</span></span>
    
  - <span data-ttu-id="7dd8f-154">内存使用率</span><span class="sxs-lookup"><span data-stu-id="7dd8f-154">Memory usage</span></span>
    
  - <span data-ttu-id="7dd8f-155">CPU 使用率</span><span class="sxs-lookup"><span data-stu-id="7dd8f-155">CPU usage</span></span>  

  <span data-ttu-id="7dd8f-p108">管治可阻止拒绝服务攻击并将外接程序性能保持在合理的水平。业务栏通知最终用户 Outlook 富客户端已根据此类管治控制禁用的 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-p108">Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.</span></span>

- <span data-ttu-id="7dd8f-158">无论何时，最终用户都可以验证所安装 Outlook 外接程序请求的权限，在 Exchange 管理中心禁用或随后启用任何 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-158">At any time, end users can verify the permissions requested by installed Outlook add-ins, and disable or subsequently enable any Outlook add-in in the Exchange Admin Center.</span></span>


## <a name="developers-permission-choices-and-resource-usage-limits"></a><span data-ttu-id="7dd8f-159">开发人员：权限选择和资源使用限制</span><span class="sxs-lookup"><span data-stu-id="7dd8f-159">Developers: permission choices and resource usage limits</span></span>

<span data-ttu-id="7dd8f-160">安全模型向开发人员提供精细级别的权限以供选择，以及严格的性能准则以供遵循。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-160">The security model provides developers granular levels of permissions to choose from, and strict performance guidelines to observe.</span></span>

### <a name="tiered-permissions-increases-transparency"></a><span data-ttu-id="7dd8f-161">多层权限将增加透明度</span><span class="sxs-lookup"><span data-stu-id="7dd8f-161">Tiered permissions increases transparency</span></span>

<span data-ttu-id="7dd8f-162">开发人员应按照多层权限模型提供透明度，并解决用户有关哪些加载项可以处理其数据和邮箱的问题，间接促进加载项采用：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-162">Developers should follow the tiered permissions model to provide transparency and alleviate users' concern about what add-ins can do to their data and mailbox, indirectly promoting add-in adoption:</span></span>

- <span data-ttu-id="7dd8f-163">开发人员根据 Outlook 外接程序应激活的方式、Outlook 外接程序读取或写入项目特定属性的需求，或者创建和发送项目的需求来针对 Outlook 外接程序请求适当级别的权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-163">Developers request an appropriate level of permission for an Outlook add-in, based on how the Outlook add-in should be activated, and its need to read or write certain properties of an item, or to create and send an item.</span></span>

- <span data-ttu-id="7dd8f-164">开发人员使用 Outlook 加载项清单中的 [Permissions](../reference/manifest/permissions.md) 元素，并根据需要分配 **Restricted**、**ReadItem**、**ReadWriteItem** 或 **ReadWriteMailbox** 的值来请求权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-164">Developers request permission by using the [Permissions](../reference/manifest/permissions.md) element in the manifest of the Outlook add-in, by assigning a value of **Restricted**, **ReadItem**, **ReadWriteItem** or **ReadWriteMailbox**, as appropriate.</span></span>

  > [!NOTE]
  > <span data-ttu-id="7dd8f-165">请注意，从清单架构 v1.1 开始就提供 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-165">Note that the **ReadWriteItem** permission is available starting in manifest schema v1.1.</span></span>

  <span data-ttu-id="7dd8f-166">下面的示例请求**读取项**权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-166">The following example requests the **read item** permission.</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- <span data-ttu-id="7dd8f-167">如果 Outlook 加载项激活特定类型的 Outlook 项目（约会或邮件）或存在于项目主题或正文中的特定提取的实体（电话号码、地址、URL），开发人员可以请求“**受限**”权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-167">Developers can request the **restricted** permission if the Outlook add-in activates on a specific type of Outlook items (appointment or message), or on specific extracted entities (phone number, address, URL) being present in the item's subject or body.</span></span> <span data-ttu-id="7dd8f-168">例如，如果在当前邮件的主题或正文中找到一个或多个实体（共三个）- 电话号码、邮寄地址或 URL，以下规则将激活 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-168">For example, the following rule activates the Outlook add-in if one or more of three entities - phone number, postal address, or URL - are found in the subject or body of the current message.</span></span>
    
  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- <span data-ttu-id="7dd8f-169">如果 Outlook 加载项需要读取当前项目的属性而非默认提取实体的属性，或者需要通过当前项目上的加载项写入自定义属性集，但无需读写其他项目或在用户的邮箱中创建或发送邮件，则开发人员应请求“**读取项目**”权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-169">Developers should request the **read item** permission if the Outlook add-in needs to read properties of the current item other than the default extracted entities, or write custom properties set by the add-in on the current item, but does not require reading or writing to other items, or creating or sending a message in the user's mailbox.</span></span> <span data-ttu-id="7dd8f-170">例如，如果 Outlook 外接程序需要寻找项目主体或正文中的会议建议、任务建议、电子邮件地址或联系人姓名等实体，或者需要使用一个正则表达式来激活，则开发人员应请求“**读取项目**”权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-170">For example, a developer should request **read item** permission if an Outlook add-in needs to look for an entity like a meeting suggestion, task suggestion, email address, or contact name in the item's subject or body, or uses a regular expression to activate.</span></span>

- <span data-ttu-id="7dd8f-171">如果 Outlook 加载项需要向撰写的项目的属性（如收件人姓名、电子邮件地址、正文和主题）写入，或需要添加或删除项目附件，那么开发人员应请求“**读/写项目**”权限。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-171">Developers should request the **read/write item** permission if the Outlook add-in needs to write to properties of the composed item, such as recipient names, email addresses, body, and subject, or needs to add or remove item attachments.</span></span>

- <span data-ttu-id="7dd8f-172">仅在 Outlook 外接程序需要使用 [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法执行下列一个或多个操作时，开发人员才请求“读/写邮箱”权限：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-172">Developers request the **read/write mailbox** permission only if the Outlook add-in needs to do one or more of the following actions by using the [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method:</span></span>

  - <span data-ttu-id="7dd8f-173">读取或写入邮箱中项目的属性。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-173">Read or write to properties of items in the mailbox.</span></span>
  - <span data-ttu-id="7dd8f-174">创建、读取、写入或发送邮箱中的项目。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-174">Create, read, write, or send items in the mailbox.</span></span>
  - <span data-ttu-id="7dd8f-175">创建、读取或写入邮箱文件夹。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-175">Create, read, or write to folders in the mailbox.</span></span>


### <a name="resource-usage-tuning"></a><span data-ttu-id="7dd8f-176">资源使用调整</span><span class="sxs-lookup"><span data-stu-id="7dd8f-176">Resource usage tuning</span></span>

<span data-ttu-id="7dd8f-p111">开发人员应注意激活资源的使用限制，在他们的开发工作流中加入性能调整功能，以便减少主机对低性能外接程序的拒绝服务机会。开发人员应遵循 [Outlook 外接程序的激活和 JavaScript API 的限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)中所述的设计激活规则准则。如果 Outlook 外接程序适合运行于 Outlook 富客户端之上，那么开发人员应验证该外接程序能否在资源使用限制之内执行。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-p111">Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.</span></span>


### <a name="other-measures-to-promote-user-security"></a><span data-ttu-id="7dd8f-179">提高用户安全性的其他措施</span><span class="sxs-lookup"><span data-stu-id="7dd8f-179">Other measures to promote user security</span></span>

<span data-ttu-id="7dd8f-180">开发人员还应该注意并规划以下内容：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-180">Developers should be aware of and plan for the following as well:</span></span>

- <span data-ttu-id="7dd8f-181">开发人员无法在加载项中使用 ActiveX 控件，因为它们不受支持。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-181">Developers cannot use ActiveX controls in add-ins because they are not supported.</span></span>
    
- <span data-ttu-id="7dd8f-182">开发人员应在将 Outlook 加载项提交到 AppSource 时执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-182">Developers should do the following when submitting an Outlook add-in to AppSource:</span></span>
    
  - <span data-ttu-id="7dd8f-183">生成扩展验证 (EV) SSL 证书作为身份证明。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-183">Produce an Extended Validation (EV) SSL certificate as a proof of identity.</span></span>
    
  - <span data-ttu-id="7dd8f-184">在支持 SSL 的 Web 服务器上承载其提交的加载项。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-184">Host the add-in they are submitting on a web server that supports SSL.</span></span>
    
  - <span data-ttu-id="7dd8f-185">生成合规隐私策略。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-185">Produce a compliant privacy policy.</span></span>
    
  - <span data-ttu-id="7dd8f-186">准备好在提交加载项后签订合约协议。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-186">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>
    

## <a name="administrators-privileges"></a><span data-ttu-id="7dd8f-187">管理员：权限</span><span class="sxs-lookup"><span data-stu-id="7dd8f-187">Administrators: privileges</span></span>

<span data-ttu-id="7dd8f-188">安全模型向管理员提供以下权限和责任：</span><span class="sxs-lookup"><span data-stu-id="7dd8f-188">The security model provides the following rights and responsibilities to administrators:</span></span>

- <span data-ttu-id="7dd8f-189">可以阻止最终用户安装任何 Outlook 加载项，包括来自 AppSource 的加载项。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-189">Can prevent end users from installing any Outlook add-in, including add-ins from AppSource.</span></span>
    
- <span data-ttu-id="7dd8f-190">可以在 Exchange 管理中心上禁用或启用任何 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-190">Can disable or enable any Outlook add-in on the Exchange Admin Center.</span></span>
    
- <span data-ttu-id="7dd8f-191">仅适用于 Windows 版 Outlook：可以通过 GPO 注册表设置覆盖性能阈值设置。</span><span class="sxs-lookup"><span data-stu-id="7dd8f-191">Applicable to only Outlook on Windows: Can override performance threshold settings by GPO registry settings.</span></span>
    


## <a name="see-also"></a><span data-ttu-id="7dd8f-192">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7dd8f-192">See also</span></span>

- [<span data-ttu-id="7dd8f-193">Office 加载项的隐私和安全性</span><span class="sxs-lookup"><span data-stu-id="7dd8f-193">Privacy and security for Office Add-ins</span></span>](../develop/privacy-and-security.md)    
- [<span data-ttu-id="7dd8f-194">Outlook 外接程序 API</span><span class="sxs-lookup"><span data-stu-id="7dd8f-194">Outlook add-in APIs</span></span>](apis.md)    
- [<span data-ttu-id="7dd8f-195">Outlook 外接程序的激活和 JavaScript API 限制</span><span class="sxs-lookup"><span data-stu-id="7dd8f-195">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
