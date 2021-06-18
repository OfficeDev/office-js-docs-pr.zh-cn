---
title: 在加载项中启用共享文件夹Outlook邮箱方案
description: 讨论如何为共享文件夹配置外接程序支持 (。例如， 委派访问) 和共享邮箱。
ms.date: 06/17/2021
localization_priority: Normal
ms.openlocfilehash: 5d7fb712b8f814184c2a444c32416d35fb1da49c
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007767"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="f2904-104">在加载项中启用共享文件夹Outlook邮箱方案</span><span class="sxs-lookup"><span data-stu-id="f2904-104">Enable shared folders and shared mailbox scenarios in an Outlook add-in</span></span>

<span data-ttu-id="f2904-105">本文介绍如何在 Outlook 外接程序中启用共享文件夹 (也称为委派访问) 和共享邮箱 (（预览[) ](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)方案，包括 Office JavaScript API 支持哪些权限）。</span><span class="sxs-lookup"><span data-stu-id="f2904-105">This article describes how to enable shared folders (also known as delegate access) and shared mailbox (now in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)) scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f2904-106">要求集 [1.8 中引入了对此功能的支持](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)。</span><span class="sxs-lookup"><span data-stu-id="f2904-106">Support for this feature was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="f2904-107">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="f2904-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-setups"></a><span data-ttu-id="f2904-108">支持的安装程序</span><span class="sxs-lookup"><span data-stu-id="f2904-108">Supported setups</span></span>

<span data-ttu-id="f2904-109">以下各节介绍共享邮箱和共享文件夹 (预览) 的配置。</span><span class="sxs-lookup"><span data-stu-id="f2904-109">The following sections describe supported configurations for shared mailboxes (now in preview) and shared folders.</span></span> <span data-ttu-id="f2904-110">在其他配置中，功能 API 可能无法如预期工作。</span><span class="sxs-lookup"><span data-stu-id="f2904-110">The feature APIs may not work as expected in other configurations.</span></span> <span data-ttu-id="f2904-111">选择要了解如何配置的平台。</span><span class="sxs-lookup"><span data-stu-id="f2904-111">Select the platform you'd like to learn how to configure.</span></span>

### <a name="windows"></a>[<span data-ttu-id="f2904-112">Windows</span><span class="sxs-lookup"><span data-stu-id="f2904-112">Windows</span></span>](#tab/windows)

#### <a name="shared-folders"></a><span data-ttu-id="f2904-113">共享文件夹</span><span class="sxs-lookup"><span data-stu-id="f2904-113">Shared folders</span></span>

<span data-ttu-id="f2904-114">邮箱所有者必须先 [向代理提供访问权限](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="f2904-114">The mailbox owner must first [provide access to a delegate](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="f2904-115">然后，代理人必须遵循管理其他人的邮件和日历项目一文的"将其他人的邮箱添加到你的配置文件"部分中 [概述的说明](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)。</span><span class="sxs-lookup"><span data-stu-id="f2904-115">The delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="f2904-116">共享邮箱 (预览) </span><span class="sxs-lookup"><span data-stu-id="f2904-116">Shared mailboxes (preview)</span></span>

<span data-ttu-id="f2904-117">Exchange管理员可创建和管理共享邮箱，供多组用户访问。</span><span class="sxs-lookup"><span data-stu-id="f2904-117">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="f2904-118">目前[，Exchange Online](/exchange/collaboration-exo/shared-mailboxes)是此功能唯一受支持的服务器版本。</span><span class="sxs-lookup"><span data-stu-id="f2904-118">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="f2904-119">默认情况下Exchange Server自动映射"功能是启用的，这意味着共享邮箱随后应在关闭并重新打开共享邮箱后自动[](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)显示在用户的 Outlook Outlook 应用中。</span><span class="sxs-lookup"><span data-stu-id="f2904-119">An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened.</span></span> <span data-ttu-id="f2904-120">但是，如果管理员关闭自动映射，用户必须按照在 Outlook 中打开和使用共享邮箱一文的"将共享邮箱添加到 Outlook"部分中概述的[手动](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd)步骤操作。</span><span class="sxs-lookup"><span data-stu-id="f2904-120">However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span></span>

> [!WARNING]
> <span data-ttu-id="f2904-121">请勿 **使用** 密码登录共享邮箱。</span><span class="sxs-lookup"><span data-stu-id="f2904-121">Do **NOT** sign into the shared mailbox with a password.</span></span> <span data-ttu-id="f2904-122">在这种情况下，功能 API 将不起作用。</span><span class="sxs-lookup"><span data-stu-id="f2904-122">The feature APIs won't work in that case.</span></span>

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="f2904-123">Web 浏览器 - 新式 Outlook</span><span class="sxs-lookup"><span data-stu-id="f2904-123">Web browser - modern Outlook</span></span>](#tab/modern)

#### <a name="shared-folders"></a><span data-ttu-id="f2904-124">共享文件夹</span><span class="sxs-lookup"><span data-stu-id="f2904-124">Shared folders</span></span>

<span data-ttu-id="f2904-125">邮箱所有者必须先 [通过更新邮箱文件夹](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) 权限向代理提供访问权限。</span><span class="sxs-lookup"><span data-stu-id="f2904-125">The mailbox owner must first [provide access to a delegate](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) by updating the mailbox folder permissions.</span></span> <span data-ttu-id="f2904-126">然后，代理必须遵循文章访问其他人的邮箱 的"将其他人的邮箱添加到 Outlook Web App 中的文件夹列表"部分中概述[的说明](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081)。</span><span class="sxs-lookup"><span data-stu-id="f2904-126">The delegate must then follow the instructions outlined in the "Add another person’s mailbox to your folder list in Outlook Web App" section of the article [Access another person's mailbox](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="f2904-127">共享邮箱 (预览) </span><span class="sxs-lookup"><span data-stu-id="f2904-127">Shared mailboxes (preview)</span></span>

<span data-ttu-id="f2904-128">Exchange管理员可创建和管理共享邮箱，供多组用户访问。</span><span class="sxs-lookup"><span data-stu-id="f2904-128">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="f2904-129">目前[，Exchange Online](/exchange/collaboration-exo/shared-mailboxes)是此功能唯一受支持的服务器版本。</span><span class="sxs-lookup"><span data-stu-id="f2904-129">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="f2904-130">获得访问权限后，共享邮箱用户必须遵循在"在邮箱中打开和使用共享邮箱"一文的"添加共享邮箱，以便它显示在主邮箱[下"一节中Outlook 网页版。](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207)</span><span class="sxs-lookup"><span data-stu-id="f2904-130">After receiving access, a shared mailbox user must follow the steps outlined in the "Add the shared mailbox so it displays under your primary mailbox" section of the article [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span></span>

> [!WARNING]
> <span data-ttu-id="f2904-131">请勿 **使用** "打开另一个邮箱"等其他选项。</span><span class="sxs-lookup"><span data-stu-id="f2904-131">Do **NOT** use other options like "Open another mailbox".</span></span> <span data-ttu-id="f2904-132">然后，功能 API 可能无法正常运行。</span><span class="sxs-lookup"><span data-stu-id="f2904-132">The feature APIs may not work properly then.</span></span>

---

<span data-ttu-id="f2904-133">若要了解有关外接程序在一般情况下是在哪里激活和不激活的更多信息，请参阅 Outlook[](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)外接程序概述页的"可用于外接程序的邮箱项目"部分。</span><span class="sxs-lookup"><span data-stu-id="f2904-133">To learn more about where add-ins do and do not activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.</span></span>

## <a name="supported-permissions"></a><span data-ttu-id="f2904-134">支持的权限</span><span class="sxs-lookup"><span data-stu-id="f2904-134">Supported permissions</span></span>

<span data-ttu-id="f2904-135">下表介绍了 JavaScript API 支持Office和共享邮箱用户的权限。</span><span class="sxs-lookup"><span data-stu-id="f2904-135">The following table describes the permissions that the Office JavaScript API supports for delegates and shared mailbox users.</span></span>

|<span data-ttu-id="f2904-136">权限</span><span class="sxs-lookup"><span data-stu-id="f2904-136">Permission</span></span>|<span data-ttu-id="f2904-137">值</span><span class="sxs-lookup"><span data-stu-id="f2904-137">Value</span></span>|<span data-ttu-id="f2904-138">说明</span><span class="sxs-lookup"><span data-stu-id="f2904-138">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="f2904-139">读取</span><span class="sxs-lookup"><span data-stu-id="f2904-139">Read</span></span>|<span data-ttu-id="f2904-140">1 (0000001) </span><span class="sxs-lookup"><span data-stu-id="f2904-140">1 (000001)</span></span>|<span data-ttu-id="f2904-141">可读取项目。</span><span class="sxs-lookup"><span data-stu-id="f2904-141">Can read items.</span></span>|
|<span data-ttu-id="f2904-142">写入</span><span class="sxs-lookup"><span data-stu-id="f2904-142">Write</span></span>|<span data-ttu-id="f2904-143">2 (000010) </span><span class="sxs-lookup"><span data-stu-id="f2904-143">2 (000010)</span></span>|<span data-ttu-id="f2904-144">可以创建项目。</span><span class="sxs-lookup"><span data-stu-id="f2904-144">Can create items.</span></span>|
|<span data-ttu-id="f2904-145">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="f2904-145">DeleteOwn</span></span>|<span data-ttu-id="f2904-146">4 (000100) </span><span class="sxs-lookup"><span data-stu-id="f2904-146">4 (000100)</span></span>|<span data-ttu-id="f2904-147">只能删除他们创建的项。</span><span class="sxs-lookup"><span data-stu-id="f2904-147">Can delete only the items they created.</span></span>|
|<span data-ttu-id="f2904-148">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="f2904-148">DeleteAll</span></span>|<span data-ttu-id="f2904-149">8 (001000) </span><span class="sxs-lookup"><span data-stu-id="f2904-149">8 (001000)</span></span>|<span data-ttu-id="f2904-150">可以删除任何项目。</span><span class="sxs-lookup"><span data-stu-id="f2904-150">Can delete any items.</span></span>|
|<span data-ttu-id="f2904-151">EditOwn</span><span class="sxs-lookup"><span data-stu-id="f2904-151">EditOwn</span></span>|<span data-ttu-id="f2904-152">16 (010000) </span><span class="sxs-lookup"><span data-stu-id="f2904-152">16 (010000)</span></span>|<span data-ttu-id="f2904-153">只能编辑他们创建的项。</span><span class="sxs-lookup"><span data-stu-id="f2904-153">Can edit only the items they created.</span></span>|
|<span data-ttu-id="f2904-154">EditAll</span><span class="sxs-lookup"><span data-stu-id="f2904-154">EditAll</span></span>|<span data-ttu-id="f2904-155">32 (1000000) </span><span class="sxs-lookup"><span data-stu-id="f2904-155">32 (100000)</span></span>|<span data-ttu-id="f2904-156">可以编辑任何项目。</span><span class="sxs-lookup"><span data-stu-id="f2904-156">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="f2904-157">目前，API 支持获取现有权限，但不支持设置权限。</span><span class="sxs-lookup"><span data-stu-id="f2904-157">Currently the API supports getting existing permissions, but not setting permissions.</span></span>

<span data-ttu-id="f2904-158">使用位掩码来指示权限实现 [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) 对象。</span><span class="sxs-lookup"><span data-stu-id="f2904-158">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the permissions.</span></span> <span data-ttu-id="f2904-159">位掩码中的每个位置表示特定权限，如果设置为 ， `1` 则用户具有各自的权限。</span><span class="sxs-lookup"><span data-stu-id="f2904-159">Each position in the bitmask represents a particular permission and if it's set to `1` then the user has the respective permission.</span></span> <span data-ttu-id="f2904-160">例如，如果右边的第二位是 ， `1` 则用户具有 **写入** 权限。</span><span class="sxs-lookup"><span data-stu-id="f2904-160">For example, if the second bit from the right is `1`, then the user has **Write** permission.</span></span> <span data-ttu-id="f2904-161">您可以在本文稍后的以委派或共享邮箱用户角色执行操作部分查看[](#perform-an-operation-as-delegate-or-shared-mailbox-user)如何检查特定权限的示例。</span><span class="sxs-lookup"><span data-stu-id="f2904-161">You can see an example of how to check for a specific permission in the [Perform an operation as delegate or shared mailbox user](#perform-an-operation-as-delegate-or-shared-mailbox-user) section later in this article.</span></span>

## <a name="sync-across-shared-folder-clients"></a><span data-ttu-id="f2904-162">跨共享文件夹客户端同步</span><span class="sxs-lookup"><span data-stu-id="f2904-162">Sync across shared folder clients</span></span>

<span data-ttu-id="f2904-163">代理对所有者邮箱的更新通常会立即跨邮箱同步。</span><span class="sxs-lookup"><span data-stu-id="f2904-163">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="f2904-164">但是，如果使用 REST 或 Exchange Web (EWS) 操作来设置项目的扩展属性，则此类更改可能需要几个小时才能同步。我们建议你改为使用[CustomProperties](/javascript/api/outlook/office.customproperties)对象和相关 API 以避免此类延迟。</span><span class="sxs-lookup"><span data-stu-id="f2904-164">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="f2904-165">若要了解更多信息，请参阅"[](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)在加载项中获取和设置Outlook元数据"一文的自定义属性部分。</span><span class="sxs-lookup"><span data-stu-id="f2904-165">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f2904-166">在委派方案中，不能将 EWS 与当前由 office.js API 提供的令牌一起使用。</span><span class="sxs-lookup"><span data-stu-id="f2904-166">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f2904-167">配置清单</span><span class="sxs-lookup"><span data-stu-id="f2904-167">Configure the manifest</span></span>

<span data-ttu-id="f2904-168">若要在加载项中启用共享文件夹和共享邮箱方案，必须在父元素 下的清单中将 [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` 元素设置为 `DesktopFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="f2904-168">To enable shared folders and shared mailbox scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="f2904-169">目前，不支持其他外形因素。</span><span class="sxs-lookup"><span data-stu-id="f2904-169">At present, other form factors are not supported.</span></span>

<span data-ttu-id="f2904-170">若要支持从代理进行 REST 调用，将清单 [中的"权限"](../reference/manifest/permissions.md) 节点设置为 `ReadWriteMailbox` 。</span><span class="sxs-lookup"><span data-stu-id="f2904-170">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="f2904-171">以下示例显示清单 `SupportsSharedFolders` 的一节中设置为 `true` 的 元素。</span><span class="sxs-lookup"><span data-stu-id="f2904-171">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a><span data-ttu-id="f2904-172">以委派邮箱用户或共享邮箱用户模式执行操作</span><span class="sxs-lookup"><span data-stu-id="f2904-172">Perform an operation as delegate or shared mailbox user</span></span>

<span data-ttu-id="f2904-173">可以通过调用 [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法在撰写或阅读模式下获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="f2904-173">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="f2904-174">这将返回 [一个 SharedProperties](/javascript/api/outlook/office.sharedproperties) 对象，该对象当前提供用户的权限、所有者的电子邮件地址、REST API 的基本 URL 和目标邮箱。</span><span class="sxs-lookup"><span data-stu-id="f2904-174">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the user's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="f2904-175">以下示例演示如何获取邮件或约会的共享属性、检查代理或共享邮箱用户是否具有写入权限以及进行 REST调用。</span><span class="sxs-lookup"><span data-stu-id="f2904-175">The following example shows how to get the shared properties of a message or appointment, check if the delegate or shared mailbox user has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="f2904-176">作为代理，您可以使用 REST 获取附加到项目或组帖子Outlook邮件Outlook[内容](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。</span><span class="sxs-lookup"><span data-stu-id="f2904-176">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="f2904-177">处理对共享项和非共享项的调用 REST</span><span class="sxs-lookup"><span data-stu-id="f2904-177">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="f2904-178">如果要对项目调用 REST 操作（无论该项是否共享）都可以使用 API 来确定 `getSharedPropertiesAsync` 该项目是否共享。</span><span class="sxs-lookup"><span data-stu-id="f2904-178">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="f2904-179">然后，您可以使用适当的对象构造该操作的 REST URL。</span><span class="sxs-lookup"><span data-stu-id="f2904-179">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="f2904-180">限制</span><span class="sxs-lookup"><span data-stu-id="f2904-180">Limitations</span></span>

<span data-ttu-id="f2904-181">根据外接程序的方案，在处理共享文件夹或共享邮箱情况时需要考虑一些限制。</span><span class="sxs-lookup"><span data-stu-id="f2904-181">Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="f2904-182">邮件撰写模式</span><span class="sxs-lookup"><span data-stu-id="f2904-182">Message Compose mode</span></span>

<span data-ttu-id="f2904-183">在邮件撰写模式下[，getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_)在 Outlook 网页版 或 Windows都不受支持，除非满足以下条件。</span><span class="sxs-lookup"><span data-stu-id="f2904-183">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or on Windows unless the following conditions are met.</span></span>

<span data-ttu-id="f2904-184">a.</span><span class="sxs-lookup"><span data-stu-id="f2904-184">a.</span></span> <span data-ttu-id="f2904-185">**委派访问权限/共享文件夹**</span><span class="sxs-lookup"><span data-stu-id="f2904-185">**Delegate access/Shared folders**</span></span>

1. <span data-ttu-id="f2904-186">邮箱所有者启动一封邮件。</span><span class="sxs-lookup"><span data-stu-id="f2904-186">The mailbox owner starts a message.</span></span> <span data-ttu-id="f2904-187">这可以是新邮件、回复或转发。</span><span class="sxs-lookup"><span data-stu-id="f2904-187">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="f2904-188">他们保存邮件，然后将邮件从自己的 **"草稿** "文件夹移动到与代理共享的文件夹。</span><span class="sxs-lookup"><span data-stu-id="f2904-188">They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.</span></span>
1. <span data-ttu-id="f2904-189">代理从共享文件夹打开草稿，然后继续撰写。</span><span class="sxs-lookup"><span data-stu-id="f2904-189">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="f2904-190">b.</span><span class="sxs-lookup"><span data-stu-id="f2904-190">b.</span></span> <span data-ttu-id="f2904-191">**共享邮箱**</span><span class="sxs-lookup"><span data-stu-id="f2904-191">**Shared mailbox**</span></span>

1. <span data-ttu-id="f2904-192">共享邮箱用户启动邮件。</span><span class="sxs-lookup"><span data-stu-id="f2904-192">A shared mailbox user starts a message.</span></span> <span data-ttu-id="f2904-193">这可以是新邮件、回复或转发。</span><span class="sxs-lookup"><span data-stu-id="f2904-193">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="f2904-194">他们保存邮件，然后将邮件从自己的 **"草稿** "文件夹移动到共享邮箱中的文件夹。</span><span class="sxs-lookup"><span data-stu-id="f2904-194">They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.</span></span>
1. <span data-ttu-id="f2904-195">另一个共享邮箱用户从共享邮箱打开草稿，然后继续撰写。</span><span class="sxs-lookup"><span data-stu-id="f2904-195">Another shared mailbox user opens the draft from the shared mailbox then continues composing.</span></span>

<span data-ttu-id="f2904-196">消息现在位于共享上下文中，支持这些共享方案的外接程序可以获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="f2904-196">The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.</span></span> <span data-ttu-id="f2904-197">邮件发送后，通常会在发件人的"已发送邮件" **文件夹中找到** 该邮件。</span><span class="sxs-lookup"><span data-stu-id="f2904-197">After the message has been sent, it's usually found in the sender's **Sent Items** folder.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="f2904-198">REST 和 EWS</span><span class="sxs-lookup"><span data-stu-id="f2904-198">REST and EWS</span></span>

<span data-ttu-id="f2904-199">您的外接程序可以使用 REST，并且外接程序的权限必须设置为，才能启用对所有者邮箱或共享邮箱的 `ReadWriteMailbox` REST 访问（如果适用）。</span><span class="sxs-lookup"><span data-stu-id="f2904-199">Your add-in can use REST and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox or to the shared mailbox as applicable.</span></span> <span data-ttu-id="f2904-200">不支持 EWS。</span><span class="sxs-lookup"><span data-stu-id="f2904-200">EWS is not supported.</span></span>

## <a name="see-also"></a><span data-ttu-id="f2904-201">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f2904-201">See also</span></span>

- [<span data-ttu-id="f2904-202">允许其他人管理邮件和日历</span><span class="sxs-lookup"><span data-stu-id="f2904-202">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="f2904-203">日历中的日历Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="f2904-203">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="f2904-204">将共享邮箱添加到Outlook</span><span class="sxs-lookup"><span data-stu-id="f2904-204">Add a shared mailbox to Outlook</span></span>](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [<span data-ttu-id="f2904-205">如何对清单元素排序</span><span class="sxs-lookup"><span data-stu-id="f2904-205">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="f2904-206">[计算 (的) ](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="f2904-206">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="f2904-207">JavaScript 位运算符</span><span class="sxs-lookup"><span data-stu-id="f2904-207">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)