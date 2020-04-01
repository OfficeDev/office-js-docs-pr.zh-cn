---
title: Outlook 加载项的 Onsend 功能
description: 提供了一种处理项目或阻止用户进行特定操作的方法，并允许加载项在发送时设置某些属性。
ms.date: 03/30/2020
localization_priority: Normal
ms.openlocfilehash: 59d633169fa74687032691bef65fb7f0b114822a
ms.sourcegitcommit: 73a3df90a51acf13416d6a049bddcd9aabc32441
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2020
ms.locfileid: "43069307"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="ba4c1-103">Outlook 加载项的 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="ba4c1-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="ba4c1-p101">Outlook 加载项的 Onsend 功能提供了一种处理邮件或会议项目，或阻止用户进行特定操作的方法，并允许加载项在发送时设置某些属性。例如，可以使用 Onsend 功能执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="ba4c1-106">防止用户发送敏感信息或将主题行留空。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="ba4c1-107">将特定的收件人添加到邮件中的“抄送”行中，或添加到会议中的“可选收件人”行中。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

> [!NOTE]
> <span data-ttu-id="ba4c1-108">Exchange Online (Office 365)、Exchange 2016 本地版本（累积更新 6 或更高版本）和 Exchange 2019 本地版本（累积更新 1 或更高版本）中 Outlook 网页版支持 on-send 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-108">The on-send feature is currently supported for Outlook on the web in Exchange Online (Office 365), Exchange 2016 on-premises (Cumulative Update 6 or later), and Exchange 2019 on-premises (Cumulative Update 1 or later).</span></span> <span data-ttu-id="ba4c1-109">Windows 和 Mac 上的最新 Outlook 内部版本中也提供了此功能，与 Exchange Online (Office 365) 连接。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-109">This feature is also available in the latest Outlook builds on Windows and Mac, connected to Exchange Online (Office 365).</span></span> <span data-ttu-id="ba4c1-110">在要求集1.8 （[当前服务器和客户端支持](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)）中引入了此功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-110">The feature was introduced in requirement set 1.8 ([current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ba4c1-111">[AppSource](https://appsource.microsoft.com)中不允许使用 "发送时" 功能的外接程序。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-111">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="ba4c1-112">on-send 功能是由事件类型 `ItemSend` 触发的，无 UI。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-112">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="ba4c1-113">有关 Onsend 功能的限制信息，请参阅本文稍后部分中介绍的[限制](#limitations)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-113">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="ba4c1-114">Onsend 功能的工作原理</span><span class="sxs-lookup"><span data-stu-id="ba4c1-114">How does the on-send feature work?</span></span>

<span data-ttu-id="ba4c1-115">可使用 Onsend 功能生成集成了 `ItemSend` 同步事件的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-115">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="ba4c1-116">此事件检测到用户正在按“**发送**”按钮（或现有会议的“**发送更新**”按钮），并且如果验证失败，则可用于阻止该项目发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-116">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="ba4c1-117">例如，当用户触发邮件发送事件时，使用 Onsend 功能的 Outlook 加载项可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-117">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="ba4c1-118">读取和验证电子邮件内容</span><span class="sxs-lookup"><span data-stu-id="ba4c1-118">Read and validate the email message contents</span></span>
- <span data-ttu-id="ba4c1-119">验证邮件是否包含主题行</span><span class="sxs-lookup"><span data-stu-id="ba4c1-119">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="ba4c1-120">设置预先确定的收件人</span><span class="sxs-lookup"><span data-stu-id="ba4c1-120">Set a predetermined recipient</span></span>

<span data-ttu-id="ba4c1-121">当触发 send 事件时，将在 Outlook 中对客户端进行验证，并且外接程序在超时之前最长可达5分钟。如果验证失败，将阻止发送项目，并在信息栏中显示一条错误消息，提示用户执行操作。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-121">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

<span data-ttu-id="ba4c1-122">以下屏幕截图显示了通知发件人添加主题的信息栏。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-122">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![屏幕截图显示一个错误消息，提示用户输入缺失的主题行](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="ba4c1-124">以下屏幕截图显示了一个信息栏，通知发件人已找到禁止使用的词语。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-124">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![屏幕截图显示一条错误消息，告诉用户已找到禁止使用的词语](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="ba4c1-126">限制</span><span class="sxs-lookup"><span data-stu-id="ba4c1-126">Limitations</span></span>

<span data-ttu-id="ba4c1-127">Onsend 功能目前具有以下限制。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-127">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="ba4c1-128">**AppSource** &ndash; 无法在 [AppSource](https://appsource.microsoft.com) 中发布使用 Onsend 功能的 Outlook 加载项，因为它们将无法通过 AppSource 验证。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-128">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="ba4c1-129">使用 Onsend 功能的加载项应由管理员部署。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-129">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="ba4c1-130">**清单**&ndash; - 每个加载项仅支持一个 `ItemSend` 事件。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-130">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="ba4c1-131">如果清单中有两个或多个 `ItemSend` 事件，则该清单将无法通过验证。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-131">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="ba4c1-p106">**性能** &ndash; 多次往返到托管加载项的 Web 服务器可能会影响加载项的性能。创建需要多个基于邮件或会议操作的加载项时，请考虑性能影响。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-p106">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="ba4c1-134">**稍后发送**（仅适用于 Mac）&ndash; 如果有 Onsend 加载项，**稍后发送**功能将不可用。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-134">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="ba4c1-135">邮箱类型/模式限制</span><span class="sxs-lookup"><span data-stu-id="ba4c1-135">Mailbox type/mode limitations</span></span>

<span data-ttu-id="ba4c1-136">只有 Outlook 网页版、Windows 版和 Mac 版中的用户邮箱支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-136">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="ba4c1-137">当前不可对以下邮箱类型和模式使用此功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-137">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="ba4c1-138">共享邮箱\*</span><span class="sxs-lookup"><span data-stu-id="ba4c1-138">Shared mailboxes\*</span></span>
- <span data-ttu-id="ba4c1-139">组邮箱</span><span class="sxs-lookup"><span data-stu-id="ba4c1-139">Group mailboxes</span></span>
- <span data-ttu-id="ba4c1-140">脱机模式</span><span class="sxs-lookup"><span data-stu-id="ba4c1-140">Offline mode</span></span>

<span data-ttu-id="ba4c1-141">如果对这些邮箱场景启用了 Onsend 功能，则 Outlook 将不允许进行发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-141">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="ba4c1-142">但是，如果用户答复组邮箱中的电子邮件，则 Onsend 加载项将不运行且系统将发送邮件。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-142">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ba4c1-143">\*如果外接程序还[实现对代理访问方案的支持](delegate-access.md)，则发送时功能应适用于共享邮箱或文件夹。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-143">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="ba4c1-144">多个 Onsend 加载项</span><span class="sxs-lookup"><span data-stu-id="ba4c1-144">Multiple on-send add-ins</span></span>

<span data-ttu-id="ba4c1-145">如果安装了多个 Onsend 加载项，则加载项将按照从 API `getAppManifestCall` 或 `getExtensibilityContext` 接收到的顺序运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-145">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="ba4c1-146">如果第一个外接程序允许发送，则第二个外接程序可以更改阻止第一个外接程序进行发送的某些设置。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-146">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="ba4c1-147">但是，如果所有已安装的外接程序均允许发送，则第一个外接程序将不会重新运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-147">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="ba4c1-148">例如，Add-in1 和 Add-in2 均使用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-148">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="ba4c1-149">首先安装的是 Add-in1，接着安装的是 Add-in2。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-149">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="ba4c1-150">Add-in1 验证邮件中出现的 Fabrikam 一词作为外接程序允许发送的条件。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-150">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="ba4c1-151">但是，Add-in2 可以删除出现的所有 Fabrikam 词语。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-151">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="ba4c1-152">邮件将与已删除 Fabrikam 的所有实例一同发送（归因于 Add-in1 和 Add-in2 的安装顺序）。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-152">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="ba4c1-153">部署使用 Onsend 的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="ba4c1-153">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="ba4c1-154">建议管理员部署使用 Onsend 功能的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-154">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="ba4c1-155">管理员必须确保 Onsend 加载项满足以下条件：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-155">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="ba4c1-156">任何时候打开撰写项目时均可用（针对电子邮件：新建、回复或转发）。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-156">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="ba4c1-157">用户无法关闭或禁用。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-157">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="ba4c1-158">安装使用 Onsend 的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="ba4c1-158">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="ba4c1-159">Outlook 中的 Onsend 功能要求针对发送事件类型配置加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-159">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="ba4c1-160">选择要配置的平台。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-160">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="ba4c1-161">Web 浏览器 - 经典 Outlook</span><span class="sxs-lookup"><span data-stu-id="ba4c1-161">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="ba4c1-162">对于分配了将 *OnSendAddinsEnabled* 标志设置为 **true** 的 Outlook 网页版邮箱策略的用户，系统会为其运行使用 Onsend 功能的 Outlook 网页版（经典）的加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-162">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="ba4c1-163">若要安装新的外接程序，请运行以下 Exchange Online PowerShell cmdlet。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-163">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="ba4c1-164">若要了解如何使用远程 PowerShell 连接到 Exchange Online，请参阅[连接到 Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-164">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="ba4c1-165">启用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="ba4c1-165">Enable the on-send feature</span></span>

<span data-ttu-id="ba4c1-166">默认情况下，Onsend 功能处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-166">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="ba4c1-167">管理员可以通过运行 Exchange Online PowerShell cmdlet 启用 Onsend。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-167">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="ba4c1-168">要为所有用户启用 Onsend 加载项，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-168">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="ba4c1-169">创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-169">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="ba4c1-170">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-170">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="ba4c1-171">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-171">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ba4c1-172">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-172">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ba4c1-173">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-173">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="ba4c1-174">为一组用户启用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="ba4c1-174">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="ba4c1-175">为特定用户组启用 Onsend 功能的步骤如下。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-175">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="ba4c1-176">在此示例中，管理员仅希望在财务用户（其中财务用户属于财务部门）的环境中启用 Outlook 网页版 Onsend 加载项功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-176">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="ba4c1-177">为该组创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-177">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="ba4c1-178">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能（有关详细信息，请参阅本文前面介绍的[邮箱类型限制](#multiple-on-send-add-ins)）。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-178">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="ba4c1-179">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-179">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ba4c1-180">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-180">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ba4c1-181">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-181">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="ba4c1-182">需要等待 60 分钟该策略才能生效，或重启 Internet Information Services (IIS)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-182">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="ba4c1-183">策略生效后，将为该组启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-183">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="ba4c1-184">禁用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="ba4c1-184">Disable the on-send feature</span></span>

<span data-ttu-id="ba4c1-185">若要禁用用户的 Onsend 功能或分配未启用该标志的 Outlook 网页版邮箱策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-185">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="ba4c1-186">在此示例中，该邮箱策略是 *ContosoCorpOWAPolicy*。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-186">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="ba4c1-187">有关如何使用 **Set-OwaMailboxPolicy** cmdlet 配置现有 Outlook 网页版邮箱策略的详细信息，请参阅 [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-187">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="ba4c1-188">若要禁用所有分配了指定 Outlook 网页版邮箱策略的用户的 Onsend 功能，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-188">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="ba4c1-189">Web 浏览器 - 新式 Outlook</span><span class="sxs-lookup"><span data-stu-id="ba4c1-189">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="ba4c1-190">对于安装了使用 Onsend 功能的 Outlook 网页版（新式）加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-190">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="ba4c1-191">但是，如果用户需要运行该加载项来满足合规性标准，则邮箱策略必须将 *OnSendAddinsEnabled* 标志设置为 **true**。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-191">However, if users are required to run the add-in to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="ba4c1-192">若要安装新的外接程序，请运行以下 Exchange Online PowerShell cmdlet。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-192">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="ba4c1-193">若要了解如何使用远程 PowerShell 连接到 Exchange Online，请参阅[连接到 Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-193">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-policy"></a><span data-ttu-id="ba4c1-194">启用 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="ba4c1-194">Enable the on-send policy</span></span>

<span data-ttu-id="ba4c1-195">默认情况下，Onsend 策略处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-195">By default, on-send policy is disabled.</span></span> <span data-ttu-id="ba4c1-196">管理员可以通过运行 Exchange Online PowerShell cmdlet 启用 Onsend。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-196">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="ba4c1-197">要为所有用户启用 Onsend 加载项，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-197">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="ba4c1-198">创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-198">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="ba4c1-199">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-199">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="ba4c1-200">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-200">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ba4c1-201">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-201">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ba4c1-202">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-202">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-policy-for-a-group-of-users"></a><span data-ttu-id="ba4c1-203">为一组用户启用 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="ba4c1-203">Enable the on-send policy for a group of users</span></span>

<span data-ttu-id="ba4c1-204">为特定用户组启用 Onsend 策略的步骤如下。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-204">To enable the on-send policy for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="ba4c1-205">在此示例中，管理员仅希望在财务用户（其中财务用户属于财务部门）的环境中启用 Outlook 网页版 Onsend 加载项策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-205">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="ba4c1-206">为该组创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-206">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="ba4c1-207">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能（有关详细信息，请参阅本文前面介绍的[邮箱类型限制](#multiple-on-send-add-ins)）。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-207">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="ba4c1-208">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-208">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ba4c1-209">启用 Onsend 策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-209">Enable the on-send policy.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ba4c1-210">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-210">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="ba4c1-211">需要等待 60 分钟该策略才能生效，或重启 Internet Information Services (IIS)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-211">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="ba4c1-212">策略生效后，将为该组强制执行 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-212">When the policy takes effect, the on-send feature will be enforced for the group.</span></span>

#### <a name="disable-the-on-send-policy"></a><span data-ttu-id="ba4c1-213">禁用 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="ba4c1-213">Disable the on-send policy</span></span>

<span data-ttu-id="ba4c1-214">若要禁用用户的 Onsend 策略或分配未启用该标志的 Outlook 网页版邮箱策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-214">To disable the on-send policy for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="ba4c1-215">在此示例中，该邮箱策略是 *ContosoCorpOWAPolicy*。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-215">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="ba4c1-216">有关如何使用 **Set-OwaMailboxPolicy** cmdlet 配置现有 Outlook 网页版邮箱策略的详细信息，请参阅 [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-216">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="ba4c1-217">若要禁用所有分配了指定 Outlook 网页版邮箱策略的用户的 Onsend 策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-217">To disable the on-send policy for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="ba4c1-218">Windows</span><span class="sxs-lookup"><span data-stu-id="ba4c1-218">Windows</span></span>](#tab/windows)

<span data-ttu-id="ba4c1-219">对于安装了使用 Onsend 功能的 Windows 版 Outlook 加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-219">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="ba4c1-220">但是，如果用户需要运行该加载项来满足合规性标准，则必须在每台适用的计算机上将组策略“**无法加载 Web 扩展时禁用发送**”设置为“**已启用**”。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-220">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="ba4c1-221">若要设置邮箱策略，管理员可以下载[管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)，然后通过运行本地组策略编辑器 **(gpedit.msc)** 访问最新的管理模板。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-221">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="ba4c1-222">策略的用途</span><span class="sxs-lookup"><span data-stu-id="ba4c1-222">What the policy does</span></span>

<span data-ttu-id="ba4c1-223">出于合规性原因，管理员可能需要在用户具有可供运行的最新 Onsend 加载项前，确保其无法发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-223">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="ba4c1-224">管理员必须启用组策略“**无法加载 Web 扩展时禁用发送**”，以便所有加载项都从 Exchange 进行更新，并可用于在发送时验证每封邮件或每个会议项目是否符合预期的规则和规定。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-224">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="ba4c1-225">策略状态</span><span class="sxs-lookup"><span data-stu-id="ba4c1-225">Policy status</span></span>|<span data-ttu-id="ba4c1-226">结果</span><span class="sxs-lookup"><span data-stu-id="ba4c1-226">Result</span></span>|
|---|---|
|<span data-ttu-id="ba4c1-227">已禁用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-227">Disabled</span></span>|<span data-ttu-id="ba4c1-228">允许发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-228">Send allowed.</span></span> <span data-ttu-id="ba4c1-229">即使尚未从 Exchange 中更新加载项，也可以在不运行 Onsend 加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-229">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="ba4c1-230">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-230">Enabled</span></span>|<span data-ttu-id="ba4c1-231">仅当加载项已从 Exchange 更新时才允许发送；否则，将阻止发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-231">Send allowed only when the add-in has been updated from Exchange; otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="ba4c1-232">管理 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="ba4c1-232">Manage the on-send policy</span></span>

<span data-ttu-id="ba4c1-233">默认情况下，Onsend 策略处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-233">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="ba4c1-234">管理员可以通过确保用户的组策略设置“**无法加载 Web 扩展时禁用发送**”设置为“**已启用**”来启用 Onsend 策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-234">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="ba4c1-235">若为用户禁用策略，管理员应将其设置为“**已禁用**”。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-235">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="ba4c1-236">若要管理此策略设置，可执行下列操作。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-236">To manage this policy setting, you can do the following.</span></span>

1. <span data-ttu-id="ba4c1-237">下载最新的[管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-237">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="ba4c1-238">打开本地组策略编辑器 (**gpedit.msc**)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-238">Open the Local Group Policy editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="ba4c1-239">导航到 **“用户配置”>“管理模板”>“Microsoft Outlook 2016”>“安全性”>“信任中心”**。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-239">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="ba4c1-240">选择“**无法加载 Web 扩展时禁用发送**”设置。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-240">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="ba4c1-241">打开链接以编辑策略设置。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-241">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="ba4c1-242">在“**无法加载 Web 扩展时禁用发送**”对话框窗口中，根据需要选择“**已启用**”或“**已禁用**”，然后选择“**确定**”或“**应用**”以使更新生效。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-242">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="ba4c1-243">Mac</span><span class="sxs-lookup"><span data-stu-id="ba4c1-243">Mac</span></span>](#tab/unix)

<span data-ttu-id="ba4c1-244">对于安装了使用 Onsend 功能的 Mac 版 Outlook 加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-244">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="ba4c1-245">但是，如果用户需要运行该加载项来满足合规性标准，则必须在每个用户的计算机上应用以下邮箱设置。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-245">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="ba4c1-246">此设置或键与 CFPreferences 兼容，这意味着可以使用适用于 Mac 的企业管理软件（例如，Jamf Pro）来对其进行设置。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-246">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

|||
|:---|:---|
|<span data-ttu-id="ba4c1-247">**域**</span><span class="sxs-lookup"><span data-stu-id="ba4c1-247">**Domain**</span></span>|<span data-ttu-id="ba4c1-248">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="ba4c1-248">com.microsoft.outlook</span></span>|
|<span data-ttu-id="ba4c1-249">**键**</span><span class="sxs-lookup"><span data-stu-id="ba4c1-249">**Key**</span></span>|<span data-ttu-id="ba4c1-250">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="ba4c1-250">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="ba4c1-251">**DataType**</span><span class="sxs-lookup"><span data-stu-id="ba4c1-251">**DataType**</span></span>|<span data-ttu-id="ba4c1-252">Boolean</span><span class="sxs-lookup"><span data-stu-id="ba4c1-252">Boolean</span></span>|
|<span data-ttu-id="ba4c1-253">**可能的值**</span><span class="sxs-lookup"><span data-stu-id="ba4c1-253">**Possible values**</span></span>|<span data-ttu-id="ba4c1-254">false（默认值）</span><span class="sxs-lookup"><span data-stu-id="ba4c1-254">false (default)</span></span><br><span data-ttu-id="ba4c1-255">true</span><span class="sxs-lookup"><span data-stu-id="ba4c1-255">true</span></span>|
|<span data-ttu-id="ba4c1-256">**可用性**</span><span class="sxs-lookup"><span data-stu-id="ba4c1-256">**Availability**</span></span>|<span data-ttu-id="ba4c1-257">16.27</span><span class="sxs-lookup"><span data-stu-id="ba4c1-257">16.27</span></span>|
|<span data-ttu-id="ba4c1-258">**备注**</span><span class="sxs-lookup"><span data-stu-id="ba4c1-258">**Comments**</span></span>|<span data-ttu-id="ba4c1-259">此键将创建 onSendMailbox 策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-259">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="ba4c1-260">设置的用途</span><span class="sxs-lookup"><span data-stu-id="ba4c1-260">What the setting does</span></span>

<span data-ttu-id="ba4c1-261">出于合规性原因，管理员可能需要在用户具有可供运行的最新 Onsend 加载项前，确保其无法发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-261">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="ba4c1-262">管理员必须启用键 **OnSendAddinsWaitForLoad**，以便所有加载项都从 Exchange 进行更新，并可用于在发送时验证每封邮件或每个会议项目是否符合预期的规则和规定。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-262">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="ba4c1-263">键的状态</span><span class="sxs-lookup"><span data-stu-id="ba4c1-263">Key's state</span></span>|<span data-ttu-id="ba4c1-264">结果</span><span class="sxs-lookup"><span data-stu-id="ba4c1-264">Result</span></span>|
|---|---|
|<span data-ttu-id="ba4c1-265">false</span><span class="sxs-lookup"><span data-stu-id="ba4c1-265">false</span></span>|<span data-ttu-id="ba4c1-266">允许发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-266">Send allowed.</span></span> <span data-ttu-id="ba4c1-267">即使尚未从 Exchange 中更新加载项，也可以在不运行 Onsend 加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-267">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="ba4c1-268">true</span><span class="sxs-lookup"><span data-stu-id="ba4c1-268">true</span></span>|<span data-ttu-id="ba4c1-269">仅当加载项已从 Exchange 更新时才允许发送；否则，将阻止发送，并且禁用“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-269">Send allowed only when add-ins have been updated from Exchange; otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="ba4c1-270">Onsend 功能的应用场景</span><span class="sxs-lookup"><span data-stu-id="ba4c1-270">On-send feature scenarios</span></span>

<span data-ttu-id="ba4c1-271">以下是支持和不支持使用 Onsend 功能的加载项的应用场景。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-271">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="ba4c1-272">用户邮箱启用了 Onsend 加载项功能，但未安装任何加载项</span><span class="sxs-lookup"><span data-stu-id="ba4c1-272">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="ba4c1-273">在这种场景中，用户将能够在不执行任何加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-273">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="ba4c1-274">用户邮箱启用了 Onsend 加载项功能，并且安装并启用了支持 Onsend 的加载项</span><span class="sxs-lookup"><span data-stu-id="ba4c1-274">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="ba4c1-275">外接程序在发送事件期间运行，然后允许或阻止用户发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-275">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="ba4c1-276">邮箱委派，其中邮箱 1 具有对邮箱 2 的完全访问权限</span><span class="sxs-lookup"><span data-stu-id="ba4c1-276">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="ba4c1-277">Web 浏览器（经典 Outlook）</span><span class="sxs-lookup"><span data-stu-id="ba4c1-277">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="ba4c1-278">应用场景</span><span class="sxs-lookup"><span data-stu-id="ba4c1-278">Scenario</span></span>|<span data-ttu-id="ba4c1-279">邮箱 1 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="ba4c1-279">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="ba4c1-280">邮箱 2 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="ba4c1-280">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="ba4c1-281">Outlook Web 会话（经典）</span><span class="sxs-lookup"><span data-stu-id="ba4c1-281">Outlook web session (classic)</span></span>|<span data-ttu-id="ba4c1-282">结果</span><span class="sxs-lookup"><span data-stu-id="ba4c1-282">Result</span></span>|<span data-ttu-id="ba4c1-283">是否支持？</span><span class="sxs-lookup"><span data-stu-id="ba4c1-283">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="ba4c1-284">1</span><span class="sxs-lookup"><span data-stu-id="ba4c1-284">1</span></span>|<span data-ttu-id="ba4c1-285">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-285">Enabled</span></span>|<span data-ttu-id="ba4c1-286">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-286">Enabled</span></span>|<span data-ttu-id="ba4c1-287">新会话</span><span class="sxs-lookup"><span data-stu-id="ba4c1-287">New session</span></span>|<span data-ttu-id="ba4c1-288">邮箱 1 无法从邮箱 2 发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-288">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="ba4c1-p133">目前尚不支持。可以使用方案 3 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-p133">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="ba4c1-291">双面</span><span class="sxs-lookup"><span data-stu-id="ba4c1-291">2</span></span>|<span data-ttu-id="ba4c1-292">Disabled</span><span class="sxs-lookup"><span data-stu-id="ba4c1-292">Disabled</span></span>|<span data-ttu-id="ba4c1-293">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-293">Enabled</span></span>|<span data-ttu-id="ba4c1-294">新会话</span><span class="sxs-lookup"><span data-stu-id="ba4c1-294">New session</span></span>|<span data-ttu-id="ba4c1-295">邮箱 1 无法从邮箱 2 发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-295">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="ba4c1-p134">目前尚不支持。可以使用方案 3 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-p134">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="ba4c1-298">第三章</span><span class="sxs-lookup"><span data-stu-id="ba4c1-298">3</span></span>|<span data-ttu-id="ba4c1-299">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-299">Enabled</span></span>|<span data-ttu-id="ba4c1-300">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-300">Enabled</span></span>|<span data-ttu-id="ba4c1-301">同一个会话</span><span class="sxs-lookup"><span data-stu-id="ba4c1-301">Same session</span></span>|<span data-ttu-id="ba4c1-302">分配给邮箱 1 的 Onsend 加载项运行 Onsend。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-302">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="ba4c1-303">支持。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-303">Supported.</span></span>|
|<span data-ttu-id="ba4c1-304">4 </span><span class="sxs-lookup"><span data-stu-id="ba4c1-304">4</span></span>|<span data-ttu-id="ba4c1-305">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-305">Enabled</span></span>|<span data-ttu-id="ba4c1-306">已禁用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-306">Disabled</span></span>|<span data-ttu-id="ba4c1-307">新会话</span><span class="sxs-lookup"><span data-stu-id="ba4c1-307">New session</span></span>|<span data-ttu-id="ba4c1-308">未运行 Onsend 加载项；邮件或会议项目已发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-308">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="ba4c1-309">支持。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-309">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="ba4c1-310">Web 浏览器（新式 Outlook）、Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="ba4c1-310">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="ba4c1-311">若要强制执行 Onsend，管理员应确保对两个邮箱都启用了该策略。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-311">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="ba4c1-312">若要了解如何在加载项中支持委派访问，请参阅[在 Outlook 加载项中启用委派访问方案](delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-312">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="ba4c1-313">组 1 是新式组邮箱，用户邮箱 1 是组 1 的成员</span><span class="sxs-lookup"><span data-stu-id="ba4c1-313">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="ba4c1-314">方案</span><span class="sxs-lookup"><span data-stu-id="ba4c1-314">Scenario</span></span>|<span data-ttu-id="ba4c1-315">邮箱 1 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="ba4c1-315">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="ba4c1-316">是否启用了 Onsend 加载项？</span><span class="sxs-lookup"><span data-stu-id="ba4c1-316">On-send add-ins enabled?</span></span>|<span data-ttu-id="ba4c1-317">邮箱 1 操作</span><span class="sxs-lookup"><span data-stu-id="ba4c1-317">Mailbox 1 action</span></span>|<span data-ttu-id="ba4c1-318">结果</span><span class="sxs-lookup"><span data-stu-id="ba4c1-318">Result</span></span>|<span data-ttu-id="ba4c1-319">是否支持？</span><span class="sxs-lookup"><span data-stu-id="ba4c1-319">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="ba4c1-320">1</span><span class="sxs-lookup"><span data-stu-id="ba4c1-320">1</span></span>|<span data-ttu-id="ba4c1-321">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-321">Enabled</span></span>|<span data-ttu-id="ba4c1-322">是</span><span class="sxs-lookup"><span data-stu-id="ba4c1-322">Yes</span></span>|<span data-ttu-id="ba4c1-323">邮箱 1 撰写发送到组 1 的新邮件或会议。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-323">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="ba4c1-324">发送期间，Onsend 加载项运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-324">On-send add-ins run during send.</span></span>|<span data-ttu-id="ba4c1-325">是</span><span class="sxs-lookup"><span data-stu-id="ba4c1-325">Yes</span></span>|
|<span data-ttu-id="ba4c1-326">双面</span><span class="sxs-lookup"><span data-stu-id="ba4c1-326">2</span></span>|<span data-ttu-id="ba4c1-327">已启用</span><span class="sxs-lookup"><span data-stu-id="ba4c1-327">Enabled</span></span>|<span data-ttu-id="ba4c1-328">是</span><span class="sxs-lookup"><span data-stu-id="ba4c1-328">Yes</span></span>|<span data-ttu-id="ba4c1-329">邮箱 1 在 Outlook 网页版组 1 的组窗口中撰写发送到组 1 的新邮件或会议。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-329">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="ba4c1-330">Onsend 加载项不会在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-330">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="ba4c1-331">目前尚不支持。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-331">Not currently supported.</span></span> <span data-ttu-id="ba4c1-332">可以使用方案 1 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-332">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="ba4c1-333">用户邮箱启用了 Onsend 加载项功能/策略，并且安装并启用了支持 Onsend 的加载项，启用了脱机模式</span><span class="sxs-lookup"><span data-stu-id="ba4c1-333">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="ba4c1-334">Onsend 加载项将根据用户、加载项后端和 Exchange 的联机状态运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-334">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="ba4c1-335">用户的状态</span><span class="sxs-lookup"><span data-stu-id="ba4c1-335">User's state</span></span>

<span data-ttu-id="ba4c1-336">如果用户处于联机状态，则 Onsend 加载项将在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-336">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="ba4c1-337">如果用户处于脱机状态，Onsend 加载项不会在发送期间运行，也不会发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-337">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="ba4c1-338">加载项后端的状态</span><span class="sxs-lookup"><span data-stu-id="ba4c1-338">Add-in backend's state</span></span>

<span data-ttu-id="ba4c1-339">如果 Onsend 加载项的后端处于联机状态且可访问，则将运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-339">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="ba4c1-340">如果后端处于脱机状态，则将禁用发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-340">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="ba4c1-341">Exchange 的状态</span><span class="sxs-lookup"><span data-stu-id="ba4c1-341">Exchange's state</span></span>

<span data-ttu-id="ba4c1-342">如果 Exchange 服务器处于联机状态且可访问，则 Onsend 加载项将在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-342">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="ba4c1-343">如果 Onsend 加载项无法访问 Exchange 并且已启用适用的策略或 cmdlet，则将禁用发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-343">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="ba4c1-344">在处于任何脱机状态的 Mac 上，“**发送**”按钮（或现有会议的“**发送更新**”按钮）将被禁用，并显示当用户脱机时其组织不允许发送的通知。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-344">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>


## <a name="code-examples"></a><span data-ttu-id="ba4c1-345">代码示例</span><span class="sxs-lookup"><span data-stu-id="ba4c1-345">Code examples</span></span>

<span data-ttu-id="ba4c1-346">以下代码示例说明如何创建一个简单的 Onsend 加载项。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-346">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="ba4c1-347">若要下载这些示例所基于的代码示例，请参阅 [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-347">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="ba4c1-348">清单、版本重写和事件</span><span class="sxs-lookup"><span data-stu-id="ba4c1-348">Manifest, version override, and event</span></span>

<span data-ttu-id="ba4c1-349">[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) 代码示例包括两个清单：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-349">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="ba4c1-350">`Contoso Message Body Checker.xml` &ndash; 展示了如何在发送时检查邮件正文是否包含限制字词或敏感信息。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-350">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="ba4c1-351">`Contoso Subject and CC Checker.xml` &ndash; 展示了如何将收件人添加到抄送行，并在发送时验证邮件是否包含主题行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-351">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="ba4c1-352">在 `Contoso Message Body Checker.xml` 清单文件中，将包含在 `ItemSend` 事件中应调用的函数文件和函数名称。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-352">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="ba4c1-353">该操作将同步运行。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-353">The operation runs synchronously.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> <span data-ttu-id="ba4c1-354">如果使用 Visual Studio 2019 开发你的发送外接程序，则可能会收到类似于以下的验证警告： "这是一个无效的 xsi： type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events"。 "若要解决此问题，您需要在[有关此警告的博客](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/)中提供了 MailAppVersionOverridesV1_1 的较新版本的 .Xsd 作为 GitHub gist 提供。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-354">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="ba4c1-355">对于 `Contoso Subject and CC Checker.xml` 清单文件，以下示例中显示了邮件发送事件中要调用的函数文件和函数名称。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-355">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

<br/>

<span data-ttu-id="ba4c1-356">Onsend API 需要 `VersionOverrides v1_1`。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-356">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="ba4c1-357">以下显示如何在清单中添加 `VersionOverrides` 节点。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-357">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On Send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="ba4c1-358">有关详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-358">For more information, see the following:</span></span>
> - [<span data-ttu-id="ba4c1-359">Outlook 外接程序清单</span><span class="sxs-lookup"><span data-stu-id="ba4c1-359">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="ba4c1-360">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="ba4c1-360">VersionOverrides</span></span>](../develop/create-addin-commands.md#step-3-add-versionoverrides-element)
> - [<span data-ttu-id="ba4c1-361">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="ba4c1-361">Office Add-ins XML manifest</span></span>](../overview/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="ba4c1-362">`Event` 和 `item` 对象以及 `body.getAsync` 和 `body.setAsync` 方法</span><span class="sxs-lookup"><span data-stu-id="ba4c1-362">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="ba4c1-363">若要访问当前选择的邮件或会议项目（在本示例中为新撰写的邮件），请使用 `Office.context.mailbox.item` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-363">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="ba4c1-364">`ItemSend` 事件由 Onsend 功能自动传递到清单中指定的函数&mdash;在本示例中为 `validateBody` 函数。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-364">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

```js
var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

<span data-ttu-id="ba4c1-365">`validateBody` 函数以指定格式 (HTML) 获取当前正文，并在回调方法中传递代码想要访问的 `ItemSend` 事件对象。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-365">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="ba4c1-366">除 `getAsync` 方法之外，`Body` 对象还提供了 `setAsync` 方法，可用于将正文替换为指定的文本。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-366">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="ba4c1-367">有关详细信息，请参阅 [Event 对象](/javascript/api/office/office.addincommands.event)和 [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-367">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="ba4c1-368">`NotificationMessages` 对象和 `event.completed` 方法</span><span class="sxs-lookup"><span data-stu-id="ba4c1-368">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="ba4c1-369">`checkBodyOnlyOnSendCallBack` 函数使用正则表达式来确定邮件正文是否包含禁止使用的词语。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-369">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="ba4c1-370">如果该函数发现受限词语数组的匹配项，则将阻止发送电子邮件，并通过信息栏通知发件人。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-370">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="ba4c1-371">为了做到这一点，它使用 `Item` 对象的 `notificationMessages` 属性来返回 `NotificationMessages` 对象。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-371">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="ba4c1-372">然后，通过调用 `addAsync` 方法向该项目添加通知，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-372">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

<span data-ttu-id="ba4c1-373">以下是 `addAsync` 方法的参数：</span><span class="sxs-lookup"><span data-stu-id="ba4c1-373">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="ba4c1-374">`NoSend` &ndash; 一个字符串，即开发人员指定用于引用通知邮件的密钥。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-374">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="ba4c1-375">可用于在以后修改此邮件。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-375">You can use it to modify this message later.</span></span> <span data-ttu-id="ba4c1-376">密钥长度不能超过32个字符。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-376">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="ba4c1-377">`type` &ndash; JSON 对象参数的一个属性。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-377">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="ba4c1-378">表示邮件的类型；类型对应于 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) 枚举的值。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-378">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="ba4c1-379">可能的值是进度指示器、信息消息或错误消息。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-379">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="ba4c1-380">在此示例中，`type` 是错误消息。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-380">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="ba4c1-381">`message` &ndash; JSON 对象参数的一个属性。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-381">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="ba4c1-382">在此示例中，`message` 是通知邮件的文本。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-382">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="ba4c1-383">为表明加载项对由发送操作触发的 `ItemSend` 事件的处理已完成，请调用 `event.completed({allowEvent:Boolean})` 方法。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-383">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="ba4c1-384">`allowEvent` 属性是一个布尔值。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-384">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="ba4c1-385">如果设置为 `true`，则允许发送。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-385">If set to `true`, send is allowed.</span></span> <span data-ttu-id="ba4c1-386">如果设置为 `false`，则将阻止发送电子邮件。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-386">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="ba4c1-387">有关详细信息，请参阅 [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [completed](/javascript/api/office/office.addincommands.event)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-387">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="ba4c1-388">`replaceAsync`、`removeAsync` 和 `getAllAsync` 方法</span><span class="sxs-lookup"><span data-stu-id="ba4c1-388">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="ba4c1-389">除了 `addAsync` 方法之外，`NotificationMessages` 对象还包括 `replaceAsync`、`removeAsync` 和 `getAllAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-389">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="ba4c1-390">此代码示例中不使用这些方法。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-390">These methods are not used in this code sample.</span></span>  <span data-ttu-id="ba4c1-391">有关详细信息，请参阅 [NotificationMessages](/javascript/api/outlook/office.NotificationMessages)。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-391">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="ba4c1-392">主题和抄送检查器代码</span><span class="sxs-lookup"><span data-stu-id="ba4c1-392">Subject and CC checker code</span></span>

<span data-ttu-id="ba4c1-393">以下代码示例介绍如何将收件人添加到抄送行，并验证邮件在发送时是否包含主题。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-393">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="ba4c1-394">此示例使用 Onsend 功能允许或禁止发送电子邮件。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-394">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

<span data-ttu-id="ba4c1-p152">若要详细了解如何将收件人添加到抄送行、验证电子邮件在发送时是否包主题行，以及查看可以使用的 API，请参阅 [Outlook-Add-in-On-Send 示例](https://github.com/OfficeDev/Outlook-Add-in-On-Send)。已充分注释代码。</span><span class="sxs-lookup"><span data-stu-id="ba4c1-p152">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="ba4c1-397">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ba4c1-397">See also</span></span>

- [<span data-ttu-id="ba4c1-398">Outlook 加载项体系结构和功能概述</span><span class="sxs-lookup"><span data-stu-id="ba4c1-398">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="ba4c1-399">加载项命令演示 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="ba4c1-399">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)
