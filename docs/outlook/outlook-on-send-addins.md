---
title: Outlook 加载项的 Onsend 功能
description: 提供了一种处理项目或阻止用户进行特定操作的方法，并允许加载项在发送时设置某些属性。
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: e21082736bea5ac53caecc9222de317906cd220d
ms.sourcegitcommit: e9f23a2857b90a7c17e3152292b548a13a90aa33
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/19/2020
ms.locfileid: "46803770"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="5036c-103">Outlook 加载项的 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="5036c-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="5036c-p101">Outlook 加载项的 Onsend 功能提供了一种处理邮件或会议项目，或阻止用户进行特定操作的方法，并允许加载项在发送时设置某些属性。例如，可以使用 Onsend 功能执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5036c-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="5036c-106">防止用户发送敏感信息或将主题行留空。</span><span class="sxs-lookup"><span data-stu-id="5036c-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="5036c-107">将特定的收件人添加到邮件中的“抄送”行中，或添加到会议中的“可选收件人”行中。</span><span class="sxs-lookup"><span data-stu-id="5036c-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="5036c-108">on-send 功能是由事件类型 `ItemSend` 触发的，无 UI。</span><span class="sxs-lookup"><span data-stu-id="5036c-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="5036c-109">有关 Onsend 功能的限制信息，请参阅本文稍后部分中介绍的[限制](#limitations)。</span><span class="sxs-lookup"><span data-stu-id="5036c-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="5036c-110">支持的客户端和平台</span><span class="sxs-lookup"><span data-stu-id="5036c-110">Supported clients and platforms</span></span>

<span data-ttu-id="5036c-111">下表显示了用于 "发送" 功能的受支持的客户端/服务器组合。</span><span class="sxs-lookup"><span data-stu-id="5036c-111">The following table shows supported client-server combinations for the on-send feature.</span></span> <span data-ttu-id="5036c-112">不支持排除的组合。</span><span class="sxs-lookup"><span data-stu-id="5036c-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="5036c-113">客户端</span><span class="sxs-lookup"><span data-stu-id="5036c-113">Client</span></span> | <span data-ttu-id="5036c-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="5036c-114">Exchange Online</span></span> | <span data-ttu-id="5036c-115">Exchange 2016 本地</span><span class="sxs-lookup"><span data-stu-id="5036c-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="5036c-116"> (累积更新6或更高版本) </span><span class="sxs-lookup"><span data-stu-id="5036c-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="5036c-117">Exchange 2019 本地</span><span class="sxs-lookup"><span data-stu-id="5036c-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="5036c-118"> (累积更新1或更高版本) </span><span class="sxs-lookup"><span data-stu-id="5036c-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="5036c-119">Windows：</span><span class="sxs-lookup"><span data-stu-id="5036c-119">Windows:</span></span><br><span data-ttu-id="5036c-120">版本 1910 (内部版本 12130.20272) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="5036c-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="5036c-121">是</span><span class="sxs-lookup"><span data-stu-id="5036c-121">Yes</span></span>|<span data-ttu-id="5036c-122">是</span><span class="sxs-lookup"><span data-stu-id="5036c-122">Yes</span></span>|<span data-ttu-id="5036c-123">是</span><span class="sxs-lookup"><span data-stu-id="5036c-123">Yes</span></span>|
|<span data-ttu-id="5036c-124">Mac</span><span class="sxs-lookup"><span data-stu-id="5036c-124">Mac:</span></span><br><span data-ttu-id="5036c-125">生成16.30 或更高版本</span><span class="sxs-lookup"><span data-stu-id="5036c-125">build 16.30 or later</span></span>|<span data-ttu-id="5036c-126">是</span><span class="sxs-lookup"><span data-stu-id="5036c-126">Yes</span></span>|<span data-ttu-id="5036c-127">否</span><span class="sxs-lookup"><span data-stu-id="5036c-127">No</span></span>|<span data-ttu-id="5036c-128">否</span><span class="sxs-lookup"><span data-stu-id="5036c-128">No</span></span>|
|<span data-ttu-id="5036c-129">Web 浏览器：</span><span class="sxs-lookup"><span data-stu-id="5036c-129">Web browser:</span></span><br><span data-ttu-id="5036c-130">新式 Outlook UI</span><span class="sxs-lookup"><span data-stu-id="5036c-130">modern Outlook UI</span></span>|<span data-ttu-id="5036c-131">是</span><span class="sxs-lookup"><span data-stu-id="5036c-131">Yes</span></span>|<span data-ttu-id="5036c-132">不适用</span><span class="sxs-lookup"><span data-stu-id="5036c-132">Not applicable</span></span>|<span data-ttu-id="5036c-133">不适用</span><span class="sxs-lookup"><span data-stu-id="5036c-133">Not applicable</span></span>|
|<span data-ttu-id="5036c-134">Web 浏览器：</span><span class="sxs-lookup"><span data-stu-id="5036c-134">Web browser:</span></span><br><span data-ttu-id="5036c-135">经典 Outlook UI</span><span class="sxs-lookup"><span data-stu-id="5036c-135">classic Outlook UI</span></span>|<span data-ttu-id="5036c-136">不适用</span><span class="sxs-lookup"><span data-stu-id="5036c-136">Not applicable</span></span>|<span data-ttu-id="5036c-137">是</span><span class="sxs-lookup"><span data-stu-id="5036c-137">Yes</span></span>|<span data-ttu-id="5036c-138">是</span><span class="sxs-lookup"><span data-stu-id="5036c-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="5036c-139">按发送功能已发布在要求集1.8 中 (有关详细信息) ，请参阅 [当前服务器和客户端支持](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) 。</span><span class="sxs-lookup"><span data-stu-id="5036c-139">The on-send feature was released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5036c-140">[AppSource](https://appsource.microsoft.com)中不允许使用 "发送时" 功能的外接程序。</span><span class="sxs-lookup"><span data-stu-id="5036c-140">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="5036c-141">Onsend 功能的工作原理</span><span class="sxs-lookup"><span data-stu-id="5036c-141">How does the on-send feature work?</span></span>

<span data-ttu-id="5036c-142">可使用 Onsend 功能生成集成了 `ItemSend` 同步事件的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-142">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="5036c-143">此事件检测到用户正在按“**发送**”按钮（或现有会议的“**发送更新**”按钮），并且如果验证失败，则可用于阻止该项目发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-143">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="5036c-144">例如，当用户触发邮件发送事件时，使用 Onsend 功能的 Outlook 加载项可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5036c-144">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="5036c-145">读取和验证电子邮件内容</span><span class="sxs-lookup"><span data-stu-id="5036c-145">Read and validate the email message contents</span></span>
- <span data-ttu-id="5036c-146">验证邮件是否包含主题行</span><span class="sxs-lookup"><span data-stu-id="5036c-146">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="5036c-147">设置预先确定的收件人</span><span class="sxs-lookup"><span data-stu-id="5036c-147">Set a predetermined recipient</span></span>

<span data-ttu-id="5036c-148">当触发 send 事件时，将在 Outlook 中对客户端进行验证，并且外接程序在超时之前最长可达5分钟。如果验证失败，将阻止发送项目，并在信息栏中显示一条错误消息，提示用户执行操作。</span><span class="sxs-lookup"><span data-stu-id="5036c-148">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

<span data-ttu-id="5036c-149">以下屏幕截图显示了通知发件人添加主题的信息栏。</span><span class="sxs-lookup"><span data-stu-id="5036c-149">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![屏幕截图显示一个错误消息，提示用户输入缺失的主题行](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="5036c-151">以下屏幕截图显示了一个信息栏，通知发件人已找到禁止使用的词语。</span><span class="sxs-lookup"><span data-stu-id="5036c-151">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![屏幕截图显示一条错误消息，告诉用户已找到禁止使用的词语](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="5036c-153">限制</span><span class="sxs-lookup"><span data-stu-id="5036c-153">Limitations</span></span>

<span data-ttu-id="5036c-154">Onsend 功能目前具有以下限制。</span><span class="sxs-lookup"><span data-stu-id="5036c-154">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="5036c-155">如果您调用正文， (preview) **的追加-发送**功能 &ndash; [。AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)在发送处理程序中，返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="5036c-155">**Append-on-send** feature (preview) &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="5036c-156">**AppSource** &ndash; 无法在 [AppSource](https://appsource.microsoft.com) 中发布使用 Onsend 功能的 Outlook 加载项，因为它们将无法通过 AppSource 验证。</span><span class="sxs-lookup"><span data-stu-id="5036c-156">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="5036c-157">使用 Onsend 功能的加载项应由管理员部署。</span><span class="sxs-lookup"><span data-stu-id="5036c-157">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="5036c-158">**清单**&ndash; - 每个加载项仅支持一个 `ItemSend` 事件。</span><span class="sxs-lookup"><span data-stu-id="5036c-158">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="5036c-159">如果清单中有两个或多个 `ItemSend` 事件，则该清单将无法通过验证。</span><span class="sxs-lookup"><span data-stu-id="5036c-159">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="5036c-p106">**性能** &ndash; 多次往返到托管加载项的 Web 服务器可能会影响加载项的性能。创建需要多个基于邮件或会议操作的加载项时，请考虑性能影响。</span><span class="sxs-lookup"><span data-stu-id="5036c-p106">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="5036c-162">**稍后发送**（仅适用于 Mac）&ndash; 如果有 Onsend 加载项，**稍后发送**功能将不可用。</span><span class="sxs-lookup"><span data-stu-id="5036c-162">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="5036c-163">邮箱类型/模式限制</span><span class="sxs-lookup"><span data-stu-id="5036c-163">Mailbox type/mode limitations</span></span>

<span data-ttu-id="5036c-164">只有 Outlook 网页版、Windows 版和 Mac 版中的用户邮箱支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-164">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="5036c-165">当前不可对以下邮箱类型和模式使用此功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-165">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="5036c-166">共享邮箱\*</span><span class="sxs-lookup"><span data-stu-id="5036c-166">Shared mailboxes\*</span></span>
- <span data-ttu-id="5036c-167">组邮箱</span><span class="sxs-lookup"><span data-stu-id="5036c-167">Group mailboxes</span></span>
- <span data-ttu-id="5036c-168">脱机模式</span><span class="sxs-lookup"><span data-stu-id="5036c-168">Offline mode</span></span>

<span data-ttu-id="5036c-169">如果对这些邮箱场景启用了 Onsend 功能，则 Outlook 将不允许进行发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-169">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="5036c-170">但是，如果用户答复组邮箱中的电子邮件，则 Onsend 加载项将不运行且系统将发送邮件。</span><span class="sxs-lookup"><span data-stu-id="5036c-170">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5036c-171">\* 如果外接程序还 [实现对代理访问方案的支持](delegate-access.md)，则发送时功能应适用于共享邮箱或文件夹。</span><span class="sxs-lookup"><span data-stu-id="5036c-171">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="5036c-172">多个 Onsend 加载项</span><span class="sxs-lookup"><span data-stu-id="5036c-172">Multiple on-send add-ins</span></span>

<span data-ttu-id="5036c-173">如果安装了多个 Onsend 加载项，则加载项将按照从 API `getAppManifestCall` 或 `getExtensibilityContext` 接收到的顺序运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-173">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="5036c-174">如果第一个外接程序允许发送，则第二个外接程序可以更改阻止第一个外接程序进行发送的某些设置。</span><span class="sxs-lookup"><span data-stu-id="5036c-174">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="5036c-175">但是，如果所有已安装的外接程序均允许发送，则第一个外接程序将不会重新运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-175">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="5036c-176">例如，Add-in1 和 Add-in2 均使用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-176">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="5036c-177">首先安装的是 Add-in1，接着安装的是 Add-in2。</span><span class="sxs-lookup"><span data-stu-id="5036c-177">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="5036c-178">Add-in1 验证邮件中出现的 Fabrikam 一词作为外接程序允许发送的条件。</span><span class="sxs-lookup"><span data-stu-id="5036c-178">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="5036c-179">但是，Add-in2 可以删除出现的所有 Fabrikam 词语。</span><span class="sxs-lookup"><span data-stu-id="5036c-179">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="5036c-180">邮件将与已删除 Fabrikam 的所有实例一同发送（归因于 Add-in1 和 Add-in2 的安装顺序）。</span><span class="sxs-lookup"><span data-stu-id="5036c-180">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="5036c-181">部署使用 Onsend 的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="5036c-181">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="5036c-182">建议管理员部署使用 Onsend 功能的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-182">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="5036c-183">管理员必须确保 Onsend 加载项满足以下条件：</span><span class="sxs-lookup"><span data-stu-id="5036c-183">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="5036c-184">任何时候打开撰写项目时均可用（针对电子邮件：新建、回复或转发）。</span><span class="sxs-lookup"><span data-stu-id="5036c-184">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="5036c-185">用户无法关闭或禁用。</span><span class="sxs-lookup"><span data-stu-id="5036c-185">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="5036c-186">安装使用 Onsend 的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="5036c-186">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="5036c-187">Outlook 中的 Onsend 功能要求针对发送事件类型配置加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-187">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="5036c-188">选择要配置的平台。</span><span class="sxs-lookup"><span data-stu-id="5036c-188">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="5036c-189">Web 浏览器 - 经典 Outlook</span><span class="sxs-lookup"><span data-stu-id="5036c-189">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="5036c-190">对于分配了将 *OnSendAddinsEnabled* 标志设置为 **true** 的 Outlook 网页版邮箱策略的用户，系统会为其运行使用 Onsend 功能的 Outlook 网页版（经典）的加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-190">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="5036c-191">若要安装新的外接程序，请运行以下 Exchange Online PowerShell cmdlet。</span><span class="sxs-lookup"><span data-stu-id="5036c-191">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="5036c-192">若要了解如何使用远程 PowerShell 连接到 Exchange Online，请参阅[连接到 Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)。</span><span class="sxs-lookup"><span data-stu-id="5036c-192">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="5036c-193">启用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="5036c-193">Enable the on-send feature</span></span>

<span data-ttu-id="5036c-194">默认情况下，Onsend 功能处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="5036c-194">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="5036c-195">管理员可以通过运行 Exchange Online PowerShell cmdlet 启用 Onsend。</span><span class="sxs-lookup"><span data-stu-id="5036c-195">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="5036c-196">要为所有用户启用 Onsend 加载项，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5036c-196">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="5036c-197">创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-197">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="5036c-198">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-198">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="5036c-199">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-199">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="5036c-200">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-200">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="5036c-201">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="5036c-201">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="5036c-202">为一组用户启用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="5036c-202">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="5036c-203">为特定用户组启用 Onsend 功能的步骤如下。</span><span class="sxs-lookup"><span data-stu-id="5036c-203">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="5036c-204">在此示例中，管理员仅希望在财务用户（其中财务用户属于财务部门）的环境中启用 Outlook 网页版 Onsend 加载项功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-204">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="5036c-205">为该组创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-205">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="5036c-206">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能（有关详细信息，请参阅本文前面介绍的[邮箱类型限制](#multiple-on-send-add-ins)）。</span><span class="sxs-lookup"><span data-stu-id="5036c-206">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="5036c-207">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-207">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="5036c-208">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-208">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="5036c-209">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="5036c-209">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="5036c-210">需要等待 60 分钟该策略才能生效，或重启 Internet Information Services (IIS)。</span><span class="sxs-lookup"><span data-stu-id="5036c-210">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="5036c-211">策略生效后，将为该组启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-211">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="5036c-212">禁用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="5036c-212">Disable the on-send feature</span></span>

<span data-ttu-id="5036c-213">若要禁用用户的 Onsend 功能或分配未启用该标志的 Outlook 网页版邮箱策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="5036c-213">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="5036c-214">在此示例中，该邮箱策略是 *ContosoCorpOWAPolicy*。</span><span class="sxs-lookup"><span data-stu-id="5036c-214">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="5036c-215">有关如何使用 **Set-OwaMailboxPolicy** cmdlet 配置现有 Outlook 网页版邮箱策略的详细信息，请参阅 [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)。</span><span class="sxs-lookup"><span data-stu-id="5036c-215">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="5036c-216">若要禁用所有分配了指定 Outlook 网页版邮箱策略的用户的 Onsend 功能，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="5036c-216">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="5036c-217">Web 浏览器 - 新式 Outlook</span><span class="sxs-lookup"><span data-stu-id="5036c-217">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="5036c-218">对于安装了使用 Onsend 功能的 Outlook 网页版（新式）加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-218">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="5036c-219">但是，如果用户需要运行该加载项来满足合规性标准，则邮箱策略必须将 *OnSendAddinsEnabled* 标志设置为 **true**。</span><span class="sxs-lookup"><span data-stu-id="5036c-219">However, if users are required to run the add-in to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="5036c-220">若要安装新的外接程序，请运行以下 Exchange Online PowerShell cmdlet。</span><span class="sxs-lookup"><span data-stu-id="5036c-220">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="5036c-221">若要了解如何使用远程 PowerShell 连接到 Exchange Online，请参阅[连接到 Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)。</span><span class="sxs-lookup"><span data-stu-id="5036c-221">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="disable-the-on-send-policy"></a><span data-ttu-id="5036c-222">禁用 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="5036c-222">Disable the on-send policy</span></span>

<span data-ttu-id="5036c-223">默认情况下，启用发送策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-223">By default, on-send policy is enabled.</span></span> <span data-ttu-id="5036c-224">若要禁用用户的 Onsend 策略或分配未启用该标志的 Outlook 网页版邮箱策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="5036c-224">To disable the on-send policy for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="5036c-225">在此示例中，该邮箱策略是 *ContosoCorpOWAPolicy*。</span><span class="sxs-lookup"><span data-stu-id="5036c-225">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="5036c-226">有关如何使用 **Set-OwaMailboxPolicy** cmdlet 配置现有 Outlook 网页版邮箱策略的详细信息，请参阅 [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)。</span><span class="sxs-lookup"><span data-stu-id="5036c-226">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="5036c-227">若要禁用所有分配了指定 Outlook 网页版邮箱策略的用户的 Onsend 策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="5036c-227">To disable the on-send policy for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

#### <a name="enable-the-on-send-policy"></a><span data-ttu-id="5036c-228">启用 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="5036c-228">Enable the on-send policy</span></span>

<span data-ttu-id="5036c-229">管理员可以通过运行 Exchange Online PowerShell cmdlet 启用 Onsend。</span><span class="sxs-lookup"><span data-stu-id="5036c-229">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="5036c-230">要为所有用户启用 Onsend 加载项，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5036c-230">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="5036c-231">创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-231">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="5036c-232">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-232">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="5036c-233">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-233">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="5036c-234">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-234">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="5036c-235">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="5036c-235">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-policy-for-a-group-of-users"></a><span data-ttu-id="5036c-236">为一组用户启用 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="5036c-236">Enable the on-send policy for a group of users</span></span>

<span data-ttu-id="5036c-237">为特定用户组启用 Onsend 策略的步骤如下。</span><span class="sxs-lookup"><span data-stu-id="5036c-237">To enable the on-send policy for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="5036c-238">在此示例中，管理员仅希望在财务用户（其中财务用户属于财务部门）的环境中启用 Outlook 网页版 Onsend 加载项策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-238">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="5036c-239">为该组创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-239">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="5036c-240">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能（有关详细信息，请参阅本文前面介绍的[邮箱类型限制](#multiple-on-send-add-ins)）。</span><span class="sxs-lookup"><span data-stu-id="5036c-240">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="5036c-241">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-241">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="5036c-242">启用 Onsend 策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-242">Enable the on-send policy.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="5036c-243">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="5036c-243">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="5036c-244">需要等待 60 分钟该策略才能生效，或重启 Internet Information Services (IIS)。</span><span class="sxs-lookup"><span data-stu-id="5036c-244">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="5036c-245">策略生效后，将为该组强制执行 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="5036c-245">When the policy takes effect, the on-send feature will be enforced for the group.</span></span>

### <a name="windows"></a>[<span data-ttu-id="5036c-246">Windows</span><span class="sxs-lookup"><span data-stu-id="5036c-246">Windows</span></span>](#tab/windows)

<span data-ttu-id="5036c-247">对于安装了使用 Onsend 功能的 Windows 版 Outlook 加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-247">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="5036c-248">但是，如果用户需要运行该加载项来满足合规性标准，则必须在每台适用的计算机上将组策略“**无法加载 Web 扩展时禁用发送**”设置为“**已启用**”。</span><span class="sxs-lookup"><span data-stu-id="5036c-248">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="5036c-249">若要设置邮箱策略，管理员可以下载[管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)，然后通过运行本地组策略编辑器 **(gpedit.msc)** 访问最新的管理模板。</span><span class="sxs-lookup"><span data-stu-id="5036c-249">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="5036c-250">策略的用途</span><span class="sxs-lookup"><span data-stu-id="5036c-250">What the policy does</span></span>

<span data-ttu-id="5036c-251">出于合规性原因，管理员可能需要在用户具有可供运行的最新 Onsend 加载项前，确保其无法发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-251">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="5036c-252">管理员必须启用组策略“**无法加载 Web 扩展时禁用发送**”，以便所有加载项都从 Exchange 进行更新，并可用于在发送时验证每封邮件或每个会议项目是否符合预期的规则和规定。</span><span class="sxs-lookup"><span data-stu-id="5036c-252">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="5036c-253">策略状态</span><span class="sxs-lookup"><span data-stu-id="5036c-253">Policy status</span></span>|<span data-ttu-id="5036c-254">结果</span><span class="sxs-lookup"><span data-stu-id="5036c-254">Result</span></span>|
|---|---|
|<span data-ttu-id="5036c-255">已禁用</span><span class="sxs-lookup"><span data-stu-id="5036c-255">Disabled</span></span>|<span data-ttu-id="5036c-256">允许发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-256">Send allowed.</span></span> <span data-ttu-id="5036c-257">即使尚未从 Exchange 中更新加载项，也可以在不运行 Onsend 加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-257">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="5036c-258">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-258">Enabled</span></span>|<span data-ttu-id="5036c-259">仅当加载项已从 Exchange 更新时才允许发送；否则，将阻止发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-259">Send allowed only when the add-in has been updated from Exchange; otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="5036c-260">管理 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="5036c-260">Manage the on-send policy</span></span>

<span data-ttu-id="5036c-261">默认情况下，Onsend 策略处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="5036c-261">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="5036c-262">管理员可以通过确保用户的组策略设置“**无法加载 Web 扩展时禁用发送**”设置为“**已启用**”来启用 Onsend 策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-262">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="5036c-263">若为用户禁用策略，管理员应将其设置为“**已禁用**”。</span><span class="sxs-lookup"><span data-stu-id="5036c-263">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="5036c-264">若要管理此策略设置，可执行下列操作。</span><span class="sxs-lookup"><span data-stu-id="5036c-264">To manage this policy setting, you can do the following.</span></span>

1. <span data-ttu-id="5036c-265">下载最新的[管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)。</span><span class="sxs-lookup"><span data-stu-id="5036c-265">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="5036c-266">打开本地组策略编辑器 (**gpedit.msc**)。</span><span class="sxs-lookup"><span data-stu-id="5036c-266">Open the Local Group Policy editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="5036c-267">导航到 **“用户配置”>“管理模板”>“Microsoft Outlook 2016”>“安全性”>“信任中心”**。</span><span class="sxs-lookup"><span data-stu-id="5036c-267">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="5036c-268">选择“**无法加载 Web 扩展时禁用发送**”设置。</span><span class="sxs-lookup"><span data-stu-id="5036c-268">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="5036c-269">打开链接以编辑策略设置。</span><span class="sxs-lookup"><span data-stu-id="5036c-269">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="5036c-270">在“**无法加载 Web 扩展时禁用发送**”对话框窗口中，根据需要选择“**已启用**”或“**已禁用**”，然后选择“**确定**”或“**应用**”以使更新生效。</span><span class="sxs-lookup"><span data-stu-id="5036c-270">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="5036c-271">Mac</span><span class="sxs-lookup"><span data-stu-id="5036c-271">Mac</span></span>](#tab/unix)

<span data-ttu-id="5036c-272">对于安装了使用 Onsend 功能的 Mac 版 Outlook 加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-272">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="5036c-273">但是，如果用户需要运行该加载项来满足合规性标准，则必须在每个用户的计算机上应用以下邮箱设置。</span><span class="sxs-lookup"><span data-stu-id="5036c-273">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="5036c-274">此设置或键与 CFPreferences 兼容，这意味着可以使用适用于 Mac 的企业管理软件（例如，Jamf Pro）来对其进行设置。</span><span class="sxs-lookup"><span data-stu-id="5036c-274">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

|||
|:---|:---|
|<span data-ttu-id="5036c-275">**域**</span><span class="sxs-lookup"><span data-stu-id="5036c-275">**Domain**</span></span>|<span data-ttu-id="5036c-276">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="5036c-276">com.microsoft.outlook</span></span>|
|<span data-ttu-id="5036c-277">**键**</span><span class="sxs-lookup"><span data-stu-id="5036c-277">**Key**</span></span>|<span data-ttu-id="5036c-278">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="5036c-278">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="5036c-279">**DataType**</span><span class="sxs-lookup"><span data-stu-id="5036c-279">**DataType**</span></span>|<span data-ttu-id="5036c-280">Boolean</span><span class="sxs-lookup"><span data-stu-id="5036c-280">Boolean</span></span>|
|<span data-ttu-id="5036c-281">**可能的值**</span><span class="sxs-lookup"><span data-stu-id="5036c-281">**Possible values**</span></span>|<span data-ttu-id="5036c-282">false（默认值）</span><span class="sxs-lookup"><span data-stu-id="5036c-282">false (default)</span></span><br><span data-ttu-id="5036c-283">true</span><span class="sxs-lookup"><span data-stu-id="5036c-283">true</span></span>|
|<span data-ttu-id="5036c-284">**可用性**</span><span class="sxs-lookup"><span data-stu-id="5036c-284">**Availability**</span></span>|<span data-ttu-id="5036c-285">16.27</span><span class="sxs-lookup"><span data-stu-id="5036c-285">16.27</span></span>|
|<span data-ttu-id="5036c-286">**备注**</span><span class="sxs-lookup"><span data-stu-id="5036c-286">**Comments**</span></span>|<span data-ttu-id="5036c-287">此键将创建 onSendMailbox 策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-287">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="5036c-288">设置的用途</span><span class="sxs-lookup"><span data-stu-id="5036c-288">What the setting does</span></span>

<span data-ttu-id="5036c-289">出于合规性原因，管理员可能需要在用户具有可供运行的最新 Onsend 加载项前，确保其无法发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-289">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="5036c-290">管理员必须启用键 **OnSendAddinsWaitForLoad**，以便所有加载项都从 Exchange 进行更新，并可用于在发送时验证每封邮件或每个会议项目是否符合预期的规则和规定。</span><span class="sxs-lookup"><span data-stu-id="5036c-290">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="5036c-291">键的状态</span><span class="sxs-lookup"><span data-stu-id="5036c-291">Key's state</span></span>|<span data-ttu-id="5036c-292">结果</span><span class="sxs-lookup"><span data-stu-id="5036c-292">Result</span></span>|
|---|---|
|<span data-ttu-id="5036c-293">false</span><span class="sxs-lookup"><span data-stu-id="5036c-293">false</span></span>|<span data-ttu-id="5036c-294">允许发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-294">Send allowed.</span></span> <span data-ttu-id="5036c-295">即使尚未从 Exchange 中更新加载项，也可以在不运行 Onsend 加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-295">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="5036c-296">true</span><span class="sxs-lookup"><span data-stu-id="5036c-296">true</span></span>|<span data-ttu-id="5036c-297">仅当加载项已从 Exchange 更新时才允许发送；否则，将阻止发送，并且禁用“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="5036c-297">Send allowed only when add-ins have been updated from Exchange; otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="5036c-298">Onsend 功能的应用场景</span><span class="sxs-lookup"><span data-stu-id="5036c-298">On-send feature scenarios</span></span>

<span data-ttu-id="5036c-299">以下是支持和不支持使用 Onsend 功能的加载项的应用场景。</span><span class="sxs-lookup"><span data-stu-id="5036c-299">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="5036c-300">用户邮箱启用了 Onsend 加载项功能，但未安装任何加载项</span><span class="sxs-lookup"><span data-stu-id="5036c-300">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="5036c-301">在这种场景中，用户将能够在不执行任何加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-301">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="5036c-302">用户邮箱启用了 Onsend 加载项功能，并且安装并启用了支持 Onsend 的加载项</span><span class="sxs-lookup"><span data-stu-id="5036c-302">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="5036c-303">外接程序在发送事件期间运行，然后允许或阻止用户发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-303">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="5036c-304">邮箱委派，其中邮箱 1 具有对邮箱 2 的完全访问权限</span><span class="sxs-lookup"><span data-stu-id="5036c-304">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="5036c-305">Web 浏览器（经典 Outlook）</span><span class="sxs-lookup"><span data-stu-id="5036c-305">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="5036c-306">方案</span><span class="sxs-lookup"><span data-stu-id="5036c-306">Scenario</span></span>|<span data-ttu-id="5036c-307">邮箱 1 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="5036c-307">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="5036c-308">邮箱 2 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="5036c-308">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="5036c-309">Outlook Web 会话（经典）</span><span class="sxs-lookup"><span data-stu-id="5036c-309">Outlook web session (classic)</span></span>|<span data-ttu-id="5036c-310">结果</span><span class="sxs-lookup"><span data-stu-id="5036c-310">Result</span></span>|<span data-ttu-id="5036c-311">是否支持？</span><span class="sxs-lookup"><span data-stu-id="5036c-311">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="5036c-312">1</span><span class="sxs-lookup"><span data-stu-id="5036c-312">1</span></span>|<span data-ttu-id="5036c-313">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-313">Enabled</span></span>|<span data-ttu-id="5036c-314">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-314">Enabled</span></span>|<span data-ttu-id="5036c-315">新会话</span><span class="sxs-lookup"><span data-stu-id="5036c-315">New session</span></span>|<span data-ttu-id="5036c-316">邮箱 1 无法从邮箱 2 发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-316">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="5036c-p132">目前尚不支持。可以使用方案 3 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="5036c-p132">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="5036c-319">双面</span><span class="sxs-lookup"><span data-stu-id="5036c-319">2</span></span>|<span data-ttu-id="5036c-320">已禁用</span><span class="sxs-lookup"><span data-stu-id="5036c-320">Disabled</span></span>|<span data-ttu-id="5036c-321">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-321">Enabled</span></span>|<span data-ttu-id="5036c-322">新会话</span><span class="sxs-lookup"><span data-stu-id="5036c-322">New session</span></span>|<span data-ttu-id="5036c-323">邮箱 1 无法从邮箱 2 发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-323">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="5036c-p133">目前尚不支持。可以使用方案 3 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="5036c-p133">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="5036c-326">第三章</span><span class="sxs-lookup"><span data-stu-id="5036c-326">3</span></span>|<span data-ttu-id="5036c-327">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-327">Enabled</span></span>|<span data-ttu-id="5036c-328">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-328">Enabled</span></span>|<span data-ttu-id="5036c-329">同一个会话</span><span class="sxs-lookup"><span data-stu-id="5036c-329">Same session</span></span>|<span data-ttu-id="5036c-330">分配给邮箱 1 的 Onsend 加载项运行 Onsend。</span><span class="sxs-lookup"><span data-stu-id="5036c-330">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="5036c-331">支持。</span><span class="sxs-lookup"><span data-stu-id="5036c-331">Supported.</span></span>|
|<span data-ttu-id="5036c-332">4 </span><span class="sxs-lookup"><span data-stu-id="5036c-332">4</span></span>|<span data-ttu-id="5036c-333">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-333">Enabled</span></span>|<span data-ttu-id="5036c-334">已禁用</span><span class="sxs-lookup"><span data-stu-id="5036c-334">Disabled</span></span>|<span data-ttu-id="5036c-335">新会话</span><span class="sxs-lookup"><span data-stu-id="5036c-335">New session</span></span>|<span data-ttu-id="5036c-336">未运行 Onsend 加载项；邮件或会议项目已发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-336">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="5036c-337">支持。</span><span class="sxs-lookup"><span data-stu-id="5036c-337">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="5036c-338">Web 浏览器（新式 Outlook）、Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="5036c-338">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="5036c-339">若要强制执行 Onsend，管理员应确保对两个邮箱都启用了该策略。</span><span class="sxs-lookup"><span data-stu-id="5036c-339">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="5036c-340">若要了解如何在加载项中支持委派访问，请参阅[在 Outlook 加载项中启用委派访问方案](delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="5036c-340">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="5036c-341">组 1 是新式组邮箱，用户邮箱 1 是组 1 的成员</span><span class="sxs-lookup"><span data-stu-id="5036c-341">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="5036c-342">方案</span><span class="sxs-lookup"><span data-stu-id="5036c-342">Scenario</span></span>|<span data-ttu-id="5036c-343">邮箱 1 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="5036c-343">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="5036c-344">是否启用了 Onsend 加载项？</span><span class="sxs-lookup"><span data-stu-id="5036c-344">On-send add-ins enabled?</span></span>|<span data-ttu-id="5036c-345">邮箱 1 操作</span><span class="sxs-lookup"><span data-stu-id="5036c-345">Mailbox 1 action</span></span>|<span data-ttu-id="5036c-346">结果</span><span class="sxs-lookup"><span data-stu-id="5036c-346">Result</span></span>|<span data-ttu-id="5036c-347">是否支持？</span><span class="sxs-lookup"><span data-stu-id="5036c-347">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="5036c-348">1</span><span class="sxs-lookup"><span data-stu-id="5036c-348">1</span></span>|<span data-ttu-id="5036c-349">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-349">Enabled</span></span>|<span data-ttu-id="5036c-350">是</span><span class="sxs-lookup"><span data-stu-id="5036c-350">Yes</span></span>|<span data-ttu-id="5036c-351">邮箱 1 撰写发送到组 1 的新邮件或会议。</span><span class="sxs-lookup"><span data-stu-id="5036c-351">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="5036c-352">发送期间，Onsend 加载项运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-352">On-send add-ins run during send.</span></span>|<span data-ttu-id="5036c-353">是</span><span class="sxs-lookup"><span data-stu-id="5036c-353">Yes</span></span>|
|<span data-ttu-id="5036c-354">双面</span><span class="sxs-lookup"><span data-stu-id="5036c-354">2</span></span>|<span data-ttu-id="5036c-355">已启用</span><span class="sxs-lookup"><span data-stu-id="5036c-355">Enabled</span></span>|<span data-ttu-id="5036c-356">是</span><span class="sxs-lookup"><span data-stu-id="5036c-356">Yes</span></span>|<span data-ttu-id="5036c-357">邮箱 1 在 Outlook 网页版组 1 的组窗口中撰写发送到组 1 的新邮件或会议。</span><span class="sxs-lookup"><span data-stu-id="5036c-357">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="5036c-358">Onsend 加载项不会在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-358">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="5036c-359">目前尚不支持。</span><span class="sxs-lookup"><span data-stu-id="5036c-359">Not currently supported.</span></span> <span data-ttu-id="5036c-360">可以使用方案 1 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="5036c-360">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="5036c-361">用户邮箱启用了 Onsend 加载项功能/策略，并且安装并启用了支持 Onsend 的加载项，启用了脱机模式</span><span class="sxs-lookup"><span data-stu-id="5036c-361">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="5036c-362">Onsend 加载项将根据用户、加载项后端和 Exchange 的联机状态运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-362">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="5036c-363">用户的状态</span><span class="sxs-lookup"><span data-stu-id="5036c-363">User's state</span></span>

<span data-ttu-id="5036c-364">如果用户处于联机状态，则 Onsend 加载项将在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-364">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="5036c-365">如果用户处于脱机状态，Onsend 加载项不会在发送期间运行，也不会发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-365">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="5036c-366">加载项后端的状态</span><span class="sxs-lookup"><span data-stu-id="5036c-366">Add-in backend's state</span></span>

<span data-ttu-id="5036c-367">如果 Onsend 加载项的后端处于联机状态且可访问，则将运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-367">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="5036c-368">如果后端处于脱机状态，则将禁用发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-368">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="5036c-369">Exchange 的状态</span><span class="sxs-lookup"><span data-stu-id="5036c-369">Exchange's state</span></span>

<span data-ttu-id="5036c-370">如果 Exchange 服务器处于联机状态且可访问，则 Onsend 加载项将在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-370">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="5036c-371">如果 Onsend 加载项无法访问 Exchange 并且已启用适用的策略或 cmdlet，则将禁用发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-371">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="5036c-372">在处于任何脱机状态的 Mac 上，“**发送**”按钮（或现有会议的“**发送更新**”按钮）将被禁用，并显示当用户脱机时其组织不允许发送的通知。</span><span class="sxs-lookup"><span data-stu-id="5036c-372">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="5036c-373">在发送外接程序正在运行时，用户可以编辑项目</span><span class="sxs-lookup"><span data-stu-id="5036c-373">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="5036c-374">在发送外接程序处理项目时，用户可以通过添加（例如，不适当的文本或附件）来编辑项目。</span><span class="sxs-lookup"><span data-stu-id="5036c-374">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="5036c-375">如果要在加载项在发送过程中进行处理时阻止用户编辑项目，可以使用对话框实施解决方法。</span><span class="sxs-lookup"><span data-stu-id="5036c-375">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="5036c-376">在您的发送处理程序中：</span><span class="sxs-lookup"><span data-stu-id="5036c-376">In your on-send handler:</span></span>

1. <span data-ttu-id="5036c-377">调用 [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview#displaydialogasync-startaddress--options--callback-) 以打开对话框，以便禁用鼠标单击和键击。</span><span class="sxs-lookup"><span data-stu-id="5036c-377">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="5036c-378">若要在 web 上的 Outlook 中获取此行为，您应在调用的参数中将 [displayInIframe 属性](/javascript/api/office/office.dialogoptions?view=outlook-js-preview#displayiniframe) 设置为 `true` `options` `displayDialogAsync` 。</span><span class="sxs-lookup"><span data-stu-id="5036c-378">To get this behavior in Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="5036c-379">实现项目的处理。</span><span class="sxs-lookup"><span data-stu-id="5036c-379">Implement processing of the item.</span></span>
1. <span data-ttu-id="5036c-380">关闭该对话框。</span><span class="sxs-lookup"><span data-stu-id="5036c-380">Close the dialog.</span></span> <span data-ttu-id="5036c-381">此外，处理当用户关闭对话框时会发生的情况。</span><span class="sxs-lookup"><span data-stu-id="5036c-381">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="5036c-382">代码示例</span><span class="sxs-lookup"><span data-stu-id="5036c-382">Code examples</span></span>

<span data-ttu-id="5036c-383">以下代码示例说明如何创建一个简单的 Onsend 加载项。</span><span class="sxs-lookup"><span data-stu-id="5036c-383">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="5036c-384">若要下载这些示例所基于的代码示例，请参阅 [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send)。</span><span class="sxs-lookup"><span data-stu-id="5036c-384">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="5036c-385">如果将对话框与发送时事件结合使用，请确保在完成该事件之前关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="5036c-385">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="5036c-386">清单、版本重写和事件</span><span class="sxs-lookup"><span data-stu-id="5036c-386">Manifest, version override, and event</span></span>

<span data-ttu-id="5036c-387">[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) 代码示例包括两个清单：</span><span class="sxs-lookup"><span data-stu-id="5036c-387">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="5036c-388">`Contoso Message Body Checker.xml` &ndash; 展示了如何在发送时检查邮件正文是否包含限制字词或敏感信息。</span><span class="sxs-lookup"><span data-stu-id="5036c-388">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="5036c-389">`Contoso Subject and CC Checker.xml` &ndash; 展示了如何将收件人添加到抄送行，并在发送时验证邮件是否包含主题行。</span><span class="sxs-lookup"><span data-stu-id="5036c-389">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="5036c-390">在 `Contoso Message Body Checker.xml` 清单文件中，将包含在 `ItemSend` 事件中应调用的函数文件和函数名称。</span><span class="sxs-lookup"><span data-stu-id="5036c-390">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="5036c-391">该操作将同步运行。</span><span class="sxs-lookup"><span data-stu-id="5036c-391">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="5036c-392">如果使用 Visual Studio 2019 开发你的发送外接程序，则可能会收到类似于以下的验证警告： "这是一个无效的 xsi： type ' http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events "。 "若要解决此问题，您需要在 [有关此警告的博客](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/)中提供了 MailAppVersionOverridesV1_1 的较新版本的 .Xsd 作为 GitHub gist 提供。</span><span class="sxs-lookup"><span data-stu-id="5036c-392">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="5036c-393">对于 `Contoso Subject and CC Checker.xml` 清单文件，以下示例中显示了邮件发送事件中要调用的函数文件和函数名称。</span><span class="sxs-lookup"><span data-stu-id="5036c-393">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="5036c-394">Onsend API 需要 `VersionOverrides v1_1`。</span><span class="sxs-lookup"><span data-stu-id="5036c-394">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="5036c-395">以下显示如何在清单中添加 `VersionOverrides` 节点。</span><span class="sxs-lookup"><span data-stu-id="5036c-395">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="5036c-396">有关详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="5036c-396">For more information, see the following:</span></span>
> - [<span data-ttu-id="5036c-397">Outlook 外接程序清单</span><span class="sxs-lookup"><span data-stu-id="5036c-397">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="5036c-398">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="5036c-398">Office Add-ins XML manifest</span></span>](../overview/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="5036c-399">`Event` 和 `item` 对象以及 `body.getAsync` 和 `body.setAsync` 方法</span><span class="sxs-lookup"><span data-stu-id="5036c-399">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="5036c-400">若要访问当前选择的邮件或会议项目（在本示例中为新撰写的邮件），请使用 `Office.context.mailbox.item` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="5036c-400">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="5036c-401">`ItemSend` 事件由 Onsend 功能自动传递到清单中指定的函数&mdash;在本示例中为 `validateBody` 函数。</span><span class="sxs-lookup"><span data-stu-id="5036c-401">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="5036c-402">`validateBody` 函数以指定格式 (HTML) 获取当前正文，并在回调方法中传递代码想要访问的 `ItemSend` 事件对象。</span><span class="sxs-lookup"><span data-stu-id="5036c-402">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="5036c-403">除 `getAsync` 方法之外，`Body` 对象还提供了 `setAsync` 方法，可用于将正文替换为指定的文本。</span><span class="sxs-lookup"><span data-stu-id="5036c-403">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="5036c-404">有关详细信息，请参阅 [Event 对象](/javascript/api/office/office.addincommands.event)和 [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="5036c-404">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="5036c-405">`NotificationMessages` 对象和 `event.completed` 方法</span><span class="sxs-lookup"><span data-stu-id="5036c-405">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="5036c-406">`checkBodyOnlyOnSendCallBack` 函数使用正则表达式来确定邮件正文是否包含禁止使用的词语。</span><span class="sxs-lookup"><span data-stu-id="5036c-406">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="5036c-407">如果该函数发现受限词语数组的匹配项，则将阻止发送电子邮件，并通过信息栏通知发件人。</span><span class="sxs-lookup"><span data-stu-id="5036c-407">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="5036c-408">为了做到这一点，它使用 `Item` 对象的 `notificationMessages` 属性来返回 `NotificationMessages` 对象。</span><span class="sxs-lookup"><span data-stu-id="5036c-408">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="5036c-409">然后，通过调用 `addAsync` 方法向该项目添加通知，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="5036c-409">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="5036c-410">以下是 `addAsync` 方法的参数：</span><span class="sxs-lookup"><span data-stu-id="5036c-410">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="5036c-411">`NoSend` &ndash; 一个字符串，即开发人员指定用于引用通知邮件的密钥。</span><span class="sxs-lookup"><span data-stu-id="5036c-411">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="5036c-412">可用于在以后修改此邮件。</span><span class="sxs-lookup"><span data-stu-id="5036c-412">You can use it to modify this message later.</span></span> <span data-ttu-id="5036c-413">密钥长度不能超过32个字符。</span><span class="sxs-lookup"><span data-stu-id="5036c-413">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="5036c-414">`type` &ndash; JSON 对象参数的一个属性。</span><span class="sxs-lookup"><span data-stu-id="5036c-414">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="5036c-415">表示邮件的类型；类型对应于 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) 枚举的值。</span><span class="sxs-lookup"><span data-stu-id="5036c-415">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="5036c-416">可能的值是进度指示器、信息消息或错误消息。</span><span class="sxs-lookup"><span data-stu-id="5036c-416">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="5036c-417">在此示例中，`type` 是错误消息。</span><span class="sxs-lookup"><span data-stu-id="5036c-417">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="5036c-418">`message` &ndash; JSON 对象参数的一个属性。</span><span class="sxs-lookup"><span data-stu-id="5036c-418">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="5036c-419">在此示例中，`message` 是通知邮件的文本。</span><span class="sxs-lookup"><span data-stu-id="5036c-419">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="5036c-420">为表明加载项对由发送操作触发的 `ItemSend` 事件的处理已完成，请调用 `event.completed({allowEvent:Boolean})` 方法。</span><span class="sxs-lookup"><span data-stu-id="5036c-420">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="5036c-421">`allowEvent` 属性是一个布尔值。</span><span class="sxs-lookup"><span data-stu-id="5036c-421">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="5036c-422">如果设置为 `true`，则允许发送。</span><span class="sxs-lookup"><span data-stu-id="5036c-422">If set to `true`, send is allowed.</span></span> <span data-ttu-id="5036c-423">如果设置为 `false`，则将阻止发送电子邮件。</span><span class="sxs-lookup"><span data-stu-id="5036c-423">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="5036c-424">有关详细信息，请参阅 [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [completed](/javascript/api/office/office.addincommands.event)。</span><span class="sxs-lookup"><span data-stu-id="5036c-424">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="5036c-425">`replaceAsync`、`removeAsync` 和 `getAllAsync` 方法</span><span class="sxs-lookup"><span data-stu-id="5036c-425">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="5036c-426">除了 `addAsync` 方法之外，`NotificationMessages` 对象还包括 `replaceAsync`、`removeAsync` 和 `getAllAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="5036c-426">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="5036c-427">此代码示例中不使用这些方法。</span><span class="sxs-lookup"><span data-stu-id="5036c-427">These methods are not used in this code sample.</span></span>  <span data-ttu-id="5036c-428">有关详细信息，请参阅 [NotificationMessages](/javascript/api/outlook/office.NotificationMessages)。</span><span class="sxs-lookup"><span data-stu-id="5036c-428">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="5036c-429">主题和抄送检查器代码</span><span class="sxs-lookup"><span data-stu-id="5036c-429">Subject and CC checker code</span></span>

<span data-ttu-id="5036c-430">以下代码示例介绍如何将收件人添加到抄送行，并验证邮件在发送时是否包含主题。</span><span class="sxs-lookup"><span data-stu-id="5036c-430">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="5036c-431">此示例使用 Onsend 功能允许或禁止发送电子邮件。</span><span class="sxs-lookup"><span data-stu-id="5036c-431">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="5036c-p153">若要详细了解如何将收件人添加到抄送行、验证电子邮件在发送时是否包主题行，以及查看可以使用的 API，请参阅 [Outlook-Add-in-On-Send 示例](https://github.com/OfficeDev/Outlook-Add-in-On-Send)。已充分注释代码。</span><span class="sxs-lookup"><span data-stu-id="5036c-p153">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="5036c-434">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5036c-434">See also</span></span>

- [<span data-ttu-id="5036c-435">Outlook 加载项体系结构和功能概述</span><span class="sxs-lookup"><span data-stu-id="5036c-435">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="5036c-436">加载项命令演示 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="5036c-436">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)
