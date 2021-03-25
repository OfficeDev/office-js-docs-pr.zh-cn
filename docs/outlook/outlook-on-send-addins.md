---
title: Outlook 加载项的 Onsend 功能
description: 提供了一种处理项目或阻止用户进行特定操作的方法，并允许加载项在发送时设置某些属性。
ms.date: 03/17/2021
localization_priority: Normal
ms.openlocfilehash: 70e255601fd36a2f9101d56161846616691f5100
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178053"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="f9172-103">Outlook 加载项的 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="f9172-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="f9172-p101">Outlook 加载项的 Onsend 功能提供了一种处理邮件或会议项目，或阻止用户进行特定操作的方法，并允许加载项在发送时设置某些属性。例如，可以使用 Onsend 功能执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="f9172-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="f9172-106">防止用户发送敏感信息或将主题行留空。</span><span class="sxs-lookup"><span data-stu-id="f9172-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="f9172-107">将特定的收件人添加到邮件中的“抄送”行中，或添加到会议中的“可选收件人”行中。</span><span class="sxs-lookup"><span data-stu-id="f9172-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="f9172-108">on-send 功能是由事件类型 `ItemSend` 触发的，无 UI。</span><span class="sxs-lookup"><span data-stu-id="f9172-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="f9172-109">有关 Onsend 功能的限制信息，请参阅本文稍后部分中介绍的[限制](#limitations)。</span><span class="sxs-lookup"><span data-stu-id="f9172-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="f9172-110">支持的客户端和平台</span><span class="sxs-lookup"><span data-stu-id="f9172-110">Supported clients and platforms</span></span>

<span data-ttu-id="f9172-111">下表显示了 Ons onss 功能支持的客户端-服务器组合。</span><span class="sxs-lookup"><span data-stu-id="f9172-111">The following table shows supported client-server combinations for the on-send feature.</span></span> <span data-ttu-id="f9172-112">不支持排除的组合。</span><span class="sxs-lookup"><span data-stu-id="f9172-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="f9172-113">Client</span><span class="sxs-lookup"><span data-stu-id="f9172-113">Client</span></span> | <span data-ttu-id="f9172-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="f9172-114">Exchange Online</span></span> | <span data-ttu-id="f9172-115">Exchange 2016 内部部署</span><span class="sxs-lookup"><span data-stu-id="f9172-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="f9172-116"> (累积更新 6 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="f9172-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="f9172-117">本地 Exchange 2019</span><span class="sxs-lookup"><span data-stu-id="f9172-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="f9172-118"> (累积更新 1 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="f9172-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="f9172-119">Windows：</span><span class="sxs-lookup"><span data-stu-id="f9172-119">Windows:</span></span><br><span data-ttu-id="f9172-120">版本 1910 (内部版本 12130.20272) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="f9172-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="f9172-121">是</span><span class="sxs-lookup"><span data-stu-id="f9172-121">Yes</span></span>|<span data-ttu-id="f9172-122">是</span><span class="sxs-lookup"><span data-stu-id="f9172-122">Yes</span></span>|<span data-ttu-id="f9172-123">是</span><span class="sxs-lookup"><span data-stu-id="f9172-123">Yes</span></span>|
|<span data-ttu-id="f9172-124">Mac：</span><span class="sxs-lookup"><span data-stu-id="f9172-124">Mac:</span></span><br><span data-ttu-id="f9172-125">内部版本 16.30 或更高版本</span><span class="sxs-lookup"><span data-stu-id="f9172-125">build 16.30 or later</span></span>|<span data-ttu-id="f9172-126">是</span><span class="sxs-lookup"><span data-stu-id="f9172-126">Yes</span></span>|<span data-ttu-id="f9172-127">否</span><span class="sxs-lookup"><span data-stu-id="f9172-127">No</span></span>|<span data-ttu-id="f9172-128">否</span><span class="sxs-lookup"><span data-stu-id="f9172-128">No</span></span>|
|<span data-ttu-id="f9172-129">Web 浏览器：</span><span class="sxs-lookup"><span data-stu-id="f9172-129">Web browser:</span></span><br><span data-ttu-id="f9172-130">新式 Outlook UI</span><span class="sxs-lookup"><span data-stu-id="f9172-130">modern Outlook UI</span></span>|<span data-ttu-id="f9172-131">是</span><span class="sxs-lookup"><span data-stu-id="f9172-131">Yes</span></span>|<span data-ttu-id="f9172-132">不适用</span><span class="sxs-lookup"><span data-stu-id="f9172-132">Not applicable</span></span>|<span data-ttu-id="f9172-133">不适用</span><span class="sxs-lookup"><span data-stu-id="f9172-133">Not applicable</span></span>|
|<span data-ttu-id="f9172-134">Web 浏览器：</span><span class="sxs-lookup"><span data-stu-id="f9172-134">Web browser:</span></span><br><span data-ttu-id="f9172-135">经典 Outlook UI</span><span class="sxs-lookup"><span data-stu-id="f9172-135">classic Outlook UI</span></span>|<span data-ttu-id="f9172-136">不适用</span><span class="sxs-lookup"><span data-stu-id="f9172-136">Not applicable</span></span>|<span data-ttu-id="f9172-137">是</span><span class="sxs-lookup"><span data-stu-id="f9172-137">Yes</span></span>|<span data-ttu-id="f9172-138">是</span><span class="sxs-lookup"><span data-stu-id="f9172-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="f9172-139">Ons ons on-send 功能在要求集 1.8 中正式发布， ([当前服务器](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) 和客户端支持，了解) 。</span><span class="sxs-lookup"><span data-stu-id="f9172-139">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="f9172-140">但是，请注意，功能的支持矩阵是要求集的超集。</span><span class="sxs-lookup"><span data-stu-id="f9172-140">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9172-141">AppSource 中不允许使用 Ons onss 功能 [加载项](https://appsource.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="f9172-141">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="f9172-142">Onsend 功能的工作原理</span><span class="sxs-lookup"><span data-stu-id="f9172-142">How does the on-send feature work?</span></span>

<span data-ttu-id="f9172-143">可使用 Onsend 功能生成集成了 `ItemSend` 同步事件的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-143">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="f9172-144">此事件检测到用户正在按“**发送**”按钮（或现有会议的“**发送更新**”按钮），并且如果验证失败，则可用于阻止该项目发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-144">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="f9172-145">例如，当用户触发邮件发送事件时，使用 Onsend 功能的 Outlook 加载项可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="f9172-145">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="f9172-146">读取和验证电子邮件内容</span><span class="sxs-lookup"><span data-stu-id="f9172-146">Read and validate the email message contents</span></span>
- <span data-ttu-id="f9172-147">验证邮件是否包含主题行</span><span class="sxs-lookup"><span data-stu-id="f9172-147">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="f9172-148">设置预先确定的收件人</span><span class="sxs-lookup"><span data-stu-id="f9172-148">Set a predetermined recipient</span></span>

<span data-ttu-id="f9172-149">触发发送事件时，在 Outlook 客户端完成验证，外接程序最多有 5 分钟才能退出。如果验证失败，将阻止发送项目，并且信息栏中会显示一条错误消息，提示用户采取操作。</span><span class="sxs-lookup"><span data-stu-id="f9172-149">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="f9172-150">在 Outlook 网页 Outlook 中，当 Onss onsed 功能在 Outlook 浏览器选项卡内撰写的邮件中触发时，该项目会弹出到其自己的浏览器窗口或选项卡，以便完成验证和其他处理。</span><span class="sxs-lookup"><span data-stu-id="f9172-150">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="f9172-151">以下屏幕截图显示了通知发件人添加主题的信息栏。</span><span class="sxs-lookup"><span data-stu-id="f9172-151">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![屏幕截图显示一个错误消息，提示用户输入缺失的主题行](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="f9172-153">以下屏幕截图显示了一个信息栏，通知发件人已找到禁止使用的词语。</span><span class="sxs-lookup"><span data-stu-id="f9172-153">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![屏幕截图显示一条错误消息，告诉用户已找到禁止使用的词语](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="f9172-155">限制</span><span class="sxs-lookup"><span data-stu-id="f9172-155">Limitations</span></span>

<span data-ttu-id="f9172-156">Onsend 功能目前具有以下限制。</span><span class="sxs-lookup"><span data-stu-id="f9172-156">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="f9172-157">**Append-on-send** 功能 &ndash; 如果调用 [body。AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) 在 Onsend 处理程序中返回错误。</span><span class="sxs-lookup"><span data-stu-id="f9172-157">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="f9172-158">**AppSource** &ndash; 无法在 [AppSource](https://appsource.microsoft.com) 中发布使用 Onsend 功能的 Outlook 加载项，因为它们将无法通过 AppSource 验证。</span><span class="sxs-lookup"><span data-stu-id="f9172-158">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="f9172-159">使用 Onsend 功能的加载项应由管理员部署。</span><span class="sxs-lookup"><span data-stu-id="f9172-159">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="f9172-160">**清单**&ndash; - 每个加载项仅支持一个 `ItemSend` 事件。</span><span class="sxs-lookup"><span data-stu-id="f9172-160">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="f9172-161">如果清单中有两个或多个 `ItemSend` 事件，则该清单将无法通过验证。</span><span class="sxs-lookup"><span data-stu-id="f9172-161">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="f9172-p107">**性能** &ndash; 多次往返到托管加载项的 Web 服务器可能会影响加载项的性能。创建需要多个基于邮件或会议操作的加载项时，请考虑性能影响。</span><span class="sxs-lookup"><span data-stu-id="f9172-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="f9172-164">**稍后发送**（仅适用于 Mac）&ndash; 如果有 Onsend 加载项，**稍后发送** 功能将不可用。</span><span class="sxs-lookup"><span data-stu-id="f9172-164">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="f9172-165">此外，不建议在 Onss ons 发送事件处理程序中调用 ，因为关闭项目应在事件完成后 `item.close()` 自动发生。</span><span class="sxs-lookup"><span data-stu-id="f9172-165">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="f9172-166">邮箱类型/模式限制</span><span class="sxs-lookup"><span data-stu-id="f9172-166">Mailbox type/mode limitations</span></span>

<span data-ttu-id="f9172-167">只有 Outlook 网页版、Windows 版和 Mac 版中的用户邮箱支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-167">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="f9172-168">当前不可对以下邮箱类型和模式使用此功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-168">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="f9172-169">共享邮箱\*</span><span class="sxs-lookup"><span data-stu-id="f9172-169">Shared mailboxes\*</span></span>
- <span data-ttu-id="f9172-170">组邮箱</span><span class="sxs-lookup"><span data-stu-id="f9172-170">Group mailboxes</span></span>
- <span data-ttu-id="f9172-171">脱机模式</span><span class="sxs-lookup"><span data-stu-id="f9172-171">Offline mode</span></span>

<span data-ttu-id="f9172-172">如果对这些邮箱场景启用了 Onsend 功能，则 Outlook 将不允许进行发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-172">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="f9172-173">但是，如果用户答复组邮箱中的电子邮件，则 Onsend 加载项将不运行且系统将发送邮件。</span><span class="sxs-lookup"><span data-stu-id="f9172-173">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9172-174">\*如果加载项还实现了对委派访问方案的支持，Onss ons ons functionality should work on shared mailboxes or [folders.](delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="f9172-174">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="f9172-175">多个 Onsend 加载项</span><span class="sxs-lookup"><span data-stu-id="f9172-175">Multiple on-send add-ins</span></span>

<span data-ttu-id="f9172-176">如果安装了多个 Onsend 加载项，则加载项将按照从 API `getAppManifestCall` 或 `getExtensibilityContext` 接收到的顺序运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-176">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="f9172-177">如果第一个外接程序允许发送，则第二个外接程序可以更改阻止第一个外接程序进行发送的某些设置。</span><span class="sxs-lookup"><span data-stu-id="f9172-177">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="f9172-178">但是，如果所有已安装的外接程序均允许发送，则第一个外接程序将不会重新运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-178">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="f9172-179">例如，Add-in1 和 Add-in2 均使用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-179">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="f9172-180">首先安装的是 Add-in1，接着安装的是 Add-in2。</span><span class="sxs-lookup"><span data-stu-id="f9172-180">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="f9172-181">Add-in1 验证邮件中出现的 Fabrikam 一词作为外接程序允许发送的条件。</span><span class="sxs-lookup"><span data-stu-id="f9172-181">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="f9172-182">但是，Add-in2 可以删除出现的所有 Fabrikam 词语。</span><span class="sxs-lookup"><span data-stu-id="f9172-182">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="f9172-183">邮件将与已删除 Fabrikam 的所有实例一同发送（归因于 Add-in1 和 Add-in2 的安装顺序）。</span><span class="sxs-lookup"><span data-stu-id="f9172-183">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="f9172-184">部署使用 Onsend 的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="f9172-184">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="f9172-185">建议管理员部署使用 Onsend 功能的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-185">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="f9172-186">管理员必须确保 Onsend 加载项满足以下条件：</span><span class="sxs-lookup"><span data-stu-id="f9172-186">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="f9172-187">任何时候打开撰写项目时均可用（针对电子邮件：新建、回复或转发）。</span><span class="sxs-lookup"><span data-stu-id="f9172-187">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="f9172-188">用户无法关闭或禁用。</span><span class="sxs-lookup"><span data-stu-id="f9172-188">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="f9172-189">安装使用 Onsend 的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="f9172-189">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="f9172-190">Outlook 中的 Onsend 功能要求针对发送事件类型配置加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-190">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="f9172-191">选择要配置的平台。</span><span class="sxs-lookup"><span data-stu-id="f9172-191">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="f9172-192">Web 浏览器 - 经典 Outlook</span><span class="sxs-lookup"><span data-stu-id="f9172-192">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="f9172-193">对于分配了将 *OnSendAddinsEnabled* 标志设置为 **true** 的 Outlook 网页版邮箱策略的用户，系统会为其运行使用 Onsend 功能的 Outlook 网页版（经典）的加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-193">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="f9172-194">若要安装新的外接程序，请运行以下 Exchange Online PowerShell cmdlet。</span><span class="sxs-lookup"><span data-stu-id="f9172-194">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="f9172-195">若要了解如何使用远程 PowerShell 连接到 Exchange Online，请参阅[连接到 Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)。</span><span class="sxs-lookup"><span data-stu-id="f9172-195">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="f9172-196">启用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="f9172-196">Enable the on-send feature</span></span>

<span data-ttu-id="f9172-197">默认情况下，Onsend 功能处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="f9172-197">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="f9172-198">管理员可以通过运行 Exchange Online PowerShell cmdlet 启用 Onsend。</span><span class="sxs-lookup"><span data-stu-id="f9172-198">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="f9172-199">要为所有用户启用 Onsend 加载项，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="f9172-199">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="f9172-200">创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-200">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="f9172-201">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-201">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="f9172-202">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-202">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="f9172-203">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-203">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="f9172-204">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="f9172-204">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="f9172-205">为一组用户启用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="f9172-205">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="f9172-206">为特定用户组启用 Onsend 功能的步骤如下。</span><span class="sxs-lookup"><span data-stu-id="f9172-206">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="f9172-207">在此示例中，管理员仅希望在财务用户（其中财务用户属于财务部门）的环境中启用 Outlook 网页版 Onsend 加载项功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-207">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="f9172-208">为该组创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-208">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="f9172-209">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能（有关详细信息，请参阅本文前面介绍的[邮箱类型限制](#multiple-on-send-add-ins)）。</span><span class="sxs-lookup"><span data-stu-id="f9172-209">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="f9172-210">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-210">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="f9172-211">启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-211">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="f9172-212">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="f9172-212">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="f9172-213">需要等待 60 分钟该策略才能生效，或重启 Internet Information Services (IIS)。</span><span class="sxs-lookup"><span data-stu-id="f9172-213">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="f9172-214">策略生效后，将为该组启用 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-214">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="f9172-215">禁用 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="f9172-215">Disable the on-send feature</span></span>

<span data-ttu-id="f9172-216">若要禁用用户的 Onsend 功能或分配未启用该标志的 Outlook 网页版邮箱策略，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="f9172-216">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="f9172-217">在此示例中，该邮箱策略是 *ContosoCorpOWAPolicy*。</span><span class="sxs-lookup"><span data-stu-id="f9172-217">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="f9172-218">有关如何使用 **Set-OwaMailboxPolicy** cmdlet 配置现有 Outlook 网页版邮箱策略的详细信息，请参阅 [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)。</span><span class="sxs-lookup"><span data-stu-id="f9172-218">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="f9172-219">若要禁用所有分配了指定 Outlook 网页版邮箱策略的用户的 Onsend 功能，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="f9172-219">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="f9172-220">Web 浏览器 - 新式 Outlook</span><span class="sxs-lookup"><span data-stu-id="f9172-220">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="f9172-221">对于安装了使用 Onsend 功能的 Outlook 网页版（新式）加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-221">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="f9172-222">但是，如果用户需要运行 Onsend 外接程序以满足合规性标准，则邮箱策略必须将 *OnSendAddinsEnabled* 标志设置为 ，以便不允许在外接程序在发送时编辑项目。 `true`</span><span class="sxs-lookup"><span data-stu-id="f9172-222">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="f9172-223">若要安装新的外接程序，请运行以下 Exchange Online PowerShell cmdlet。</span><span class="sxs-lookup"><span data-stu-id="f9172-223">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="f9172-224">若要了解如何使用远程 PowerShell 连接到 Exchange Online，请参阅[连接到 Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)。</span><span class="sxs-lookup"><span data-stu-id="f9172-224">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="f9172-225">启用 On-send 标志</span><span class="sxs-lookup"><span data-stu-id="f9172-225">Enable the on-send flag</span></span>

<span data-ttu-id="f9172-226">管理员可以通过运行 Exchange Online PowerShell cmdlet 强制实施 Onss onss 合规性。</span><span class="sxs-lookup"><span data-stu-id="f9172-226">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="f9172-227">对于所有用户，若要在处理 Onss on-send 外接程序时禁止编辑：</span><span class="sxs-lookup"><span data-stu-id="f9172-227">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="f9172-228">创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-228">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="f9172-229">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能。</span><span class="sxs-lookup"><span data-stu-id="f9172-229">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="f9172-230">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-230">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="f9172-231">在发送时强制执行合规性。</span><span class="sxs-lookup"><span data-stu-id="f9172-231">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="f9172-232">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="f9172-232">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="f9172-233">为一组用户打开 On-send 标志</span><span class="sxs-lookup"><span data-stu-id="f9172-233">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="f9172-234">若要对一组特定用户强制执行 On-send 合规性，步骤如下。</span><span class="sxs-lookup"><span data-stu-id="f9172-234">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="f9172-235">在此示例中，管理员仅希望在财务用户（其中财务用户属于财务部门）的环境中启用 Outlook 网页版 Onsend 加载项策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-235">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="f9172-236">为该组创建新的 Outlook 网页版邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-236">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="f9172-237">管理员可以使用现有策略，但只有某些邮箱类型才支持 Onsend 功能（有关详细信息，请参阅本文前面介绍的[邮箱类型限制](#multiple-on-send-add-ins)）。</span><span class="sxs-lookup"><span data-stu-id="f9172-237">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="f9172-238">系统将默认阻止 Outlook 网页版中不受支持的邮箱进行发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-238">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="f9172-239">在发送时强制执行合规性。</span><span class="sxs-lookup"><span data-stu-id="f9172-239">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="f9172-240">将策略分配给用户。</span><span class="sxs-lookup"><span data-stu-id="f9172-240">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="f9172-241">需要等待 60 分钟该策略才能生效，或重启 Internet Information Services (IIS)。</span><span class="sxs-lookup"><span data-stu-id="f9172-241">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="f9172-242">当策略生效时，将为该组强制执行 On-send 合规性。</span><span class="sxs-lookup"><span data-stu-id="f9172-242">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="f9172-243">关闭 On-send 标志</span><span class="sxs-lookup"><span data-stu-id="f9172-243">Turn off the on-send flag</span></span>

<span data-ttu-id="f9172-244">若要关闭用户的 Onss ons send 合规性强制，请通过运行以下 cmdlet 分配未启用该标志的 Outlook 网页邮箱策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-244">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="f9172-245">在此示例中，该邮箱策略是 *ContosoCorpOWAPolicy*。</span><span class="sxs-lookup"><span data-stu-id="f9172-245">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="f9172-246">有关如何使用 **Set-OwaMailboxPolicy** cmdlet 配置现有 Outlook 网页版邮箱策略的详细信息，请参阅 [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)。</span><span class="sxs-lookup"><span data-stu-id="f9172-246">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="f9172-247">若要为分配了特定 Outlook 网页邮箱策略的所有用户禁用 Onss onsook 合规性强制，请运行以下 cmdlet。</span><span class="sxs-lookup"><span data-stu-id="f9172-247">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="f9172-248">Windows</span><span class="sxs-lookup"><span data-stu-id="f9172-248">Windows</span></span>](#tab/windows)

<span data-ttu-id="f9172-249">对于安装了使用 Onsend 功能的 Windows 版 Outlook 加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-249">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="f9172-250">但是，如果用户需要运行该加载项来满足合规性标准，则必须在每台适用的计算机上将组策略“**无法加载 Web 扩展时禁用发送**”设置为“**已启用**”。</span><span class="sxs-lookup"><span data-stu-id="f9172-250">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="f9172-251">若要设置邮箱策略，管理员可以下载管理模板工具 [](https://www.microsoft.com/download/details.aspx?id=49030)，然后通过运行本地组策略编辑器 **gpedit.msc** 来访问最新的管理模板。</span><span class="sxs-lookup"><span data-stu-id="f9172-251">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="f9172-252">策略的用途</span><span class="sxs-lookup"><span data-stu-id="f9172-252">What the policy does</span></span>

<span data-ttu-id="f9172-253">出于合规性原因，管理员可能需要在用户具有可供运行的最新 Onsend 加载项前，确保其无法发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="f9172-253">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="f9172-254">管理员必须启用组策略“**无法加载 Web 扩展时禁用发送**”，以便所有加载项都从 Exchange 进行更新，并可用于在发送时验证每封邮件或每个会议项目是否符合预期的规则和规定。</span><span class="sxs-lookup"><span data-stu-id="f9172-254">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="f9172-255">策略状态</span><span class="sxs-lookup"><span data-stu-id="f9172-255">Policy status</span></span>|<span data-ttu-id="f9172-256">结果</span><span class="sxs-lookup"><span data-stu-id="f9172-256">Result</span></span>|
|---|---|
|<span data-ttu-id="f9172-257">已禁用</span><span class="sxs-lookup"><span data-stu-id="f9172-257">Disabled</span></span>|<span data-ttu-id="f9172-258">当前下载的 Ons ons ons 外接程序清单 (在发送的邮件或会议项目) 运行的最新版本。</span><span class="sxs-lookup"><span data-stu-id="f9172-258">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="f9172-259">这是默认状态/行为。</span><span class="sxs-lookup"><span data-stu-id="f9172-259">This is the default status/behavior.</span></span>|
|<span data-ttu-id="f9172-260">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-260">Enabled</span></span>|<span data-ttu-id="f9172-261">从 Exchange 下载 Ons ons 外接程序的最新清单后，外接程序将运行在要发送的邮件或会议项目上。</span><span class="sxs-lookup"><span data-stu-id="f9172-261">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="f9172-262">否则，将阻止发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-262">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="f9172-263">管理 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="f9172-263">Manage the on-send policy</span></span>

<span data-ttu-id="f9172-264">默认情况下，Onsend 策略处于禁用状态。</span><span class="sxs-lookup"><span data-stu-id="f9172-264">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="f9172-265">管理员可以通过确保用户的组策略设置“**无法加载 Web 扩展时禁用发送**”设置为“**已启用**”来启用 Onsend 策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-265">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="f9172-266">若为用户禁用策略，管理员应将其设置为“**已禁用**”。</span><span class="sxs-lookup"><span data-stu-id="f9172-266">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="f9172-267">若要管理此策略设置，可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="f9172-267">To manage this policy setting, you can do the following:</span></span>

1. <span data-ttu-id="f9172-268">下载最新的[管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)。</span><span class="sxs-lookup"><span data-stu-id="f9172-268">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="f9172-269">打开 **gpedit.msc (本地组策略**) 。</span><span class="sxs-lookup"><span data-stu-id="f9172-269">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="f9172-270">导航到 **“用户配置”>“管理模板”>“Microsoft Outlook 2016”>“安全性”>“信任中心”**。</span><span class="sxs-lookup"><span data-stu-id="f9172-270">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="f9172-271">选择“**无法加载 Web 扩展时禁用发送**”设置。</span><span class="sxs-lookup"><span data-stu-id="f9172-271">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="f9172-272">打开链接以编辑策略设置。</span><span class="sxs-lookup"><span data-stu-id="f9172-272">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="f9172-273">在“**无法加载 Web 扩展时禁用发送**”对话框窗口中，根据需要选择“**已启用**”或“**已禁用**”，然后选择“**确定**”或“**应用**”以使更新生效。</span><span class="sxs-lookup"><span data-stu-id="f9172-273">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="f9172-274">Mac</span><span class="sxs-lookup"><span data-stu-id="f9172-274">Mac</span></span>](#tab/unix)

<span data-ttu-id="f9172-275">对于安装了使用 Onsend 功能的 Mac 版 Outlook 加载项的任何用户，系统会为其运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-275">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="f9172-276">但是，如果用户需要运行该加载项来满足合规性标准，则必须在每个用户的计算机上应用以下邮箱设置。</span><span class="sxs-lookup"><span data-stu-id="f9172-276">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="f9172-277">此设置或键与 CFPreferences 兼容，这意味着可以使用适用于 Mac 的企业管理软件（例如，Jamf Pro）来对其进行设置。</span><span class="sxs-lookup"><span data-stu-id="f9172-277">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

||<span data-ttu-id="f9172-278">值</span><span class="sxs-lookup"><span data-stu-id="f9172-278">Value</span></span>|
|:---|:---|
|<span data-ttu-id="f9172-279">**域**</span><span class="sxs-lookup"><span data-stu-id="f9172-279">**Domain**</span></span>|<span data-ttu-id="f9172-280">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="f9172-280">com.microsoft.outlook</span></span>|
|<span data-ttu-id="f9172-281">**键**</span><span class="sxs-lookup"><span data-stu-id="f9172-281">**Key**</span></span>|<span data-ttu-id="f9172-282">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="f9172-282">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="f9172-283">**DataType**</span><span class="sxs-lookup"><span data-stu-id="f9172-283">**DataType**</span></span>|<span data-ttu-id="f9172-284">Boolean</span><span class="sxs-lookup"><span data-stu-id="f9172-284">Boolean</span></span>|
|<span data-ttu-id="f9172-285">**可能的值**</span><span class="sxs-lookup"><span data-stu-id="f9172-285">**Possible values**</span></span>|<span data-ttu-id="f9172-286">false（默认值）</span><span class="sxs-lookup"><span data-stu-id="f9172-286">false (default)</span></span><br><span data-ttu-id="f9172-287">true</span><span class="sxs-lookup"><span data-stu-id="f9172-287">true</span></span>|
|<span data-ttu-id="f9172-288">**可用性**</span><span class="sxs-lookup"><span data-stu-id="f9172-288">**Availability**</span></span>|<span data-ttu-id="f9172-289">16.27</span><span class="sxs-lookup"><span data-stu-id="f9172-289">16.27</span></span>|
|<span data-ttu-id="f9172-290">**备注**</span><span class="sxs-lookup"><span data-stu-id="f9172-290">**Comments**</span></span>|<span data-ttu-id="f9172-291">此键将创建 onSendMailbox 策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-291">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="f9172-292">设置的用途</span><span class="sxs-lookup"><span data-stu-id="f9172-292">What the setting does</span></span>

<span data-ttu-id="f9172-293">出于合规性原因，管理员可能需要在用户具有可供运行的最新 Onsend 加载项前，确保其无法发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="f9172-293">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="f9172-294">管理员必须启用键 **OnSendAddinsWaitForLoad**，以便所有加载项都从 Exchange 进行更新，并可用于在发送时验证每封邮件或每个会议项目是否符合预期的规则和规定。</span><span class="sxs-lookup"><span data-stu-id="f9172-294">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="f9172-295">键的状态</span><span class="sxs-lookup"><span data-stu-id="f9172-295">Key's state</span></span>|<span data-ttu-id="f9172-296">结果</span><span class="sxs-lookup"><span data-stu-id="f9172-296">Result</span></span>|
|---|---|
|<span data-ttu-id="f9172-297">false</span><span class="sxs-lookup"><span data-stu-id="f9172-297">false</span></span>|<span data-ttu-id="f9172-298">当前下载的 Ons ons ons 外接程序清单 (在发送的邮件或会议项目) 运行的最新版本。</span><span class="sxs-lookup"><span data-stu-id="f9172-298">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="f9172-299">这是默认状态/行为。</span><span class="sxs-lookup"><span data-stu-id="f9172-299">This is the default state/behavior.</span></span>|
|<span data-ttu-id="f9172-300">true</span><span class="sxs-lookup"><span data-stu-id="f9172-300">true</span></span>|<span data-ttu-id="f9172-301">从 Exchange 下载 Ons ons 外接程序的最新清单后，外接程序将运行在要发送的邮件或会议项目上。</span><span class="sxs-lookup"><span data-stu-id="f9172-301">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="f9172-302">否则，将阻止发送并禁用 **"** 发送"按钮。</span><span class="sxs-lookup"><span data-stu-id="f9172-302">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="f9172-303">Onsend 功能的应用场景</span><span class="sxs-lookup"><span data-stu-id="f9172-303">On-send feature scenarios</span></span>

<span data-ttu-id="f9172-304">以下是支持和不支持使用 Onsend 功能的加载项的应用场景。</span><span class="sxs-lookup"><span data-stu-id="f9172-304">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="f9172-305">用户邮箱启用了 Onsend 加载项功能，但未安装任何加载项</span><span class="sxs-lookup"><span data-stu-id="f9172-305">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="f9172-306">在这种场景中，用户将能够在不执行任何加载项的情况下发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="f9172-306">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="f9172-307">用户邮箱启用了 Onsend 加载项功能，并且安装并启用了支持 Onsend 的加载项</span><span class="sxs-lookup"><span data-stu-id="f9172-307">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="f9172-308">外接程序在发送事件期间运行，然后允许或阻止用户发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-308">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="f9172-309">邮箱委派，其中邮箱 1 具有对邮箱 2 的完全访问权限</span><span class="sxs-lookup"><span data-stu-id="f9172-309">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="f9172-310">Web 浏览器（经典 Outlook）</span><span class="sxs-lookup"><span data-stu-id="f9172-310">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="f9172-311">方案</span><span class="sxs-lookup"><span data-stu-id="f9172-311">Scenario</span></span>|<span data-ttu-id="f9172-312">邮箱 1 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="f9172-312">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="f9172-313">邮箱 2 Onsend 功能</span><span class="sxs-lookup"><span data-stu-id="f9172-313">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="f9172-314">Outlook Web 会话（经典）</span><span class="sxs-lookup"><span data-stu-id="f9172-314">Outlook web session (classic)</span></span>|<span data-ttu-id="f9172-315">结果</span><span class="sxs-lookup"><span data-stu-id="f9172-315">Result</span></span>|<span data-ttu-id="f9172-316">是否支持？</span><span class="sxs-lookup"><span data-stu-id="f9172-316">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="f9172-317">1</span><span class="sxs-lookup"><span data-stu-id="f9172-317">1</span></span>|<span data-ttu-id="f9172-318">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-318">Enabled</span></span>|<span data-ttu-id="f9172-319">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-319">Enabled</span></span>|<span data-ttu-id="f9172-320">新会话</span><span class="sxs-lookup"><span data-stu-id="f9172-320">New session</span></span>|<span data-ttu-id="f9172-321">邮箱 1 无法从邮箱 2 发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="f9172-321">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="f9172-p135">目前尚不支持。可以使用方案 3 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="f9172-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="f9172-324">2</span><span class="sxs-lookup"><span data-stu-id="f9172-324">2</span></span>|<span data-ttu-id="f9172-325">已禁用</span><span class="sxs-lookup"><span data-stu-id="f9172-325">Disabled</span></span>|<span data-ttu-id="f9172-326">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-326">Enabled</span></span>|<span data-ttu-id="f9172-327">新会话</span><span class="sxs-lookup"><span data-stu-id="f9172-327">New session</span></span>|<span data-ttu-id="f9172-328">邮箱 1 无法从邮箱 2 发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="f9172-328">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="f9172-p136">目前尚不支持。可以使用方案 3 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="f9172-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="f9172-331">3</span><span class="sxs-lookup"><span data-stu-id="f9172-331">3</span></span>|<span data-ttu-id="f9172-332">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-332">Enabled</span></span>|<span data-ttu-id="f9172-333">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-333">Enabled</span></span>|<span data-ttu-id="f9172-334">同一个会话</span><span class="sxs-lookup"><span data-stu-id="f9172-334">Same session</span></span>|<span data-ttu-id="f9172-335">分配给邮箱 1 的 Onsend 加载项运行 Onsend。</span><span class="sxs-lookup"><span data-stu-id="f9172-335">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="f9172-336">支持。</span><span class="sxs-lookup"><span data-stu-id="f9172-336">Supported.</span></span>|
|<span data-ttu-id="f9172-337">4 </span><span class="sxs-lookup"><span data-stu-id="f9172-337">4</span></span>|<span data-ttu-id="f9172-338">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-338">Enabled</span></span>|<span data-ttu-id="f9172-339">已禁用</span><span class="sxs-lookup"><span data-stu-id="f9172-339">Disabled</span></span>|<span data-ttu-id="f9172-340">新会话</span><span class="sxs-lookup"><span data-stu-id="f9172-340">New session</span></span>|<span data-ttu-id="f9172-341">未运行 Onsend 加载项；邮件或会议项目已发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-341">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="f9172-342">支持。</span><span class="sxs-lookup"><span data-stu-id="f9172-342">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="f9172-343">Web 浏览器（新式 Outlook）、Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="f9172-343">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="f9172-344">若要强制执行 Onsend，管理员应确保对两个邮箱都启用了该策略。</span><span class="sxs-lookup"><span data-stu-id="f9172-344">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="f9172-345">若要了解如何在加载项中支持委派访问，请参阅[在 Outlook 加载项中启用委派访问方案](delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="f9172-345">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="f9172-346">组 1 是新式组邮箱，用户邮箱 1 是组 1 的成员</span><span class="sxs-lookup"><span data-stu-id="f9172-346">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="f9172-347">方案</span><span class="sxs-lookup"><span data-stu-id="f9172-347">Scenario</span></span>|<span data-ttu-id="f9172-348">邮箱 1 Onsend 策略</span><span class="sxs-lookup"><span data-stu-id="f9172-348">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="f9172-349">是否启用了 Onsend 加载项？</span><span class="sxs-lookup"><span data-stu-id="f9172-349">On-send add-ins enabled?</span></span>|<span data-ttu-id="f9172-350">邮箱 1 操作</span><span class="sxs-lookup"><span data-stu-id="f9172-350">Mailbox 1 action</span></span>|<span data-ttu-id="f9172-351">结果</span><span class="sxs-lookup"><span data-stu-id="f9172-351">Result</span></span>|<span data-ttu-id="f9172-352">是否支持？</span><span class="sxs-lookup"><span data-stu-id="f9172-352">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="f9172-353">1</span><span class="sxs-lookup"><span data-stu-id="f9172-353">1</span></span>|<span data-ttu-id="f9172-354">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-354">Enabled</span></span>|<span data-ttu-id="f9172-355">是</span><span class="sxs-lookup"><span data-stu-id="f9172-355">Yes</span></span>|<span data-ttu-id="f9172-356">邮箱 1 撰写发送到组 1 的新邮件或会议。</span><span class="sxs-lookup"><span data-stu-id="f9172-356">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="f9172-357">发送期间，Onsend 加载项运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-357">On-send add-ins run during send.</span></span>|<span data-ttu-id="f9172-358">是</span><span class="sxs-lookup"><span data-stu-id="f9172-358">Yes</span></span>|
|<span data-ttu-id="f9172-359">2</span><span class="sxs-lookup"><span data-stu-id="f9172-359">2</span></span>|<span data-ttu-id="f9172-360">已启用</span><span class="sxs-lookup"><span data-stu-id="f9172-360">Enabled</span></span>|<span data-ttu-id="f9172-361">是</span><span class="sxs-lookup"><span data-stu-id="f9172-361">Yes</span></span>|<span data-ttu-id="f9172-362">邮箱 1 在 Outlook 网页版组 1 的组窗口中撰写发送到组 1 的新邮件或会议。</span><span class="sxs-lookup"><span data-stu-id="f9172-362">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="f9172-363">Onsend 加载项不会在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-363">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="f9172-364">目前尚不支持。</span><span class="sxs-lookup"><span data-stu-id="f9172-364">Not currently supported.</span></span> <span data-ttu-id="f9172-365">可以使用方案 1 作为一种解决办法。</span><span class="sxs-lookup"><span data-stu-id="f9172-365">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="f9172-366">用户邮箱启用了 Onsend 加载项功能/策略，并且安装并启用了支持 Onsend 的加载项，启用了脱机模式</span><span class="sxs-lookup"><span data-stu-id="f9172-366">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="f9172-367">Onsend 加载项将根据用户、加载项后端和 Exchange 的联机状态运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-367">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="f9172-368">用户的状态</span><span class="sxs-lookup"><span data-stu-id="f9172-368">User's state</span></span>

<span data-ttu-id="f9172-369">如果用户处于联机状态，则 Onsend 加载项将在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-369">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="f9172-370">如果用户处于脱机状态，Onsend 加载项不会在发送期间运行，也不会发送邮件或会议项目。</span><span class="sxs-lookup"><span data-stu-id="f9172-370">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="f9172-371">加载项后端的状态</span><span class="sxs-lookup"><span data-stu-id="f9172-371">Add-in backend's state</span></span>

<span data-ttu-id="f9172-372">如果 Onsend 加载项的后端处于联机状态且可访问，则将运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-372">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="f9172-373">如果后端处于脱机状态，则将禁用发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-373">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="f9172-374">Exchange 的状态</span><span class="sxs-lookup"><span data-stu-id="f9172-374">Exchange's state</span></span>

<span data-ttu-id="f9172-375">如果 Exchange 服务器处于联机状态且可访问，则 Onsend 加载项将在发送期间运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-375">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="f9172-376">如果 Onsend 加载项无法访问 Exchange 并且已启用适用的策略或 cmdlet，则将禁用发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-376">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="f9172-377">在处于任何脱机状态的 Mac 上，“**发送**”按钮（或现有会议的“**发送更新**”按钮）将被禁用，并显示当用户脱机时其组织不允许发送的通知。</span><span class="sxs-lookup"><span data-stu-id="f9172-377">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="f9172-378">用户可以在 Onss ons ons add-ins 处理项目时编辑项目</span><span class="sxs-lookup"><span data-stu-id="f9172-378">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="f9172-379">Ons ons ons an add-ins are processing an item， the user can edit the item by adding， for example， inappropriate text or attachments.</span><span class="sxs-lookup"><span data-stu-id="f9172-379">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="f9172-380">如果要阻止用户在加载项在发送时编辑项目，可以使用对话框实现解决方法。</span><span class="sxs-lookup"><span data-stu-id="f9172-380">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="f9172-381">此解决方法可在 Outlook 网页 (、Windows 和 Mac) 使用。</span><span class="sxs-lookup"><span data-stu-id="f9172-381">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9172-382">新式 Outlook 网页版：若要防止用户在加载项在发送时编辑项目，应设置 *OnSendAddinsEnabled* 标志，如本文前面安装使用 Onsend 的 Outlook 加载项部分所述。 `true` [](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send)</span><span class="sxs-lookup"><span data-stu-id="f9172-382">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="f9172-383">在 On-send 处理程序中：</span><span class="sxs-lookup"><span data-stu-id="f9172-383">In your on-send handler:</span></span>

1. <span data-ttu-id="f9172-384">调用 [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) 打开对话框，以便禁用鼠标单击和击键。</span><span class="sxs-lookup"><span data-stu-id="f9172-384">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="f9172-385">若要在经典 Outlook 网页中获取此行为，应在调用的参数中将 [displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) `true` `options` `displayDialogAsync` 属性设置为 。</span><span class="sxs-lookup"><span data-stu-id="f9172-385">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="f9172-386">实现对项目的处理。</span><span class="sxs-lookup"><span data-stu-id="f9172-386">Implement processing of the item.</span></span>
1. <span data-ttu-id="f9172-387">关闭该对话框。</span><span class="sxs-lookup"><span data-stu-id="f9172-387">Close the dialog.</span></span> <span data-ttu-id="f9172-388">此外，请处理用户关闭对话框时发生的情况。</span><span class="sxs-lookup"><span data-stu-id="f9172-388">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="f9172-389">代码示例</span><span class="sxs-lookup"><span data-stu-id="f9172-389">Code examples</span></span>

<span data-ttu-id="f9172-390">以下代码示例说明如何创建一个简单的 Onsend 加载项。</span><span class="sxs-lookup"><span data-stu-id="f9172-390">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="f9172-391">若要下载这些示例所基于的代码示例，请参阅 [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send)。</span><span class="sxs-lookup"><span data-stu-id="f9172-391">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="f9172-392">如果将对话框与 On-send 事件一同使用，请确保在完成该事件之前关闭该对话框。</span><span class="sxs-lookup"><span data-stu-id="f9172-392">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="f9172-393">清单、版本重写和事件</span><span class="sxs-lookup"><span data-stu-id="f9172-393">Manifest, version override, and event</span></span>

<span data-ttu-id="f9172-394">[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) 代码示例包括两个清单：</span><span class="sxs-lookup"><span data-stu-id="f9172-394">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="f9172-395">`Contoso Message Body Checker.xml` &ndash; 展示了如何在发送时检查邮件正文是否包含限制字词或敏感信息。</span><span class="sxs-lookup"><span data-stu-id="f9172-395">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="f9172-396">`Contoso Subject and CC Checker.xml` &ndash; 展示了如何将收件人添加到抄送行，并在发送时验证邮件是否包含主题行。</span><span class="sxs-lookup"><span data-stu-id="f9172-396">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="f9172-397">在 `Contoso Message Body Checker.xml` 清单文件中，将包含在 `ItemSend` 事件中应调用的函数文件和函数名称。</span><span class="sxs-lookup"><span data-stu-id="f9172-397">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="f9172-398">该操作将同步运行。</span><span class="sxs-lookup"><span data-stu-id="f9172-398">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="f9172-399">如果使用 Visual Studio 2019 开发 Onss ons ons 外接程序，则可能会收到如下验证警告："这是无效的 xsi：type ' http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events '。"若要处理此问题，你需要更高版本的 MailAppVersionOverridesV1_1.xsd，该版本在有关此警告的博客中已作为 GitHub gist [提供](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/)。</span><span class="sxs-lookup"><span data-stu-id="f9172-399">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="f9172-400">对于 `Contoso Subject and CC Checker.xml` 清单文件，以下示例中显示了邮件发送事件中要调用的函数文件和函数名称。</span><span class="sxs-lookup"><span data-stu-id="f9172-400">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="f9172-401">Onsend API 需要 `VersionOverrides v1_1`。</span><span class="sxs-lookup"><span data-stu-id="f9172-401">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="f9172-402">以下显示如何在清单中添加 `VersionOverrides` 节点。</span><span class="sxs-lookup"><span data-stu-id="f9172-402">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="f9172-403">有关详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="f9172-403">For more information, see the following:</span></span>
> - [<span data-ttu-id="f9172-404">Outlook 外接程序清单</span><span class="sxs-lookup"><span data-stu-id="f9172-404">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="f9172-405">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="f9172-405">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="f9172-406">`Event` 和 `item` 对象以及 `body.getAsync` 和 `body.setAsync` 方法</span><span class="sxs-lookup"><span data-stu-id="f9172-406">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="f9172-407">若要访问当前选择的邮件或会议项目（在本示例中为新撰写的邮件），请使用 `Office.context.mailbox.item` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="f9172-407">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="f9172-408">`ItemSend` 事件由 Onsend 功能自动传递到清单中指定的函数&mdash;在本示例中为 `validateBody` 函数。</span><span class="sxs-lookup"><span data-stu-id="f9172-408">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="f9172-409">`validateBody` 函数以指定格式 (HTML) 获取当前正文，并在回调方法中传递代码想要访问的 `ItemSend` 事件对象。</span><span class="sxs-lookup"><span data-stu-id="f9172-409">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="f9172-410">除 `getAsync` 方法之外，`Body` 对象还提供了 `setAsync` 方法，可用于将正文替换为指定的文本。</span><span class="sxs-lookup"><span data-stu-id="f9172-410">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="f9172-411">有关详细信息，请参阅 [Event 对象](/javascript/api/office/office.addincommands.event)和 [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="f9172-411">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="f9172-412">`NotificationMessages` 对象和 `event.completed` 方法</span><span class="sxs-lookup"><span data-stu-id="f9172-412">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="f9172-413">`checkBodyOnlyOnSendCallBack` 函数使用正则表达式来确定邮件正文是否包含禁止使用的词语。</span><span class="sxs-lookup"><span data-stu-id="f9172-413">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="f9172-414">如果该函数发现受限词语数组的匹配项，则将阻止发送电子邮件，并通过信息栏通知发件人。</span><span class="sxs-lookup"><span data-stu-id="f9172-414">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="f9172-415">为了做到这一点，它使用 `Item` 对象的 `notificationMessages` 属性来返回 `NotificationMessages` 对象。</span><span class="sxs-lookup"><span data-stu-id="f9172-415">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="f9172-416">然后，通过调用 `addAsync` 方法向该项目添加通知，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="f9172-416">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="f9172-417">以下是 `addAsync` 方法的参数：</span><span class="sxs-lookup"><span data-stu-id="f9172-417">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="f9172-418">`NoSend` &ndash; 一个字符串，即开发人员指定用于引用通知邮件的密钥。</span><span class="sxs-lookup"><span data-stu-id="f9172-418">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="f9172-419">可用于在以后修改此邮件。</span><span class="sxs-lookup"><span data-stu-id="f9172-419">You can use it to modify this message later.</span></span> <span data-ttu-id="f9172-420">该键不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="f9172-420">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="f9172-421">`type` &ndash; JSON 对象参数的一个属性。</span><span class="sxs-lookup"><span data-stu-id="f9172-421">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="f9172-422">表示邮件的类型；类型对应于 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) 枚举的值。</span><span class="sxs-lookup"><span data-stu-id="f9172-422">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="f9172-423">可能的值是进度指示器、信息消息或错误消息。</span><span class="sxs-lookup"><span data-stu-id="f9172-423">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="f9172-424">在此示例中，`type` 是错误消息。</span><span class="sxs-lookup"><span data-stu-id="f9172-424">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="f9172-425">`message` &ndash; JSON 对象参数的一个属性。</span><span class="sxs-lookup"><span data-stu-id="f9172-425">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="f9172-426">在此示例中，`message` 是通知邮件的文本。</span><span class="sxs-lookup"><span data-stu-id="f9172-426">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="f9172-427">为表明加载项对由发送操作触发的 `ItemSend` 事件的处理已完成，请调用 `event.completed({allowEvent:Boolean})` 方法。</span><span class="sxs-lookup"><span data-stu-id="f9172-427">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="f9172-428">`allowEvent` 属性是一个布尔值。</span><span class="sxs-lookup"><span data-stu-id="f9172-428">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="f9172-429">如果设置为 `true`，则允许发送。</span><span class="sxs-lookup"><span data-stu-id="f9172-429">If set to `true`, send is allowed.</span></span> <span data-ttu-id="f9172-430">如果设置为 `false`，则将阻止发送电子邮件。</span><span class="sxs-lookup"><span data-stu-id="f9172-430">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="f9172-431">有关详细信息，请参阅 [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [completed](/javascript/api/office/office.addincommands.event)。</span><span class="sxs-lookup"><span data-stu-id="f9172-431">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="f9172-432">`replaceAsync`、`removeAsync` 和 `getAllAsync` 方法</span><span class="sxs-lookup"><span data-stu-id="f9172-432">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="f9172-433">除了 `addAsync` 方法之外，`NotificationMessages` 对象还包括 `replaceAsync`、`removeAsync` 和 `getAllAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="f9172-433">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="f9172-434">此代码示例中不使用这些方法。</span><span class="sxs-lookup"><span data-stu-id="f9172-434">These methods are not used in this code sample.</span></span>  <span data-ttu-id="f9172-435">有关详细信息，请参阅 [NotificationMessages](/javascript/api/outlook/office.NotificationMessages)。</span><span class="sxs-lookup"><span data-stu-id="f9172-435">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="f9172-436">主题和抄送检查器代码</span><span class="sxs-lookup"><span data-stu-id="f9172-436">Subject and CC checker code</span></span>

<span data-ttu-id="f9172-437">以下代码示例介绍如何将收件人添加到抄送行，并验证邮件在发送时是否包含主题。</span><span class="sxs-lookup"><span data-stu-id="f9172-437">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="f9172-438">此示例使用 Onsend 功能允许或禁止发送电子邮件。</span><span class="sxs-lookup"><span data-stu-id="f9172-438">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="f9172-p156">若要详细了解如何将收件人添加到抄送行、验证电子邮件在发送时是否包主题行，以及查看可以使用的 API，请参阅 [Outlook-Add-in-On-Send 示例](https://github.com/OfficeDev/Outlook-Add-in-On-Send)。已充分注释代码。</span><span class="sxs-lookup"><span data-stu-id="f9172-p156">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="f9172-441">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f9172-441">See also</span></span>

- [<span data-ttu-id="f9172-442">Outlook 加载项体系结构和功能概述</span><span class="sxs-lookup"><span data-stu-id="f9172-442">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="f9172-443">加载项命令演示 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="f9172-443">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)