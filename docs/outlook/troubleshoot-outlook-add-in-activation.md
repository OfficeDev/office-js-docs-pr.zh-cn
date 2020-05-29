---
title: Outlook 上下文加载项激活故障排查
description: 如果加载项未按预期激活，应考虑以下几个方面的可能原因。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 555ae2a45bf49d74d1fd439258fd87035644e86a
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413180"
---
# <a name="troubleshoot-outlook-add-in-activation"></a><span data-ttu-id="40f35-103">Outlook 加载项激活故障排查</span><span class="sxs-lookup"><span data-stu-id="40f35-103">Troubleshoot Outlook add-in activation</span></span>

<span data-ttu-id="40f35-p101">Outlook 上下文加载项激活基于加载项清单中的激活规则。在当前选定项的条件满足加载项的激活规则时，主机应用程序激活，并在 Outlook UI 中显示加载项按钮（用于撰写加载项的加载项选择窗格，用于阅读加载项的加载项条）。但是，如果加载项未按预期激活，应考虑以下几个方面的原因。</span><span class="sxs-lookup"><span data-stu-id="40f35-p101">Outlook contextual add-in activation is based on the activation rules in the add-in manifest. When conditions for the currently selected item satisfy the activation rules for the add-in, the host application activates and displays the add-in button in the Outlook UI (add-in selection pane for compose add-ins, add-in bar for read add-ins). However, if your add-in doesn't activate as you expect, you should look into the following areas for possible reasons.</span></span>

## <a name="is-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a><span data-ttu-id="40f35-107">用户邮箱是否位于不低于 Exchange 2013 版本的 Exchange Server 上？</span><span class="sxs-lookup"><span data-stu-id="40f35-107">Is user mailbox on a version of Exchange Server that is at least Exchange 2013?</span></span>

<span data-ttu-id="40f35-p102">首先，确保你正在测试的用户电子邮件帐户位于至少为 Exchange 2013 的某个版本的 Exchange Server 上。如果你正在使用在Exchange 2013 之后发布的特定功能，那么请确保用户的帐户使用合适的 Exchange 版本。</span><span class="sxs-lookup"><span data-stu-id="40f35-p102">First, ensure that the user's email account you're testing with is on a version of Exchange Server that is at least Exchange 2013. If you are using specific features that are released after Exchange 2013, make sure the user's account is on the appropriate version of Exchange.</span></span>

<span data-ttu-id="40f35-110">你可使用以下方法之一验证 Exchange 2013 的版本：</span><span class="sxs-lookup"><span data-stu-id="40f35-110">You can verify the version of Exchange 2013 by using one of the following approaches:</span></span>

- <span data-ttu-id="40f35-111">咨询你的 Exchange Server 管理员。</span><span class="sxs-lookup"><span data-stu-id="40f35-111">Check with your Exchange Server administrator.</span></span>

- <span data-ttu-id="40f35-p103">若要在 Outlook 网页版或移动设备版上测试加载项，请在脚本调试器（例如，Internet Explorer 随附的 JScript 调试器）中，查找指定脚本加载位置的 **script** 标记的 **src** 属性。路径应包含子字符串 **owa/15.0.516.x/owa2/...**，其中 **15.0.516.x** 表示 Exchange Server 版本（如 **15.0.516.2**）。</span><span class="sxs-lookup"><span data-stu-id="40f35-p103">If you are testing the add-in on Outlook on the web or mobile devices, in a script debugger (for example, the JScript Debugger that comes with Internet Explorer), look for the **src** attribute of the **script** tag that specifies the location from which scripts are loaded. The path should contain a substring **owa/15.0.516.x/owa2/...**, where **15.0.516.x** represents the version of the Exchange Server, such as **15.0.516.2**.</span></span>

- <span data-ttu-id="40f35-p104">也可以使用 [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) 属性来验证版本。在 Outlook 网页版和移动设备版上，此属性会返回 Exchange Server 版本。</span><span class="sxs-lookup"><span data-stu-id="40f35-p104">Alternatively, you can use the [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) property to verify the version. On Outlook on the web and mobile devices, this property returns the version of the Exchange Server.</span></span>

- <span data-ttu-id="40f35-116">如果能够在 Outlook 上测试加载项，则可使用采用 Outlook 对象模型和 Visual Basic 编辑器的以下简单调试技术：</span><span class="sxs-lookup"><span data-stu-id="40f35-116">If you can test the add-in on Outlook, you can use the following simple debugging technique that uses the Outlook object model and Visual Basic Editor:</span></span>

    1. <span data-ttu-id="40f35-p105">首先，确认已对 Outlook 启用了宏。依次选择“**文件**”、“**选项**”、“**信任中心**”、“**信任中心设置**”、“**宏设置**”。确保在“信任中心”中选择了“**为所有宏提供通知**”。还应确保在 Outlook 启动过程中选择了“**启用宏**”。</span><span class="sxs-lookup"><span data-stu-id="40f35-p105">First, verify that macros are enabled for Outlook. Choose **File**, **Options**, **Trust Center**, **Trust Center Settings**, **Macro Settings**. Ensure that **Notifications for all macros** is selected in the Trust Center. You should have also selected **Enable Macros** during Outlook startup.</span></span>

    1. <span data-ttu-id="40f35-121">在功能区的“**开发人员**”选项卡上，选择“**Visual Basic**”。</span><span class="sxs-lookup"><span data-stu-id="40f35-121">On the **Developer** tab of the ribbon, choose **Visual Basic**.</span></span>

       > [!NOTE]
       > <span data-ttu-id="40f35-p106">看不到“**开发人员**”选项卡？请参阅[如何：在功能区上显示“开发人员”选项卡](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon)，启用此选项卡。</span><span class="sxs-lookup"><span data-stu-id="40f35-p106">Not seeing the **Developer** tab? See [How to: Show the Developer Tab on the Ribbon](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon) to turn it on.</span></span>

    1. <span data-ttu-id="40f35-124">在 Visual Basic 编辑器中，依次选择“**视图**”和“**即时窗口**”。</span><span class="sxs-lookup"><span data-stu-id="40f35-124">In the Visual Basic Editor, choose **View**, **Immediate Window**.</span></span>

    1. <span data-ttu-id="40f35-p107">在即时窗口中键入以下内容以显示 Exchange Server 的版本。返回值的主版本必须等于或大于 15。</span><span class="sxs-lookup"><span data-stu-id="40f35-p107">Type the following in the Immediate window to display the version of the Exchange Server. The major version of the returned value must be equal to or greater than 15.</span></span>

       - <span data-ttu-id="40f35-127">如果用户的配置文件中只有一个 Exchange 帐户：</span><span class="sxs-lookup"><span data-stu-id="40f35-127">If there is only one Exchange account in the user's profile:</span></span>

       ```vb
        ?Session.ExchangeMailboxServerVersion
       ```

       - <span data-ttu-id="40f35-128">如果同一用户配置文件中有多个 Exchange 帐户（`emailAddress` 表示包含用户主 SMTP 地址的字符串）：</span><span class="sxs-lookup"><span data-stu-id="40f35-128">If there are multiple Exchange accounts in the same user profile (`emailAddress` represents a string that contains the user's primary SMTP address):</span></span>

       ```vb
        ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
       ```

## <a name="is-the-add-in-disabled"></a><span data-ttu-id="40f35-129">加载项否已禁用？</span><span class="sxs-lookup"><span data-stu-id="40f35-129">Is the add-in disabled?</span></span>

<span data-ttu-id="40f35-p108">任何 Outlook 富客户端可出于性能原因禁用加载项，这些原因包括超出 CPU 内核或内存的使用阈值、超出崩溃容忍度以及超出处理加载项的所有正则表达式的时间。发生这种情况时，Outlook 富客户端会显示一条禁用加载项的通知。</span><span class="sxs-lookup"><span data-stu-id="40f35-p108">Any one of the Outlook rich clients can disable an add-in for performance reasons, including exceeding usage thresholds for CPU core or memory, tolerance for crashes, and length of time to process all the regular expressions for an add-in. When this happens, the Outlook rich client displays a notification that it is disabling the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="40f35-132">仅 Outlook 富客户端可监视资源使用状况，但如果在 Outlook 富客户端中禁用加载项，也会在 Outlook 网页版和移动设备版中禁用此加载项。</span><span class="sxs-lookup"><span data-stu-id="40f35-132">Only Outlook rich clients monitor resource usage, but disabling an add-in in an Outlook rich client also disables the add-in in Outlook on the web and mobile devices.</span></span>

<span data-ttu-id="40f35-133">使用以下方法之一，验证加载项是否已禁用：</span><span class="sxs-lookup"><span data-stu-id="40f35-133">Use one of the following approaches to verify whether an add-in is disabled:</span></span>

- <span data-ttu-id="40f35-134">在 Outlook 网页版中，直接登录电子邮件帐户，选择“设置”图标，然后选择“**管理加载项**”转到 Exchange 管理中心，可在此处验证管理加载项是否已启用。</span><span class="sxs-lookup"><span data-stu-id="40f35-134">In Outlook on the web, sign in directly to the email account, choose the Settings icon, and then choose **Manage add-ins** to go to the Exchange Admin Center, where you can verify whether the add-in is enabled.</span></span>

- <span data-ttu-id="40f35-135">在 Windows 版 Outlook 中，转到 Backstage 视图并选择“**管理加载项**”。登录 Exchange 管理中心验证加载项是否已启用。</span><span class="sxs-lookup"><span data-stu-id="40f35-135">In Outlook on Windows, go to the Backstage view and choose **Manage add-ins**. Sign in to the Exchange Admin Center to verify whether the add-in is enabled.</span></span>

- <span data-ttu-id="40f35-p109">在 Mac 版 Outlook 中，选择加载项栏中的“**管理加载项**”。登录 Exchange 管理中心验证加载项是否已启用。</span><span class="sxs-lookup"><span data-stu-id="40f35-p109">In Outlook on Mac, choose **Manage add-ins** in the add-in bar. Sign in to the Exchange Admin Center to verify whether the add-in is enabled.</span></span>

## <a name="does-the-tested-item-support-outlook-add-ins-is-the-selected-item-delivered-by-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a><span data-ttu-id="40f35-p110">已测试项是否支持 Outlook 加载项？所选项目是否由至少为 Exchange 2013 的某个版本的 Exchange Server 提供？</span><span class="sxs-lookup"><span data-stu-id="40f35-p110">Does the tested item support Outlook add-ins? Is the selected item delivered by a version of Exchange Server that is at least Exchange 2013?</span></span>

<span data-ttu-id="40f35-140">如果你的 Outlook 加载项为阅读加载项，并且应该在用户查看消息（包括电子邮件、会议请求、响应和取消）或约会时激活，尽管这些项目通常支持加载项，但还是存在例外情况。</span><span class="sxs-lookup"><span data-stu-id="40f35-140">If your Outlook add-in is a read add-in and is supposed to be activated when the user is viewing a message (including email messages, meeting requests, responses, and cancellations) or appointment, even though these items generally support add-ins, there are exceptions.</span></span> <span data-ttu-id="40f35-141">检查所选的项目是否是 [Outlook 加载项未激活列表](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)中的一项。</span><span class="sxs-lookup"><span data-stu-id="40f35-141">Check if the selected item is one of those [listed where Outlook add-ins do not activate](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins).</span></span>

<span data-ttu-id="40f35-142">此外，由于约会始终以 RTF 格式保存，因此指定 [BodyAsHTML](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 的 **PropertyName** 值的 **ItemHasRegularExpressionMatch** 规则不会对以纯文本或 RTF 格式保存的约会或邮件激活加载项。</span><span class="sxs-lookup"><span data-stu-id="40f35-142">Also, because appointments are always saved in Rich Text Format, an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule that specifies a **PropertyName** value of **BodyAsHTML** would not activate an add-in on an appointment or message that is saved in plain text or Rich Text Format.</span></span>

<span data-ttu-id="40f35-p112">即使某邮件项不是以上类型之一，如果该项不是使用至少为 Exchange 2013 的某个版本的 Exchange Server 传递，则不会在该项上确定已知实体和属性（如发件人的 SMTP 地址）。依赖这些实体或属性的任何激活规则将不会得到满足，并且加载项将不会激活。</span><span class="sxs-lookup"><span data-stu-id="40f35-p112">Even if a mail item is not one of the above types, if the item was not delivered by a version of Exchange Server that is at least Exchange 2013, known entities and properties such as sender's SMTP address would not be identified on the item. Any activation rules that rely on these entities or properties would not be satisfied, and the add-in would not be activated.</span></span>

<span data-ttu-id="40f35-145">如果您的加载项为撰写加载项并且应该在用户撰写邮件或会议请求时激活，请确保该项目未受 IRM 保护。</span><span class="sxs-lookup"><span data-stu-id="40f35-145">If your add-in is a compose add-in and is supposed to be activated when the user is authoring a message or meeting request, make sure the item is not protected by IRM.</span></span>

## <a name="is-the-add-in-manifest-installed-properly-and-does-outlook-have-a-cached-copy"></a><span data-ttu-id="40f35-146">加载项清单是否安装正确，Outlook 是否有已缓存副本？</span><span class="sxs-lookup"><span data-stu-id="40f35-146">Is the add-in manifest installed properly, and does Outlook have a cached copy?</span></span>

<span data-ttu-id="40f35-p113">此方案仅适用于 Windows 版 Outlook。正常情况下，为邮箱安装 Outlook 加载项时，Exchange Server 会将加载项清单从你指示的位置复制到该 Exchange Server 上的邮箱。每次启动 Outlook 时，它都会将为该邮箱安装的所有清单读取到以下位置的临时缓存中：</span><span class="sxs-lookup"><span data-stu-id="40f35-p113">This scenario applies to only Outlook on Windows. Normally, when you install an Outlook add-in for a mailbox, the Exchange Server copies the add-in manifest from the location you indicate to the mailbox on that Exchange Server. Every time Outlook starts, it reads all the manifests installed for that mailbox into a temporary cache at the following location:</span></span>

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

<span data-ttu-id="40f35-150">例如，对于用户 John，缓存可能位于 C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF。</span><span class="sxs-lookup"><span data-stu-id="40f35-150">For example, for the user John, the cache might be at C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="40f35-151">对于 Windows 上的 Outlook 2013，请使用15.0 而不是16.0，以便位置为：</span><span class="sxs-lookup"><span data-stu-id="40f35-151">For Outlook 2013 on Windows, use 15.0 instead of 16.0 so the location would be:</span></span>
>
> ```text
> %LocalAppData%\Microsoft\Office\15.0\WEF
> ```

<span data-ttu-id="40f35-p114">如果无法对任何项目激活加载项，则清单可能未正确安装在 Exchange Server 上，或者 Outlook 未在启动时正确读取清单。使用 Exchange 管理中心确保已为您的邮箱安装和启用加载项，并在必要时重新启动 Exchange Server。</span><span class="sxs-lookup"><span data-stu-id="40f35-p114">If an add-in does not activate for any items, the manifest might not have been installed properly on the Exchange Server, or Outlook has not read the manifest properly on startup. Using the Exchange Admin Center, ensure that the add-in is installed and enabled for your mailbox, and reboot the Exchange Server, if necessary.</span></span>

<span data-ttu-id="40f35-154">图 1 显示验证 Outlook 是否具有有效版本的清单的步骤摘要。</span><span class="sxs-lookup"><span data-stu-id="40f35-154">Figure 1 shows a summary of the steps to verify whether Outlook has a valid version of the manifest.</span></span>

<span data-ttu-id="40f35-155">**图 1.验证 Outlook 是否已正确缓存清单的步骤的流程图**</span><span class="sxs-lookup"><span data-stu-id="40f35-155">**Figure 1. Flow chart of the steps to verify whether Outlook properly cached the manifest**</span></span>

![用于检查清单的流程图](../images/troubleshoot-manifest-flow.png)

<span data-ttu-id="40f35-157">以下过程描述详细信息。</span><span class="sxs-lookup"><span data-stu-id="40f35-157">The following procedure describes the details.</span></span>

1. <span data-ttu-id="40f35-158">如果你已在 Outlook 打开时修改了清单，并且未使用 Visual Studio 2012 或 Visual Studio 的更高版本开发加载项，则应卸载加载项，并使用 Exchange 管理中心重新安装它。</span><span class="sxs-lookup"><span data-stu-id="40f35-158">If you have modified the manifest while Outlook is open, and you are not using Visual Studio 2012 or a later version of Visual Studio to develop the add-in, you should uninstall the add-in and reinstall it using the Exchange Admin Center.</span></span>

1. <span data-ttu-id="40f35-159">重新启动 Outlook 并测试 Outlook 现在是否已激活加载项。</span><span class="sxs-lookup"><span data-stu-id="40f35-159">Restart Outlook and test whether Outlook now activates the add-in.</span></span>

1. <span data-ttu-id="40f35-p115">如果 Outlook 无法激活加载项，则检查 Outlook 是否具有加载项清单的正确缓存副本。请查看以下路径：</span><span class="sxs-lookup"><span data-stu-id="40f35-p115">If Outlook doesn't activate the add-in, check whether Outlook has a properly cached copy of the manifest for the add-in. Look under the following path:</span></span>

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    <span data-ttu-id="40f35-162">可以在下列子文件夹中找到清单：</span><span class="sxs-lookup"><span data-stu-id="40f35-162">You can find the manifest in the following subfolder:</span></span>

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > <span data-ttu-id="40f35-163">下面的示例展示了为用户 John 的邮箱安装的清单路径：</span><span class="sxs-lookup"><span data-stu-id="40f35-163">The following is an example of a path to a manifest installed for a mailbox for the user John:</span></span>
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    <span data-ttu-id="40f35-164">验证要测试的加载项的清单是否在已缓存清单中。</span><span class="sxs-lookup"><span data-stu-id="40f35-164">Verify whether the manifest of the add-in you're testing is among the cached manifests.</span></span>

1. <span data-ttu-id="40f35-165">如果清单在缓存中，请跳过本节的其余部分，并考虑本节后面的其他可能原因。</span><span class="sxs-lookup"><span data-stu-id="40f35-165">If the manifest is in the cache, skip the rest of this section and consider the other possible reasons following this section.</span></span>

1. <span data-ttu-id="40f35-p116">如果清单不在缓存中，请检查 Outlook 是否已确实从 Exchange Server 中成功读取清单。为此，请使用 Windows 事件查看器：</span><span class="sxs-lookup"><span data-stu-id="40f35-p116">If the manifest is not in the cache, check whether Outlook indeed successfully read the manifest from the Exchange Server. To do that, use the Windows Event Viewer:</span></span>

    1. <span data-ttu-id="40f35-168">在“**Windows 日志**”下，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="40f35-168">Under **Windows Logs**, choose **Application**.</span></span>

    1. <span data-ttu-id="40f35-169">查找其事件 ID 等于 63（表示 Outlook 从 Exchange Server 下载清单）的近期事件。</span><span class="sxs-lookup"><span data-stu-id="40f35-169">Look for a reasonably recent event for which the Event ID equals 63, which represents Outlook downloading a manifest from an Exchange Server.</span></span>

    1. <span data-ttu-id="40f35-170">如果 Outlook 成功读取了清单，则记录的事件应包含以下说明：</span><span class="sxs-lookup"><span data-stu-id="40f35-170">If Outlook successfully read a manifest, the logged event should have the following description:</span></span>

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        <span data-ttu-id="40f35-171">然后，跳过本节的其余部分，并考虑本节后面的其他可能原因。</span><span class="sxs-lookup"><span data-stu-id="40f35-171">Then skip the rest of this section and consider the other possible reasons following this section.</span></span>

1. <span data-ttu-id="40f35-172">如果看不到成功事件，请关闭 Outlook，再删除以下路径中的所有清单：</span><span class="sxs-lookup"><span data-stu-id="40f35-172">If you don't see a successful event, close Outlook, and delete all the manifests in the following path:</span></span>

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    <span data-ttu-id="40f35-173">启动 Outlook，并测试 Outlook 现在是否已激活加载项。</span><span class="sxs-lookup"><span data-stu-id="40f35-173">Start Outlook and test whether Outlook now activates the add-in.</span></span>

1. <span data-ttu-id="40f35-174">如果 Outlook 无法激活加载项，请返回到第 3 步，再次确认 Outlook 是否已正确读取清单。</span><span class="sxs-lookup"><span data-stu-id="40f35-174">If Outlook doesn't activate the add-in, go back to Step 3 to verify again whether Outlook has properly read the manifest.</span></span>

## <a name="is-the-add-in-manifest-valid"></a><span data-ttu-id="40f35-175">加载项清单有效吗？</span><span class="sxs-lookup"><span data-stu-id="40f35-175">Is the add-in manifest valid?</span></span>

<span data-ttu-id="40f35-176">请参阅[验证并排查清单问题](../testing/troubleshoot-manifest.md)来调试加载项清单问题。</span><span class="sxs-lookup"><span data-stu-id="40f35-176">See [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>

## <a name="are-you-using-the-appropriate-activation-rules"></a><span data-ttu-id="40f35-177">使用的激活规则是否合适？</span><span class="sxs-lookup"><span data-stu-id="40f35-177">Are you using the appropriate activation rules?</span></span>

<span data-ttu-id="40f35-p117">自 Office 加载项清单架构的版本 1.1 起，你可以创建当用户位于撰写窗体（撰写加载项）或阅读窗体（阅读加载项）中时激活的加载项。确保为加载项将在其中激活的每种窗体类型指定相应的激活规则。例如，你可以仅使用 [ItemIs](../reference/manifest/rule.md#itemis-rule) 规则（**FormType** 属性设置为 **Edit** 或 **ReadOrEdit**）激活撰写加载项，你无法使用任何其他类型的规则，例如用于撰写加载项的 [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 和 [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 规则。有关详细信息，请参阅 [Outlook 加载项的激活规则](activation-rules.md)。</span><span class="sxs-lookup"><span data-stu-id="40f35-p117">Starting in version 1.1 of the Office Add-ins manifests schema, you can create add-ins that are activated when the user is in a compose form (compose add-ins) or in a read form (read add-ins). Make sure you specify the appropriate activation rules for each type of form that your add-in is supposed to activate in. For example, you can activate compose add-ins using only [ItemIs](../reference/manifest/rule.md#itemis-rule) rules with the **FormType** attribute set to **Edit** or **ReadOrEdit**, and you cannot use any of the other types of rules, such as [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) and [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules for compose add-ins. For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

## <a name="if-you-use-a-regular-expression-is-it-properly-specified"></a><span data-ttu-id="40f35-181">如果使用正则表达式，该表达式的指定是否正确？</span><span class="sxs-lookup"><span data-stu-id="40f35-181">If you use a regular expression, is it properly specified?</span></span>

<span data-ttu-id="40f35-p118">由于激活规则中的正则表达式是阅读加载项的 XML 清单文件的一部分，因此当正则表达式使用特定字符时，请务必遵守 XML 处理器支持的相应转义序列。表 1 列出了这些特殊字符。</span><span class="sxs-lookup"><span data-stu-id="40f35-p118">Because regular expressions in activation rules are part of the XML manifest file for a read add-in, if a regular expression uses certain characters, be sure to follow the corresponding escape sequence that XML processors support. Table 1 lists these special characters.</span></span>

<span data-ttu-id="40f35-184">**表 1.正则表达式的转义序列**</span><span class="sxs-lookup"><span data-stu-id="40f35-184">**Table 1. Escape sequences for regular expressions**</span></span>

|<span data-ttu-id="40f35-185">**字符**</span><span class="sxs-lookup"><span data-stu-id="40f35-185">**Character**</span></span>|<span data-ttu-id="40f35-186">**说明**</span><span class="sxs-lookup"><span data-stu-id="40f35-186">**Description**</span></span>|<span data-ttu-id="40f35-187">**要使用的转义序列**</span><span class="sxs-lookup"><span data-stu-id="40f35-187">**Escape sequence to use**</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="40f35-188">双引号</span><span class="sxs-lookup"><span data-stu-id="40f35-188">Double quotation mark</span></span>|<span data-ttu-id="40f35-189">&amp;quot;</span><span class="sxs-lookup"><span data-stu-id="40f35-189">&amp;quot;</span></span>|
|`&`|<span data-ttu-id="40f35-190">与号</span><span class="sxs-lookup"><span data-stu-id="40f35-190">Ampersand</span></span>|<span data-ttu-id="40f35-191">&amp;amp;</span><span class="sxs-lookup"><span data-stu-id="40f35-191">&amp;amp;</span></span>|
|`'`|<span data-ttu-id="40f35-192">撇号</span><span class="sxs-lookup"><span data-stu-id="40f35-192">Apostrophe</span></span>|<span data-ttu-id="40f35-193">&amp;apos;</span><span class="sxs-lookup"><span data-stu-id="40f35-193">&amp;apos;</span></span>|
|`<`|<span data-ttu-id="40f35-194">小于号</span><span class="sxs-lookup"><span data-stu-id="40f35-194">Less-than sign</span></span>|<span data-ttu-id="40f35-195">&amp;lt;</span><span class="sxs-lookup"><span data-stu-id="40f35-195">&amp;lt;</span></span>|
|`>`|<span data-ttu-id="40f35-196">大于号</span><span class="sxs-lookup"><span data-stu-id="40f35-196">Greater-than sign</span></span>|<span data-ttu-id="40f35-197">&amp;gt;</span><span class="sxs-lookup"><span data-stu-id="40f35-197">&amp;gt;</span></span>|

## <a name="if-you-use-a-regular-expression-is-the-read-add-in-activating-in-outlook-on-the-web-or-mobile-devices-but-not-in-any-of-the-outlook-rich-clients"></a><span data-ttu-id="40f35-198">如果使用正则表达式，阅读加载项是否在 Outlook 网页版或移动设备版（而不是个别 Outlook 富客户端）中激活？</span><span class="sxs-lookup"><span data-stu-id="40f35-198">If you use a regular expression, is the read add-in activating in Outlook on the web or mobile devices, but not in any of the Outlook rich clients?</span></span>

<span data-ttu-id="40f35-p119">Outlook 富客户端使用的正则表达式引擎与 Outlook 网页版和移动设备版使用的正则表达式引擎不同。Outlook 富客户端使用作为 Visual Studio 标准模板库的一部分提供的 C++ 正则表达式引擎。此引擎符合 ECMAScript 5 标准。Outlook 网页版和移动设备版使用属于 JavaScript 一部分的正则表达式评估，由浏览器提供，且支持 ECMAScript 5 超集。</span><span class="sxs-lookup"><span data-stu-id="40f35-p119">Outlook rich clients use a regular expression engine that's different from the one used by Outlook on the web and mobile devices. Outlook rich clients use the C++ regular expression engine provided as part of the Visual Studio standard template library. This engine complies with ECMAScript 5 standards. Outlook on the web and mobile devices use regular expression evaluation that is part of JavaScript, is provided by the browser, and supports a superset of ECMAScript 5.</span></span>

<span data-ttu-id="40f35-p120">在大多数情况下，这些主机应用程序在激活规则中为相同的正则表达式找到相同的匹配项，但也有例外。例如，如果正则表达式包含基于预定义字符类的自定义字符类，则 Outlook 富客户端可能会返回与 Outlook 网页版和移动设备版不同的结果。作为示例，在其中包含速记字符类 `[\d\w]` 的字符类将返回不同的结果。在这种情况下，为避免不同主机上出现不同结果，请改用 `(\d|\w)`。</span><span class="sxs-lookup"><span data-stu-id="40f35-p120">While in most cases, these host applications find the same matches for the same regular expression in an activation rule, there are exceptions. For instance, if the regex includes a custom character class based on predefined character classes, an Outlook rich client may return results different from Outlook on the web and mobile devices. As an example, character classes that contain shorthand character classes  `[\d\w]` within them would return different results. In this case, to avoid different results on different hosts, use `(\d|\w)` instead.</span></span>

<span data-ttu-id="40f35-p121">全面测试正则表达式。如果返回不同的结果，请重写正则表达式以兼容两个引擎。要验证 Outlook 富客户端上的评估结果，请编写一个小型 C++ 程序，该程序可将正则表达式应用于你尝试匹配的文本示例。在 Visual Studio 上运行时，C++ 测试程序将使用标准模板库，在运行相同正则表达式时模拟 Outlook 富客户端的行为。要验证 Outlook 网页版或移动设备版上的评估结果，请使用你喜爱的 JavaScript 正则表达式测试程序。</span><span class="sxs-lookup"><span data-stu-id="40f35-p121">Test your regular expression thoroughly. If it returns different results, rewrite the regular expression for compatibility with both engines. To verify evaluation results on an Outlook rich client, write a small C++ program that applies the regular expression against a sample of the text you are trying to match. Running on Visual Studio, the C++ test program would use the standard template library, simulating the behavior of the Outlook rich client when running the same regular expression. To verify evaluation results on Outlook on the web or mobile devices, use your favorite JavaScript regular expression tester.</span></span>

## <a name="if-you-use-an-itemis-itemhasattachment-or-itemhasregularexpressionmatch-rule-have-you-verified-the-related-item-property"></a><span data-ttu-id="40f35-212">如果使用 ItemIs、ItemHasAttachment 或 ItemHasRegularExpressionMatch 规则，是否已验证相关项属性？</span><span class="sxs-lookup"><span data-stu-id="40f35-212">If you use an ItemIs, ItemHasAttachment, or ItemHasRegularExpressionMatch rule, have you verified the related item property?</span></span>

<span data-ttu-id="40f35-213">如果使用 **ItemHasRegularExpressionMatch** 激活规则，请验证 **PropertyName** 属性的值是否为选定项的预期值。</span><span class="sxs-lookup"><span data-stu-id="40f35-213">If you use an **ItemHasRegularExpressionMatch** activation rule, verify whether the value of the **PropertyName** attribute is what you expect for the selected item.</span></span> <span data-ttu-id="40f35-214">下面是调试相应属性的一些提示：</span><span class="sxs-lookup"><span data-stu-id="40f35-214">The following are some tips to debug the corresponding properties:</span></span>

- <span data-ttu-id="40f35-215">如果选定项是邮件，并且你在 **PropertyName** 属性中指定 **BodyAsHTML**，请打开该邮件，然后选择“**查看源代码**”验证该项的 HTML 形式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="40f35-215">If the selected item is a message and you specify **BodyAsHTML** in the **PropertyName** attribute, open the message, and then choose **View Source** to verify the message body in the HTML representation of that item.</span></span>

- <span data-ttu-id="40f35-216">如果选定项是约会，或者激活规则在 **PropertyName** 中指定 **BodyAsPlaintext**，则可使用 Outlook 对象模型和 Windows 版 Outlook 中的 Visual Basic 编辑器：</span><span class="sxs-lookup"><span data-stu-id="40f35-216">If the selected item is an appointment, or if the activation rule specifies **BodyAsPlaintext** in the **PropertyName**, you can use the Outlook object model and the Visual Basic Editor in Outlook on Windows:</span></span>

    1. <span data-ttu-id="40f35-217">确保已启用宏，并且 Outlook 功能区中显示“**开发人员**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="40f35-217">Ensure that macros are enabled and the **Developer** tab is displayed in the ribbon for Outlook.</span></span>

    1. <span data-ttu-id="40f35-218">在 Visual Basic 编辑器中，依次选择“**视图**”和“**即时窗口**”。</span><span class="sxs-lookup"><span data-stu-id="40f35-218">In the Visual Basic Editor, choose **View**, **Immediate Window**.</span></span>

    1. <span data-ttu-id="40f35-219">键入以下内容显示与具体应用场景相关的各个属性。</span><span class="sxs-lookup"><span data-stu-id="40f35-219">Type the following to display various properties depending on the scenario.</span></span>

        - <span data-ttu-id="40f35-220">在 Outlook 资源管理器中选择的邮件或约会项的 HTML 正文：</span><span class="sxs-lookup"><span data-stu-id="40f35-220">The HTML body of the message or appointment item selected in the Outlook explorer:</span></span>

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```
        - <span data-ttu-id="40f35-221">在 Outlook 资源管理器中选择的邮件或约会项的纯文本正文：</span><span class="sxs-lookup"><span data-stu-id="40f35-221">The plain text body of the message or appointment item selected in the Outlook explorer:</span></span>

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```
        - <span data-ttu-id="40f35-222">在当前的 Outlook 检查器中打开的邮件或约会项的 HTML 正文：</span><span class="sxs-lookup"><span data-stu-id="40f35-222">The HTML body of the message or appointment item opened in the current Outlook inspector:</span></span>

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```
        - <span data-ttu-id="40f35-223">在当前的 Outlook 检查器中打开的邮件或约会项的纯文本正文：</span><span class="sxs-lookup"><span data-stu-id="40f35-223">The plain text body of the message or appointment item opened in the current Outlook inspector:</span></span>

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

<span data-ttu-id="40f35-224">如果 **ItemHasRegularExpressionMatch** 激活规则指定 **Subject** 或 **SenderSMTPAddress**，或者你使用 **ItemIs** 或 **ItemHasAttachment** 规则，并且你熟悉或想要使用 MAPI，则可使用 [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) 来验证表 2 中你的规则所依赖的值。</span><span class="sxs-lookup"><span data-stu-id="40f35-224">If the **ItemHasRegularExpressionMatch** activation rule specifies **Subject** or **SenderSMTPAddress**, or if you use an **ItemIs** or **ItemHasAttachment** rule, and you are familiar with or would like to use MAPI, you can use [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) to verify the value in Table 2 that your rule relies on.</span></span>

<span data-ttu-id="40f35-225">**表 2. 激活规则和相应的 MAPI 属性**</span><span class="sxs-lookup"><span data-stu-id="40f35-225">**Table 2. Activation rules and corresponding MAPI properties**</span></span>

|<span data-ttu-id="40f35-226">规则类型</span><span class="sxs-lookup"><span data-stu-id="40f35-226">Type of rule</span></span>|<span data-ttu-id="40f35-227">验证此 MAPI 属性</span><span class="sxs-lookup"><span data-stu-id="40f35-227">Verify this MAPI property</span></span>|
|:-----|:-----|
|<span data-ttu-id="40f35-228">使用 **Subject** 的 **ItemHasRegularExpressionMatch** 规则</span><span class="sxs-lookup"><span data-stu-id="40f35-228">**ItemHasRegularExpressionMatch** rule with **Subject**</span></span>|[<span data-ttu-id="40f35-229">PidTagSubject</span><span class="sxs-lookup"><span data-stu-id="40f35-229">PidTagSubject</span></span>](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|<span data-ttu-id="40f35-230">使用 **SenderSMTPAddress** 的 **ItemHasRegularExpressionMatch** 规则</span><span class="sxs-lookup"><span data-stu-id="40f35-230">**ItemHasRegularExpressionMatch** rule with **SenderSMTPAddress**</span></span>|<span data-ttu-id="40f35-231">[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) 和 [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)</span><span class="sxs-lookup"><span data-stu-id="40f35-231">[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) and [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)</span></span>|
|<span data-ttu-id="40f35-232">**ItemIs**</span><span class="sxs-lookup"><span data-stu-id="40f35-232">**ItemIs**</span></span>|[<span data-ttu-id="40f35-233">PidTagMessageClass</span><span class="sxs-lookup"><span data-stu-id="40f35-233">PidTagMessageClass</span></span>](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|<span data-ttu-id="40f35-234">**ItemHasAttachment**</span><span class="sxs-lookup"><span data-stu-id="40f35-234">**ItemHasAttachment**</span></span>|[<span data-ttu-id="40f35-235">PidTagHasAttachments</span><span class="sxs-lookup"><span data-stu-id="40f35-235">PidTagHasAttachments</span></span>](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

<span data-ttu-id="40f35-236">验证属性值后，即可使用正则表达式评估工具来测试正则表达式是否在该值中找到匹配项。</span><span class="sxs-lookup"><span data-stu-id="40f35-236">After verifying the property value, you can then use a regular expression evaluation tool to test whether the regular expression finds a match in that value.</span></span>

## <a name="does-the-host-application-apply-all-the-regular-expressions-to-the-portion-of-the-item-body-as-you-expect"></a><span data-ttu-id="40f35-237">主机应用程序是否按预期将所有正则表达式应用到项目正文部分？</span><span class="sxs-lookup"><span data-stu-id="40f35-237">Does the host application apply all the regular expressions to the portion of the item body as you expect?</span></span>

<span data-ttu-id="40f35-p123">本部分适用于所有使用正则表达式的激活规则，尤其是应用于项目主体的激活规则，这些规则可能较大，需要较长的时间才能对匹配进行评估。你应该知道，即使激活规则依赖的项目属性具有你所期望的值，主机应用程序也可能无法针对项目属性的整体值评估所有正则表达式。为了提供合理的性能并通过阅读加载项来控制资源过度使用状况，Outlook、Outlook 网页版和移动设备版在运行时遵守激活规则中处理正则表达式的以下限制：</span><span class="sxs-lookup"><span data-stu-id="40f35-p123">This section applies to all activation rules that use regular expressions -- particularly those that are applied to the item body, which may be large in size and take longer to evaluate for matches. You should be aware that even if the item property that an activation rule depends on has the value you expect, the host application may not be able to evaluate all the regular expressions on the entire value of the item property. To provide reasonable performance and to control excessive resource usage by a read add-in, Outlook, Outlook on the web and mobile devices observe the following limits on processing regular expressions in activation rules at run time:</span></span>

- <span data-ttu-id="40f35-p124">评估的项目正文的大小 — 主机应用程序在其中评估正则表达式的项目正文部分存在限制。这些限制取决于主机应用程序、组成要素和项目正文的格式。请参阅[激活限制和适用于 Outlook 加载项的 JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) 中表 2 中的详细信息。</span><span class="sxs-lookup"><span data-stu-id="40f35-p124">The size of the item body evaluated -- There are limits to the portion of an item body on which the host application evaluates a regular expression. These limits depend on the host application, form factor, and format of the item body. See the details in Table 2 in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).</span></span>

- <span data-ttu-id="40f35-p125">正则表达式匹配的数量 - Outlook 富客户端、Outlook 网页版和移动设备版分别返回最多 50 个正则表达式匹配项。这些匹配项是唯一的，重复的匹配不计入此限制。请勿假定返回的匹配项有任何顺序，也不要假定 Outlook 富客户端中的顺序与 Outlook 网页版和移动设备版中的顺序相同。如果希望激活规则中存在与正则表达式匹配的许多匹配项，并且丢失匹配项，则可能会超出此限制。</span><span class="sxs-lookup"><span data-stu-id="40f35-p125">Number of regular expression matches -- The Outlook rich clients, and Outlook on the web and mobile devices each returns a maximum of 50 regular expression matches. These matches are unique, and duplicate matches do not count against this limit. Do not assume any order to the returned matches, and do not assume the order in an Outlook rich client is the same as that in Outlook on the web and mobile devices. If you expect many matches to regular expressions in your activation rules, and you're missing a match, you may be exceeding this limit.</span></span>

- <span data-ttu-id="40f35-p126">正则表达式匹配项的长度 — 主机应用程序将返回的正则表达式匹配项的长度存在限制。主机应用程序不包括超出限制的任何匹配项，并且不显示任何警告消息。你可以使用其他正则表达式评估工具或独立的 C++ 测试程序运行你的正则表达式，以验证你是否具有超出此类限制的匹配项。表 3 总结了这些限制。有关详细信息，请参阅[激活限制和适用于 Outlook 加载项的 JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) 中的表 3。</span><span class="sxs-lookup"><span data-stu-id="40f35-p126">Length of a regular expression match -- There are limits to the length of a regular expression match that the host application would return. The host application does not include any match above the limit and does not display any warning message. You can run your regular expression using other regex evaluation tools or a stand-alone C++ test program to verify whether you have a match that exceeds such limits. Table 3 summarizes the limits. For more information, see Table 3 in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).</span></span>

    <span data-ttu-id="40f35-253">**表 3.正则表达式匹配的长度限制**</span><span class="sxs-lookup"><span data-stu-id="40f35-253">**Table 3. Length limits for a regular expression match**</span></span>

    |<span data-ttu-id="40f35-254">正则表达式匹配项的长度限制</span><span class="sxs-lookup"><span data-stu-id="40f35-254">Limit on length of a regex match</span></span>|<span data-ttu-id="40f35-255">Outlook 富客户端</span><span class="sxs-lookup"><span data-stu-id="40f35-255">Outlook rich clients</span></span>|<span data-ttu-id="40f35-256">Outlook 网页版或移动设备版</span><span class="sxs-lookup"><span data-stu-id="40f35-256">Outlook on the web or mobile devices</span></span>|
    |:-----|:-----|:-----|
    |<span data-ttu-id="40f35-257">项目正文采用纯文本</span><span class="sxs-lookup"><span data-stu-id="40f35-257">Item body is plain text</span></span>|<span data-ttu-id="40f35-258">1.5 KB</span><span class="sxs-lookup"><span data-stu-id="40f35-258">1.5 KB</span></span>|<span data-ttu-id="40f35-259">3 KB</span><span class="sxs-lookup"><span data-stu-id="40f35-259">3 KB</span></span>|
    |<span data-ttu-id="40f35-260">项目正文采用 HTML</span><span class="sxs-lookup"><span data-stu-id="40f35-260">Item body is HTML</span></span>|<span data-ttu-id="40f35-261">3 KB</span><span class="sxs-lookup"><span data-stu-id="40f35-261">3 KB</span></span>|<span data-ttu-id="40f35-262">3 KB</span><span class="sxs-lookup"><span data-stu-id="40f35-262">3 KB</span></span>|

- <span data-ttu-id="40f35-p127">评估阅读加载项的所有正则表达式所花费的时间 - 对于某个 Outlook 富客户端：默认情况下，对于每个阅读加载项，Outlook 必须在 1 秒钟内完成对其激活规则中的所有正则表达式的评估。否则，如果 Outlook 无法完成评估，则 Outlook 最多尝试 3 次并禁用该加载项。Outlook 会在通知栏中显示一条消息，指示该加载项已禁用。正则表达式可用的时间可通过设置组策略或注册表项来进行修改。</span><span class="sxs-lookup"><span data-stu-id="40f35-p127">Time spent on evaluating all regular expressions of a read add-in for an Outlook rich client: By default, for each read add-in, Outlook must finish evaluating all the regular expressions in its activation rules within 1 second. Otherwise Outlook retries up to three times and disables the add-in if Outlook cannot complete the evaluation. Outlook displays a message in the notification bar that the add-in has been disabled. The amount of time available for your regular expression can be modified by setting a group policy or a registry key.</span></span> 

   > [!NOTE]
   > <span data-ttu-id="40f35-267">如果 Outlook 富客户端禁用某个读取加载项，则无法在 Outlook 富客户端、Outlook 网页版和移动设备版上的同一邮箱中使用该读取加载项。</span><span class="sxs-lookup"><span data-stu-id="40f35-267">If the Outlook rich client disables a read add-in, the read add-in is not available for use for the same mailbox on the Outlook rich client, and Outlook on the web and mobile devices.</span></span>

## <a name="see-also"></a><span data-ttu-id="40f35-268">另请参阅</span><span class="sxs-lookup"><span data-stu-id="40f35-268">See also</span></span>

- [<span data-ttu-id="40f35-269">部署和安装 Outlook 加载项以进行测试</span><span class="sxs-lookup"><span data-stu-id="40f35-269">Deploy and install Outlook add-ins for testing</span></span>](testing-and-tips.md)
- [<span data-ttu-id="40f35-270">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="40f35-270">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="40f35-271">使用正则表达式激活规则显示 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="40f35-271">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="40f35-272">Outlook 外接程序的激活和 JavaScript API 限制</span><span class="sxs-lookup"><span data-stu-id="40f35-272">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="40f35-273">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="40f35-273">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)
