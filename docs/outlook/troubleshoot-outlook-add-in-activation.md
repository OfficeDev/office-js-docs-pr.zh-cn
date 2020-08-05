---
title: Outlook 上下文加载项激活故障排查
description: 如果加载项未按预期激活，应考虑以下几个方面的可能原因。
ms.date: 08/03/2020
localization_priority: Normal
ms.openlocfilehash: e9eba8abd1207c0c521fc87e310325529c9f24ac
ms.sourcegitcommit: a3b743598025466bad19177e0ba9ca94ea66d490
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/04/2020
ms.locfileid: "46547540"
---
# <a name="troubleshoot-outlook-add-in-activation"></a>Outlook 加载项激活故障排查

Outlook 上下文加载项激活基于加载项清单中的激活规则。在当前选定项的条件满足加载项的激活规则时，主机应用程序激活，并在 Outlook UI 中显示加载项按钮（用于撰写加载项的加载项选择窗格，用于阅读加载项的加载项条）。但是，如果加载项未按预期激活，应考虑以下几个方面的原因。

## <a name="is-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>用户邮箱是否位于不低于 Exchange 2013 版本的 Exchange Server 上？

首先，确保你正在测试的用户电子邮件帐户位于至少为 Exchange 2013 的某个版本的 Exchange Server 上。如果你正在使用在Exchange 2013 之后发布的特定功能，那么请确保用户的帐户使用合适的 Exchange 版本。

你可使用以下方法之一验证 Exchange 2013 的版本：

- 咨询你的 Exchange Server 管理员。

- 若要在 Outlook 网页版或移动设备版上测试加载项，请在脚本调试器（例如，Internet Explorer 随附的 JScript 调试器）中，查找指定脚本加载位置的 **script** 标记的 **src** 属性。路径应包含子字符串 **owa/15.0.516.x/owa2/...**，其中 **15.0.516.x** 表示 Exchange Server 版本（如 **15.0.516.2**）。

- 也可以使用 [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) 属性来验证版本。在 Outlook 网页版和移动设备版上，此属性会返回 Exchange Server 版本。

- 如果能够在 Outlook 上测试加载项，则可使用采用 Outlook 对象模型和 Visual Basic 编辑器的以下简单调试技术：

    1. 首先，确认已对 Outlook 启用了宏。依次选择“**文件**”、“**选项**”、“**信任中心**”、“**信任中心设置**”、“**宏设置**”。确保在“信任中心”中选择了“**为所有宏提供通知**”。还应确保在 Outlook 启动过程中选择了“**启用宏**”。

    1. 在功能区的“**开发人员**”选项卡上，选择“**Visual Basic**”。

       > [!NOTE]
       > 看不到“**开发人员**”选项卡？请参阅[如何：在功能区上显示“开发人员”选项卡](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon)，启用此选项卡。

    1. 在 Visual Basic 编辑器中，依次选择“**视图**”和“**即时窗口**”。

    1. 在即时窗口中键入以下内容以显示 Exchange Server 的版本。返回值的主版本必须等于或大于 15。

       - 如果用户的配置文件中只有一个 Exchange 帐户：

       ```vb
        ?Session.ExchangeMailboxServerVersion
       ```

       - 如果同一用户配置文件中有多个 Exchange 帐户（`emailAddress` 表示包含用户主 SMTP 地址的字符串）：

       ```vb
        ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
       ```

## <a name="is-the-add-in-disabled"></a>加载项否已禁用？

任何 Outlook 富客户端可出于性能原因禁用加载项，这些原因包括超出 CPU 内核或内存的使用阈值、超出崩溃容忍度以及超出处理加载项的所有正则表达式的时间。发生这种情况时，Outlook 富客户端会显示一条禁用加载项的通知。

> [!NOTE]
> 仅 Outlook 富客户端可监视资源使用状况，但如果在 Outlook 富客户端中禁用加载项，也会在 Outlook 网页版和移动设备版中禁用此加载项。

使用以下方法之一，验证加载项是否已禁用：

- 在 Outlook 网页版中，直接登录电子邮件帐户，选择“设置”图标，然后选择“**管理加载项**”转到 Exchange 管理中心，可在此处验证管理加载项是否已启用。

- 在 Windows 版 Outlook 中，转到 Backstage 视图并选择“**管理加载项**”。登录 Exchange 管理中心验证加载项是否已启用。

- 在 Mac 版 Outlook 中，选择加载项栏中的“**管理加载项**”。登录 Exchange 管理中心验证加载项是否已启用。

## <a name="does-the-tested-item-support-outlook-add-ins-is-the-selected-item-delivered-by-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>已测试项是否支持 Outlook 加载项？所选项目是否由至少为 Exchange 2013 的某个版本的 Exchange Server 提供？

如果你的 Outlook 加载项为阅读加载项，并且应该在用户查看消息（包括电子邮件、会议请求、响应和取消）或约会时激活，尽管这些项目通常支持加载项，但还是存在例外情况。 检查所选的项目是否是 [Outlook 加载项未激活列表](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)中的一项。

此外，由于约会始终以 RTF 格式保存，因此指定 [BodyAsHTML](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 的 **PropertyName** 值的 **ItemHasRegularExpressionMatch** 规则不会对以纯文本或 RTF 格式保存的约会或邮件激活加载项。

即使某邮件项不是以上类型之一，如果该项不是使用至少为 Exchange 2013 的某个版本的 Exchange Server 传递，则不会在该项上确定已知实体和属性（如发件人的 SMTP 地址）。依赖这些实体或属性的任何激活规则将不会得到满足，并且加载项将不会激活。

如果您的加载项为撰写加载项并且应该在用户撰写邮件或会议请求时激活，请确保该项目未受 IRM 保护。 但是，从 Outlook 内部版本13120.1000 在 Windows 上开始，外接程序现在可以在受 IRM 保护的项目上激活。  有关预览中此支持的详细信息，请参阅[ (IRM) 的受信息权限管理保护的项上的外接程序激活](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)。

## <a name="is-the-add-in-manifest-installed-properly-and-does-outlook-have-a-cached-copy"></a>加载项清单是否安装正确，Outlook 是否有已缓存副本？

此方案仅适用于 Windows 版 Outlook。正常情况下，为邮箱安装 Outlook 加载项时，Exchange Server 会将加载项清单从你指示的位置复制到该 Exchange Server 上的邮箱。每次启动 Outlook 时，它都会将为该邮箱安装的所有清单读取到以下位置的临时缓存中：

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

例如，对于用户 John，缓存可能位于 C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF。

> [!IMPORTANT]
> 对于 Windows 上的 Outlook 2013，请使用15.0 而不是16.0，以便位置为：
>
> ```text
> %LocalAppData%\Microsoft\Office\15.0\WEF
> ```

如果无法对任何项目激活加载项，则清单可能未正确安装在 Exchange Server 上，或者 Outlook 未在启动时正确读取清单。使用 Exchange 管理中心确保已为您的邮箱安装和启用加载项，并在必要时重新启动 Exchange Server。

图 1 显示验证 Outlook 是否具有有效版本的清单的步骤摘要。

**图 1.验证 Outlook 是否已正确缓存清单的步骤的流程图**

![用于检查清单的流程图](../images/troubleshoot-manifest-flow.png)

以下过程描述详细信息。

1. 如果你已在 Outlook 打开时修改了清单，并且未使用 Visual Studio 2012 或 Visual Studio 的更高版本开发加载项，则应卸载加载项，并使用 Exchange 管理中心重新安装它。

1. 重新启动 Outlook 并测试 Outlook 现在是否已激活加载项。

1. 如果 Outlook 无法激活加载项，则检查 Outlook 是否具有加载项清单的正确缓存副本。请查看以下路径：

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    可以在下列子文件夹中找到清单：

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > 下面的示例展示了为用户 John 的邮箱安装的清单路径：
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    验证要测试的加载项的清单是否在已缓存清单中。

1. 如果清单在缓存中，请跳过本节的其余部分，并考虑本节后面的其他可能原因。

1. 如果清单不在缓存中，请检查 Outlook 是否已确实从 Exchange Server 中成功读取清单。为此，请使用 Windows 事件查看器：

    1. 在“**Windows 日志**”下，选择“**应用程序**”。

    1. 查找其事件 ID 等于 63（表示 Outlook 从 Exchange Server 下载清单）的近期事件。

    1. 如果 Outlook 成功读取了清单，则记录的事件应包含以下说明：

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        然后，跳过本节的其余部分，并考虑本节后面的其他可能原因。

1. 如果看不到成功事件，请关闭 Outlook，再删除以下路径中的所有清单：

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    启动 Outlook，并测试 Outlook 现在是否已激活加载项。

1. 如果 Outlook 无法激活加载项，请返回到第 3 步，再次确认 Outlook 是否已正确读取清单。

## <a name="is-the-add-in-manifest-valid"></a>加载项清单有效吗？

请参阅[验证并排查清单问题](../testing/troubleshoot-manifest.md)来调试加载项清单问题。

## <a name="are-you-using-the-appropriate-activation-rules"></a>使用的激活规则是否合适？

自 Office 加载项清单架构的版本 1.1 起，你可以创建当用户位于撰写窗体（撰写加载项）或阅读窗体（阅读加载项）中时激活的加载项。确保为加载项将在其中激活的每种窗体类型指定相应的激活规则。例如，你可以仅使用 [ItemIs](../reference/manifest/rule.md#itemis-rule) 规则（**FormType** 属性设置为 **Edit** 或 **ReadOrEdit**）激活撰写加载项，你无法使用任何其他类型的规则，例如用于撰写加载项的 [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 和 [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 规则。有关详细信息，请参阅 [Outlook 加载项的激活规则](activation-rules.md)。

## <a name="if-you-use-a-regular-expression-is-it-properly-specified"></a>如果使用正则表达式，该表达式的指定是否正确？

由于激活规则中的正则表达式是阅读加载项的 XML 清单文件的一部分，因此当正则表达式使用特定字符时，请务必遵守 XML 处理器支持的相应转义序列。表 1 列出了这些特殊字符。

**表 1.正则表达式的转义序列**

|**字符**|**说明**|**要使用的转义序列**|
|:-----|:-----|:-----|
|`"`|双引号|&amp;quot;|
|`&`|与号|&amp;amp;|
|`'`|撇号|&amp;apos;|
|`<`|小于号|&amp;lt;|
|`>`|大于号|&amp;gt;|

## <a name="if-you-use-a-regular-expression-is-the-read-add-in-activating-in-outlook-on-the-web-or-mobile-devices-but-not-in-any-of-the-outlook-rich-clients"></a>如果使用正则表达式，阅读加载项是否在 Outlook 网页版或移动设备版（而不是个别 Outlook 富客户端）中激活？

Outlook 富客户端使用的正则表达式引擎与 Outlook 网页版和移动设备版使用的正则表达式引擎不同。Outlook 富客户端使用作为 Visual Studio 标准模板库的一部分提供的 C++ 正则表达式引擎。此引擎符合 ECMAScript 5 标准。Outlook 网页版和移动设备版使用属于 JavaScript 一部分的正则表达式评估，由浏览器提供，且支持 ECMAScript 5 超集。

在大多数情况下，这些主机应用程序在激活规则中为相同的正则表达式找到相同的匹配项，但也有例外。例如，如果正则表达式包含基于预定义字符类的自定义字符类，则 Outlook 富客户端可能会返回与 Outlook 网页版和移动设备版不同的结果。作为示例，在其中包含速记字符类 `[\d\w]` 的字符类将返回不同的结果。在这种情况下，为避免不同主机上出现不同结果，请改用 `(\d|\w)`。

全面测试正则表达式。如果返回不同的结果，请重写正则表达式以兼容两个引擎。要验证 Outlook 富客户端上的评估结果，请编写一个小型 C++ 程序，该程序可将正则表达式应用于你尝试匹配的文本示例。在 Visual Studio 上运行时，C++ 测试程序将使用标准模板库，在运行相同正则表达式时模拟 Outlook 富客户端的行为。要验证 Outlook 网页版或移动设备版上的评估结果，请使用你喜爱的 JavaScript 正则表达式测试程序。

## <a name="if-you-use-an-itemis-itemhasattachment-or-itemhasregularexpressionmatch-rule-have-you-verified-the-related-item-property"></a>如果使用 ItemIs、ItemHasAttachment 或 ItemHasRegularExpressionMatch 规则，是否已验证相关项属性？

如果使用 **ItemHasRegularExpressionMatch** 激活规则，请验证 **PropertyName** 属性的值是否为选定项的预期值。 下面是调试相应属性的一些提示：

- 如果选定项是邮件，并且你在 **PropertyName** 属性中指定 **BodyAsHTML**，请打开该邮件，然后选择“**查看源代码**”验证该项的 HTML 形式的邮件正文。

- 如果选定项是约会，或者激活规则在 **PropertyName** 中指定 **BodyAsPlaintext**，则可使用 Outlook 对象模型和 Windows 版 Outlook 中的 Visual Basic 编辑器：

    1. 确保已启用宏，并且 Outlook 功能区中显示“**开发人员**”选项卡。

    1. 在 Visual Basic 编辑器中，依次选择“**视图**”和“**即时窗口**”。

    1. 键入以下内容显示与具体应用场景相关的各个属性。

        - 在 Outlook 资源管理器中选择的邮件或约会项的 HTML 正文：

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```
        - 在 Outlook 资源管理器中选择的邮件或约会项的纯文本正文：

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```
        - 在当前的 Outlook 检查器中打开的邮件或约会项的 HTML 正文：

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```
        - 在当前的 Outlook 检查器中打开的邮件或约会项的纯文本正文：

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

如果 **ItemHasRegularExpressionMatch** 激活规则指定 **Subject** 或 **SenderSMTPAddress**，或者你使用 **ItemIs** 或 **ItemHasAttachment** 规则，并且你熟悉或想要使用 MAPI，则可使用 [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) 来验证表 2 中你的规则所依赖的值。

**表 2. 激活规则和相应的 MAPI 属性**

|规则类型|验证此 MAPI 属性|
|:-----|:-----|
|使用 **Subject** 的 **ItemHasRegularExpressionMatch** 规则|[PidTagSubject](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|使用 **SenderSMTPAddress** 的 **ItemHasRegularExpressionMatch** 规则|[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) 和 [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)|
|**ItemIs**|[PidTagMessageClass](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|**ItemHasAttachment**|[PidTagHasAttachments](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

验证属性值后，即可使用正则表达式评估工具来测试正则表达式是否在该值中找到匹配项。

## <a name="does-the-host-application-apply-all-the-regular-expressions-to-the-portion-of-the-item-body-as-you-expect"></a>主机应用程序是否按预期将所有正则表达式应用到项目正文部分？

本部分适用于所有使用正则表达式的激活规则，尤其是应用于项目主体的激活规则，这些规则可能较大，需要较长的时间才能对匹配进行评估。你应该知道，即使激活规则依赖的项目属性具有你所期望的值，主机应用程序也可能无法针对项目属性的整体值评估所有正则表达式。为了提供合理的性能并通过阅读加载项来控制资源过度使用状况，Outlook、Outlook 网页版和移动设备版在运行时遵守激活规则中处理正则表达式的以下限制：

- 评估的项目正文的大小 — 主机应用程序在其中评估正则表达式的项目正文部分存在限制。这些限制取决于主机应用程序、组成要素和项目正文的格式。请参阅[激活限制和适用于 Outlook 加载项的 JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) 中表 2 中的详细信息。

- 正则表达式匹配的数量 - Outlook 富客户端、Outlook 网页版和移动设备版分别返回最多 50 个正则表达式匹配项。这些匹配项是唯一的，重复的匹配不计入此限制。请勿假定返回的匹配项有任何顺序，也不要假定 Outlook 富客户端中的顺序与 Outlook 网页版和移动设备版中的顺序相同。如果希望激活规则中存在与正则表达式匹配的许多匹配项，并且丢失匹配项，则可能会超出此限制。

- 正则表达式匹配项的长度 — 主机应用程序将返回的正则表达式匹配项的长度存在限制。主机应用程序不包括超出限制的任何匹配项，并且不显示任何警告消息。你可以使用其他正则表达式评估工具或独立的 C++ 测试程序运行你的正则表达式，以验证你是否具有超出此类限制的匹配项。表 3 总结了这些限制。有关详细信息，请参阅[激活限制和适用于 Outlook 加载项的 JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) 中的表 3。

    **表 3.正则表达式匹配的长度限制**

    |正则表达式匹配项的长度限制|Outlook 富客户端|Outlook 网页版或移动设备版|
    |:-----|:-----|:-----|
    |项目正文采用纯文本|1.5 KB|3 KB|
    |项目正文采用 HTML|3 KB|3 KB|

- 评估阅读加载项的所有正则表达式所花费的时间 - 对于某个 Outlook 富客户端：默认情况下，对于每个阅读加载项，Outlook 必须在 1 秒钟内完成对其激活规则中的所有正则表达式的评估。否则，如果 Outlook 无法完成评估，则 Outlook 最多尝试 3 次并禁用该加载项。Outlook 会在通知栏中显示一条消息，指示该加载项已禁用。正则表达式可用的时间可通过设置组策略或注册表项来进行修改。 

   > [!NOTE]
   > 如果 Outlook 富客户端禁用某个读取加载项，则无法在 Outlook 富客户端、Outlook 网页版和移动设备版上的同一邮箱中使用该读取加载项。

## <a name="see-also"></a>另请参阅

- [部署和安装 Outlook 加载项以进行测试](testing-and-tips.md)
- [Outlook 加载项的激活规则](activation-rules.md)
- [使用正则表达式激活规则显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook 外接程序的激活和 JavaScript API 限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [验证并排查清单问题](../testing/troubleshoot-manifest.md)
