---
title: Outlook 加载项的激活和 API 使用限制
description: 请注意某些激活和 API 使用指南，并在这些限制范围内实施加载项。
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: b09886e49b0d980dbbf2465df7d077cd16a04f4d
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616012"
---
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Outlook 加载项的激活和 JavaScript API 限制

为了向 Outlook 外接程序的用户提供令人满意的体验，您必须了解特定的激活和 API 使用准则，并执行外接程序使其不超过这些限制。 这些准则的存在使单个加载项不能要求Exchange Server或 Outlook 花费异常长的时间来处理其激活规则或对 Office JavaScript API 的调用，从而影响 Outlook 和其他加载项的总体用户体验。这些限制适用于在外接程序清单中设计激活规则，以及使用自定义属性、漫游设置、收件人、Exchange Web 服务 (EWS) 请求和响应以及异步调用。

> [!NOTE]
> 如果外接程序在 Outlook 富客户端上运行，则还必须验证外接程序是否在某些运行时资源使用限制内执行。

## <a name="limits-on-where-add-ins-activate"></a>外接程序激活位置的限制

若要详细了解加载项在何处执行且不激活，请参阅 Outlook 外接程序概述页的“外接程序”部分 [可用的邮箱项](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) 。

## <a name="limits-for-activation-rules"></a>激活规则的限制

为 Outlook 外接程序设计激活规则时，请遵循以下准则：

- 将清单的大小限制为 256 KB。 如果超过该限制，则无法安装 Exchange 邮箱的 Outlook 加载项。

- 可为外接程序最多指定 15 条激活规则。 如果超出该限制，则无法安装加载项。

- 如果您对所选项目的正文使用 [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) 规则，预计 Outlook 富客户端将仅对正文的前 1 MB 应用规则，而不会超过此限制应用于正文的其他部分。 如果仅在正文的第一个 MB 之后才存在匹配项，则加载项不会激活。 如果您期望这成为一种可能的方案，请重新设计激活条件。

- 如果在或 [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) 规则中`ItemHasKnownEntity`使用正则表达式，请注意以下限制和准则，这些限制和准则通常适用于任何 Outlook 应用程序，以及表 1、2 和 3 中所述的限制和准则，这些限制和准则因应用程序而异。
  - 在加载项的激活规则中，最多只指定五个正则表达式。 如果超出该限制，则无法安装加载项。
  - 指定正则表达式，使预期结果由 `getRegExMatches` 前 50 个匹配项中的方法调用返回。
  - **重要** 提示：文本基于匹配正则表达式后产生的字符串突出显示。 不过，突出显示的出现可能与实际正则表达式断言的结果不完全匹配，例如负前 `(?!text)`观、后 `(?<=text)`看和负面观望 `(?<!text)`。 例如，如果在“Like under、under score 和 underscore”上使用正则表达式 `under(?!score)` ，则字符串“under”将突出显示所有匹配项，而不只是前两个匹配项。

表 1 列出了限制，并描述了 Outlook 富客户端与Outlook 网页版或移动设备之间对正则表达式的支持差异。 这种支持不依赖于任何特定类型的设备和项目正文。

**表 1.各种正则表达式支持的一般区别**

|Outlook 富客户端|Outlook 网页版或移动设备版|
|:-----|:-----|
|使用作为 Visual Studio 标准模板库一部分提供的 C++ 正则表达式引擎。该引擎使用 ECMAScript 5 标准编译。 |使用属于 JavaScript 一部分的正则表达式评估，由浏览器提供，且支持 ECMAScript 5 超集。|
|由于正则表达式引擎不同，因此需要包含基于预定义字符类的自定义字符类的正则表达式在 Outlook 富客户端中返回与Outlook 网页版或移动设备不同的结果。<br/><br/>例如，正则表达式 `[\s\S]{0,100}` 匹配空格或非空格的单个字符的任何数字（介于 0 到 100 之间）。 此正则表达式在 Outlook 富客户端中返回的结果与Outlook 网页版和移动设备不同。<br/><br/>应将正则表达式重写为 `(\s\|\S){0,100}` 解决方法。 此变通正则表达式与任意数量（0 到 100）的空格字符或非空格字符匹配。<br/><br/>应在每个 Outlook 客户端上彻底测试每个正则表达式，如果正则表达式返回不同的结果，请重写正则表达式。 |应在每个 Outlook 客户端上彻底测试每个正则表达式，如果正则表达式返回不同的结果，请重写正则表达式。|
|默认情况下，外接程序的所有正则表达式的计算时间限制为 1 秒。 超出此限制将导致最多重新计算 3 次。 除了重新计算限制之外，Outlook 富客户端将禁止外接程序在任何 Outlook 客户端中为同一邮箱运行。<br/><br/>管理员可以使用 `OutlookActivationAlertThreshold` 和 `OutlookActivationManagerRetryLimit` 注册表项替代这些评估限制。|不要支持与 Outlook 富客户端相同的资源监视或注册表设置。 但是，对于所有 Outlook 客户端上的同一邮箱，使用需要过多评估时间的正则表达式的外接程序会被禁用。|

表 2 列出了这些限制并介绍了每一个 Outlook 应用了正则表达式的项正文部分的区别。如果对项正文应用了正则表达式，则其中某些限制取决于设备和项正文的类型。

**表 2.计算的项正文的大小限制**

||Outlook 富客户端|移动设备上的 Outlook|Outlook 网页版|
|:-----|:-----|:-----|:-----|
|**外形规格**|任何受支持的设备。|Android 智能手机、iPad 或 iPhone。|除 Android 智能手机、iPad 和 iPhone 以外的任何受支持的设备。|
|**纯文本项正文**|对正文数据的第一个 1 MB 而不对超出该限制的其余正文应用正则表达式。|仅当正文少于 16,000 个字符时激活加载项。|仅当正文少于 500,000 个字符时激活加载项。|
|**HTML 项正文**|对正文数据的第一个 512 KB 而不对超出该限制的其余正文应用正则表达式。（实际的字符数取决于范围可为每字符 1 到 4 字节的编码。）|对前 64,000 个字符（包括 HTML 标记字符）而不对超出该限制的其余正文应用正则表达式。|仅当正文少于 500,000 个字符时激活加载项。|

表 3 列出了限制，并描述了每个 Outlook 客户端在评估正则表达式后返回的匹配项的差异。 这种支持不依赖于任何特定设备类型，但是，如果对项正文应用了正则表达式，则该支持可能依赖于项正文的类型。

**表 3.返回的匹配项限制**

||Outlook 富客户端|Outlook 网页版或移动设备版|
|:-----|:-----|:-----|
|**返回的匹配项的顺序**|假定`getRegExMatches`在 Outlook 富客户端中应用于同一项的同一正则表达式与在Outlook 网页版或移动设备中应用的相同正则表达式返回匹配项。|假定`getRegExMatches`在 Outlook 富客户端中返回匹配项的顺序不同于Outlook 网页版或移动设备中的匹配项。|
|**纯文本项正文**|`getRegExMatches` 返回最多 1，536 (1.5 KB) 字符的任何匹配项，最多 50 个匹配项。<br/><br/>**注意**： `getRegExMatches` 在返回的数组中，不会以任何特定顺序返回匹配项。 一般情况下，假设在同一项上应用的相同正则表达式的 Outlook 富客户端中匹配的顺序不同于Outlook 网页版和移动设备中的匹配顺序。|`getRegExMatches` 返回最多 3，072 (3 KB) 字符的任何匹配项，最多 50 个匹配项。|
|**HTML 项正文**|`getRegExMatches` 返回最多 3，072 (3 KB) 字符的任何匹配项，最多 50 个匹配项。<br/> <br/> **注意**： `getRegExMatches` 在返回的数组中，不会以任何特定顺序返回匹配项。 一般情况下，假设在同一项上应用的相同正则表达式的 Outlook 富客户端中匹配的顺序不同于Outlook 网页版和移动设备中的匹配顺序。|`getRegExMatches` 返回最多 3，072 (3 KB) 字符的任何匹配项，最多 50 个匹配项。|

## <a name="limits-for-javascript-api"></a>JavaScript API 的限制

除了前面的激活规则准则外，每个 Outlook 客户端在 JavaScript 对象模型中强制实施某些限制，如表 4 中所述。

**表 4.使用 Office JavaScript API 获取或设置某些数据的限制**

|功能|限制|相关 API|说明|
|:-----|:-----|:-----|:-----|
|自定义属性|2500 个字符|[CustomProperties](/javascript/api/outlook/office.customproperties) 对象<br/> <br/>[item.loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法|约会或邮件项目的所有自定义属性的限制。 如果加载项的所有自定义属性的总大小超过此限制，则所有 Outlook 客户端都会返回错误。|
|漫游设置|32 KB 字符数|[RoamingSettings](/javascript/api/outlook/office.roamingsettings) 对象<br/><br/> [context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#properties) 属性|外接程序的所有漫游设置的限制。 如果设置超过此限制，所有 Outlook 客户端都会返回错误。|
|Internet 标头：|Exchange Online中每条消息 256 KB<br/><br/>由组织本地 Exchange 中的管理员确定的标头大小限制|[InternetHeaders.setAsync](/javascript/api/outlook/office.internetheaders) 方法|可应用于消息的标头的总大小限制。|
|正在提取已知实体|2000 个字符|[item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法<br/> <br/>[item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法<br/> <br/>[item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法|在项目正文上提取常见实体的 Exchange Server 限制。 Exchange Server 将忽略超过该限制的实体。 请注意，此限制与加载项是否使用 `ItemHasKnownEntity` 规则无关。|
|Exchange Web 服务|1 MB 字符数|[mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法|请求或响应调用的限制 `Mailbox.makeEwsRequestAsync` 。|
|收件人|100 位收件人|[item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性<br/> <br/>[item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性<br/> <br/>[item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性<br/> <br/>[item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性<br/> <br/>[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)) 方法<br/> <br/>[Recipient.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) 方法<br/> <br/>[Recipient.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) 方法|在每个属性中指定的对收件人的限制。|
|显示名称|255 个字符|[EmailAddressDetails.displayName](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-displayname-member) 属性<br/><br/> [Recipients](/javascript/api/outlook/office.recipients) 对象<br/><br/> `item.requiredAttendees` 财产<br/><br/> `item.optionalAttendees` 财产 <br/><br/>`item.to` 财产 <br/><br/>`item.cc` 财产|约会或邮件中显示名称的长度限制。|
|设置主题|255 个字符|[Mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法<br/><br/> [Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1)) 方法|新的约会窗体中的主题限制，或设置约会或邮件主题的限制。|
|设置地点|255 个字符|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) 方法|设置约会或会议请求地点的限制。|
|新的约会窗体的正文|32 KB 字符数|`Mailbox.displayNewAppointmentForm` 方法|新的约会窗体中正文的限制。|
|显示现有项目的正文|32 KB 字符数|[mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法<br/><br/> [mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法|对于Outlook 网页版和移动设备：现有约会或消息窗体中正文的限制。|
|设置正文|1 MB 字符数|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)) 方法<br/> <br/>[Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))<br/><br/>[Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1)) 方法|设置约会或邮件项目正文的限制。|
|附件数|Outlook 网页版和移动设备上的 499 个文件|[item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法|限制可附加到发送项目的文件数量。 Outlook 网页版和移动设备通常限制通过用户界面`addFileAttachmentAsync`和移动设备附加多达 499 个文件。 Outlook 富客户端不具体限制文件附件的数量。 但是，所有 Outlook 客户端都遵守用户Exchange Server已配置的附件大小的限制。 请查看下一行获取“附件大小”信息。|
|附件大小|取决于 Exchange Server|`item.addFileAttachmentAsync` 方法|对项目所有附件的大小有限制，管理员可以在用户邮箱的 Exchange Server 上配置此限制。对于 Outlook 富客户端，这限制了项目的附件数量。 对于Outlook 网页版和移动设备，这两个限制（附件数和所有附件的大小）越小，则限制项目的实际附件。|
|附件的文件名|255 个字符|`item.addFileAttachmentAsync` 方法|要添加到项目的附件的文件名长度限制。|
|附件的 URI|2048 个字符|`item.addFileAttachmentAsync` 方法|要添加为项目附件的文件名 URI 的限制。|
|附件 ID|100 个字符|[item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法<br/><br/> [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法|要添加或从项目中删除的附件 ID 的长度限制。|
|异步调用|3 次调用|`item.addFileAttachmentAsync` 方法<br/><br/>`item.addItemAttachmentAsync` 方法<br/><br/><br/>`item.removeAttachmentAsync` 方法<br/><br/> [Body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1)) 方法<br/><br/>`Body.prependAsync` 方法<br/><br/>`Body.setSelectedDataAsync` 方法<br/><br/> [CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) 方法<br/><br/><br/> [项目。LoadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法<br/><br/><br/> [Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) 方法<br/><br/>`Location.setAsync` 方法<br/><br/> [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法<br/><br/> [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法<br/><br/> [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法<br/><br/>`Recipients.addAsync` 方法<br/><br/> [Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) 方法<br/><br/>`Recipients.setAsync` 方法<br/><br/> [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)) 方法<br/><br/> [Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) 方法<br/><br/>`Subject.setAsync` 方法<br/><br/> [Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) 方法<br/><br/> [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1)) 方法|对于Outlook 网页版或移动设备：任意一次同时异步调用数的限制，因为浏览器只允许对服务器进行有限数量的异步调用。 |

## <a name="see-also"></a>另请参阅

- [部署和安装 Outlook 加载项以进行测试](testing-and-tips.md)
- [Outlook 外接程序的隐私、权限和安全性](../concepts/privacy-and-security.md)
