---
title: 将字符串作为 Outlook 加载项中的已知实体进行匹配
description: 使用 Office JavaScript API，您可以获取与特定已知实体匹配的字符串以进行进一步处理。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: a8dfb20405f4c3add35ca1ea646ffe69fc776a26
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325338"
---
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a><span data-ttu-id="87729-103">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="87729-103">Match strings in an Outlook item as well-known entities</span></span>

<span data-ttu-id="87729-p101">发送邮件或会议请求项之前，Exchange Server 将分析项目的内容、标识和标记类似于 Exchange 已知实体的主题和正文中的特定字符串，例如，电子邮件地址、电话号码和 URL。邮件和会议请求通过标有已知实体的 Outlook 收件箱中的 Exchange Server 传递。</span><span class="sxs-lookup"><span data-stu-id="87729-p101">Before sending a message or meeting request item, Exchange Server parses the contents of the item, identifies and stamps certain strings in the subject and body that resemble entities well-known to Exchange, for example, email addresses, phone numbers, and URLs. Messages and meeting requests are delivered by Exchange Server in an Outlook Inbox with well-known entities stamped.</span></span> 

<span data-ttu-id="87729-106">使用 Office JavaScript API，可以获取与特定已知实体匹配的这些字符串以进行进一步处理。</span><span class="sxs-lookup"><span data-stu-id="87729-106">Using the Office JavaScript API, you can get these strings that match specific well-known entities for further processing.</span></span> <span data-ttu-id="87729-107">还可以在外接程序清单中的某个规则中指定已知实体，以便当用户查看某个包含该实体匹配项的项目时，Outlook 可以激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-107">You can also specify a well-known entity in a rule in the add-in manifest so that Outlook can activate your add-in when the user is viewing an item that contains matches for that entity.</span></span> <span data-ttu-id="87729-108">然后您可以提取实体匹配项并对其执行操作。</span><span class="sxs-lookup"><span data-stu-id="87729-108">You can then extract and take action on matches for the entity.</span></span> 

<span data-ttu-id="87729-109">能够识别或从所选的邮件或约会中提取此类实例是很方便的。</span><span class="sxs-lookup"><span data-stu-id="87729-109">Being able to identify or extract such instances from a selected message or appointment is convenient.</span></span> <span data-ttu-id="87729-110">例如，可以构建一个反向电话查找服务，作为 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-110">For example, you can build a reverse phone look-up service as an Outlook add-in.</span></span> <span data-ttu-id="87729-111">该外接程序可从项目主题或正文中提取类似于电话号码的字符串，执行反向搜索并显示每个电话号码的注册所有者。</span><span class="sxs-lookup"><span data-stu-id="87729-111">The add-in can extract strings in the item subject or body that resemble a phone number, do a reverse lookup, and display the registered owner of each phone number.</span></span>

<span data-ttu-id="87729-112">本主题将介绍这些已知实体，显示基于已知实体的激活规则示例，以及如何独立使用激活规则中的实体提取实体匹配项。</span><span class="sxs-lookup"><span data-stu-id="87729-112">This topic introduces these well-known entities, shows examples of activation rules based on well-known entities, and how to extract entity matches independently of having used entities in activation rules.</span></span>


## <a name="support-for-well-known-entities"></a><span data-ttu-id="87729-113">支持已知实体</span><span class="sxs-lookup"><span data-stu-id="87729-113">Support for well-known entities</span></span>

<span data-ttu-id="87729-p104">在发件人发送项目之后和 Exchange 将项目传递给收件人之前，Exchange Server 将标记邮件或会议请求项目中的已知实体。因此，只标记在 Exchange 中传输的项目，用户查看此类项目时，Outlook 可以根据这些标记激活外接程序。反之，用户撰写项目或查看“已发送邮件”文件夹中的项目时，由于项目还没有进行传输，Outlook 无法根据已知实体激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-p104">Exchange Server stamps well-known entities in a message or meeting request item after the sender sends the item and before Exchange delivers the item to the recipient. Therefore, only items that have gone through transport in Exchange are stamped, and Outlook can activate add-ins based on these stamps when the user is viewing such items. On the contrary, when the user is composing an item or viewing an item that is in the Sent Items folder, because the item has not gone through transport, Outlook cannot activate add-ins based on well-known entities.</span></span> 

<span data-ttu-id="87729-p105">同样，无法提取正在撰写的项目中和“已发送邮件”文件夹中的已知实体，因为这些项目尚未进行传输和标记。有关支持激活的项目类型的其他信息，请参阅 [Outlook 外接程序的激活规则](activation-rules.md)。</span><span class="sxs-lookup"><span data-stu-id="87729-p105">Similarly, you cannot extract well-known entities in items that are being composed or in the Sent Items folder, as these items have not gone through transport and are not stamped. For additional information about the kinds of items that support activation, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="87729-p106">下表列出 Exchange Server 和 Outlook 支持和识别的实体（因而称作"已知实体"）和每个实体实例的对象类型。将字符串作为某一实体的自然语言识别基于某学习模型，该模型根据大量数据进行训练。因此，该识别具有不确定性。请参阅 [使用已知实体的提示](#tips-for-using-well-known-entities)来了解有关识别条件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="87729-p106">The following table lists the entities that Exchange Server and Outlook support and recognize (hence the name "well-known entities"), and the object type of an instance of each entity. The natural language recognition of a string as one of these entities is based on a learning model that has been trained on a large amount of data. Therefore, the recognition is non-deterministic. See [Tips for using well-known entities](#tips-for-using-well-known-entities) for more information about conditions for recognition.</span></span>

<span data-ttu-id="87729-123">**表 1.受支持的实体及其类型**</span><span class="sxs-lookup"><span data-stu-id="87729-123">**Table 1. Supported entities and their types**</span></span>

|<span data-ttu-id="87729-124">实体类型</span><span class="sxs-lookup"><span data-stu-id="87729-124">Entity type</span></span>|<span data-ttu-id="87729-125">识别条件</span><span class="sxs-lookup"><span data-stu-id="87729-125">Conditions for recognition</span></span>|<span data-ttu-id="87729-126">对象类型</span><span class="sxs-lookup"><span data-stu-id="87729-126">Object type</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="87729-127">**地址**</span><span class="sxs-lookup"><span data-stu-id="87729-127">**Address**</span></span>|<span data-ttu-id="87729-p107">美国街道地址；例如：1234 Main Street, Redmond, WA 07722。通常，对于要识别的地址，它应遵循美国邮政地址的结构，包含街道编号、街道名称、城市、州和邮政编码等大部分元素。可在一行或多行中指定地址。</span><span class="sxs-lookup"><span data-stu-id="87729-p107">United States street addresses; for example: 1234 Main Street, Redmond, WA 07722. Generally, for an address to be recognized, it should follow the structure of a United States postal address, with most of the elements of a street number, street name, city, state, and zip code present. The address can be specified in one or multiple lines.</span></span>|<span data-ttu-id="87729-131">JavaScript **String** 对象</span><span class="sxs-lookup"><span data-stu-id="87729-131">JavaScript **String** object</span></span>|
|<span data-ttu-id="87729-132">**Contact**</span><span class="sxs-lookup"><span data-stu-id="87729-132">**Contact**</span></span>|<span data-ttu-id="87729-133">对于在自然语言中识别的个人信息的引用。</span><span class="sxs-lookup"><span data-stu-id="87729-133">A reference to a person's information as recognized in natural language.</span></span> <span data-ttu-id="87729-134">联系人的识别取决于上下文。</span><span class="sxs-lookup"><span data-stu-id="87729-134">The recognition of a contact depends on the context.</span></span> <span data-ttu-id="87729-135">例如，邮件末尾的签名或在以下信息附近出现的人员姓名：电话号码、地址、电子邮件地址和 URL。</span><span class="sxs-lookup"><span data-stu-id="87729-135">For example, a signature at the end of a message, or a person's name appearing in the vicinity of some of the following information: a phone number, address, email address, and URL.</span></span>|<span data-ttu-id="87729-136">[Contact](/javascript/api/outlook/office.contact) 对象</span><span class="sxs-lookup"><span data-stu-id="87729-136">[Contact](/javascript/api/outlook/office.contact) object</span></span>|
|<span data-ttu-id="87729-137">**EmailAddress**</span><span class="sxs-lookup"><span data-stu-id="87729-137">**EmailAddress**</span></span>|<span data-ttu-id="87729-138">SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="87729-138">SMTP email addresses.</span></span>|<span data-ttu-id="87729-139">JavaScript `String`对象</span><span class="sxs-lookup"><span data-stu-id="87729-139">JavaScript `String` object</span></span>|
|<span data-ttu-id="87729-140">**MeetingSuggestion**</span><span class="sxs-lookup"><span data-stu-id="87729-140">**MeetingSuggestion**</span></span>|<span data-ttu-id="87729-p109">对事件或会议的引用。例如，Exchange 2013 会将以下文本识别为会面建议： _我们明天一起吃午饭吧。_</span><span class="sxs-lookup"><span data-stu-id="87729-p109">A reference to an event or meeting. For example, Exchange 2013 would recognize the following text as a meeting suggestion:  _Let's meet tomorrow for lunch._</span></span>|<span data-ttu-id="87729-143">[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) 对象</span><span class="sxs-lookup"><span data-stu-id="87729-143">[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object</span></span>|
|<span data-ttu-id="87729-144">**PhoneNumber**</span><span class="sxs-lookup"><span data-stu-id="87729-144">**PhoneNumber**</span></span>|<span data-ttu-id="87729-145">美国电话号码；例如：_(235) 555-0110_</span><span class="sxs-lookup"><span data-stu-id="87729-145">United States telephone numbers; for example:  _(235) 555-0110_</span></span>|<span data-ttu-id="87729-146">[PhoneNumber](/javascript/api/outlook/office.phonenumber) 对象</span><span class="sxs-lookup"><span data-stu-id="87729-146">[PhoneNumber](/javascript/api/outlook/office.phonenumber) object</span></span>|
|<span data-ttu-id="87729-147">**TaskSuggestion**</span><span class="sxs-lookup"><span data-stu-id="87729-147">**TaskSuggestion**</span></span>|<span data-ttu-id="87729-p110">电子邮件中的可操作语句。例如：_请更新电子表格。_</span><span class="sxs-lookup"><span data-stu-id="87729-p110">Actionable sentences in an email. For example:  _Please update the spreadsheet._</span></span>|<span data-ttu-id="87729-150">[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) 对象</span><span class="sxs-lookup"><span data-stu-id="87729-150">[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object</span></span>|
|<span data-ttu-id="87729-151">**Url**</span><span class="sxs-lookup"><span data-stu-id="87729-151">**Url**</span></span>|<span data-ttu-id="87729-152">明确指定了 Web 资源的网络位置和标识符的 Web 地址。</span><span class="sxs-lookup"><span data-stu-id="87729-152">A web address that explicitly specifies the network location and identifier for a web resource.</span></span> <span data-ttu-id="87729-153">Exchange Server 不需要 web 地址中的访问协议，也不识别作为`Url`实体实例嵌入到链接文本中的 url。</span><span class="sxs-lookup"><span data-stu-id="87729-153">Exchange Server does not require the access protocol in the web address, and does not recognize URLs that are embedded in link text as instances of the `Url` entity.</span></span> <span data-ttu-id="87729-154">Exchange Server 可以匹配以下示例： `www.youtube.com/user/officevideos``https://www.youtube.com/user/officevideos`</span><span class="sxs-lookup"><span data-stu-id="87729-154">Exchange Server can match the following examples: `www.youtube.com/user/officevideos` `https://www.youtube.com/user/officevideos`</span></span> |<span data-ttu-id="87729-155">JavaScript `String`对象</span><span class="sxs-lookup"><span data-stu-id="87729-155">JavaScript `String` object</span></span>|

<br/>

<span data-ttu-id="87729-p112">下图说明了 Exchange Server 和 Outlook 如何支持加载项的已知实体，以及哪些加载项可以使用已知实体。请参阅[在加载项中检索实体](#retrieving-entities-in-your-add-in)和[根据实体的存在情况激活加载项](#activating-an-add-in-based-on-the-existence-of-an-entity)，详细了解如何使用这些实体。</span><span class="sxs-lookup"><span data-stu-id="87729-p112">The following figure describes how Exchange Server and Outlook support well-known entities for add-ins, and what add-ins can do with well-known entities. See [Retrieving entities in your add-in](#retrieving-entities-in-your-add-in) and [Activating an add-in based on the existence of an entity](#activating-an-add-in-based-on-the-existence-of-an-entity) for more details on how to use these entities.</span></span>

<span data-ttu-id="87729-158">**Exchange Server、Outlook 和加载项如何支持已知实体**</span><span class="sxs-lookup"><span data-stu-id="87729-158">**How Exchange Server, Outlook, and add-ins support well-known entities**</span></span>

![邮件应用程序中已知实体的支持和使用](../images/well-known-entities-info.png)


## <a name="permissions-to-extract-entities"></a><span data-ttu-id="87729-160">提取实体的权限</span><span class="sxs-lookup"><span data-stu-id="87729-160">Permissions to extract entities</span></span>

<span data-ttu-id="87729-161">若要提取 JavaScript 代码中的实体，或根据特定已知实体的存在情况激活外接程序，请确保已在外接程序清单中请求了相应的权限。</span><span class="sxs-lookup"><span data-stu-id="87729-161">To extract entities in your JavaScript code or to have your add-in activated based on the existence of certain well-known entities, make sure you have requested the appropriate permissions in the add-in manifest.</span></span>

<span data-ttu-id="87729-162">通过指定默认的受限权限，你的外接程序可以`Address`提取`MeetingSuggestion`、或`TaskSuggestion`实体。</span><span class="sxs-lookup"><span data-stu-id="87729-162">Specifying the default restricted permission allows your add-in to extract the `Address`, `MeetingSuggestion`, or `TaskSuggestion` entity.</span></span> <span data-ttu-id="87729-163">若要提取任何其他实体，请指定读取项目、读/写项目或读/写邮箱权限。</span><span class="sxs-lookup"><span data-stu-id="87729-163">To extract any of the other entities, specify read item, read/write item, or read/write mailbox permission.</span></span> <span data-ttu-id="87729-164">若要在清单中执行该操作，请使用 [Permissions](../reference/manifest/permissions.md) 元素并指定适当的权限&mdash;**Restricted**、**ReadItem**、**ReadWriteItem** 或 **ReadWriteMailbox**&mdash;如下例所示：</span><span class="sxs-lookup"><span data-stu-id="87729-164">To do that in the manifest, use the [Permissions](../reference/manifest/permissions.md) element and specify the appropriate permission&mdash;**Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**&mdash;as in the following example:</span></span>

```xml
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a><span data-ttu-id="87729-165">在外接程序中检索实体</span><span class="sxs-lookup"><span data-stu-id="87729-165">Retrieving entities in your add-in</span></span>

<span data-ttu-id="87729-166">只要用户正在查看的项目的主题或正文包含 Exchange 和 Outlook 可以识别为已知实体的字符串，这些实例就可用于加载项。即使外接程序不是基于已知实体激活，它们也是可用的。</span><span class="sxs-lookup"><span data-stu-id="87729-166">As long as the subject or body of the item that is being viewed by the user contains strings that Exchange and Outlook can recognize as well-known entities, these instances are available to add-ins. They are available even if an add-in is not activated based on well-known entities.</span></span> <span data-ttu-id="87729-167">使用适当的权限，可以使用`getEntities` or `getEntitiesByType`方法检索当前邮件或约会中存在的已知实体。</span><span class="sxs-lookup"><span data-stu-id="87729-167">With the appropriate permission, you can use the `getEntities` or `getEntitiesByType` method to retrieve well-known entities that are present in the current message or appointment.</span></span>

<span data-ttu-id="87729-168">`getEntities`方法返回[实体](/javascript/api/outlook/office.entities)对象的数组，其中包含项目中的所有已知实体。</span><span class="sxs-lookup"><span data-stu-id="87729-168">The `getEntities` method returns an array of [Entities](/javascript/api/outlook/office.entities) objects that contains all the well-known entities in the item.</span></span>

<span data-ttu-id="87729-169">如果你对特定类型的实体感兴趣，请使用仅`getEntitiesByType`返回所需实体的数组的方法。</span><span class="sxs-lookup"><span data-stu-id="87729-169">If you're interested in a particular type of entities, use the `getEntitiesByType`method which returns an array of only the entities you want.</span></span> <span data-ttu-id="87729-170">[EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 枚举表示可以提取的所有已知实体类型。</span><span class="sxs-lookup"><span data-stu-id="87729-170">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) enumeration represents all the types of well-known entities you can extract.</span></span>

<span data-ttu-id="87729-171">调用`getEntities`之后，可以使用`Entities`对象的相应属性获取实体类型的实例数组。</span><span class="sxs-lookup"><span data-stu-id="87729-171">After calling `getEntities`, you can then use the corresponding property of the `Entities` object to obtain an array of instances of a type of entity.</span></span> <span data-ttu-id="87729-172">根据实体的类型，数组中的实例可以只是字符串，也可以映射到特定对象。</span><span class="sxs-lookup"><span data-stu-id="87729-172">Depending on the type of entity, the instances in the array can be just strings, or can map to specific objects.</span></span> 

<span data-ttu-id="87729-173">作为前面的图中的示例，若要获取该项目中的地址，请访问由 `getEntities().addresses[]` 返回的数组。</span><span class="sxs-lookup"><span data-stu-id="87729-173">As an example seen in the earlier figure, to get addresses in the item, access the array returned by `getEntities().addresses[]`.</span></span> <span data-ttu-id="87729-174">该`Entities.addresses`属性返回 Outlook 识别为邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="87729-174">The `Entities.addresses` property returns an array of strings that Outlook recognizes as postal addresses.</span></span> <span data-ttu-id="87729-175">同样，该`Entities.contacts`属性返回 Outlook 识别为`Contact`联系人信息的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="87729-175">Similarly, the `Entities.contacts` property returns an array of `Contact` objects that Outlook recognizes as contact information.</span></span> <span data-ttu-id="87729-176">表 1 列出了每个受支持实体的实例的对象类型。</span><span class="sxs-lookup"><span data-stu-id="87729-176">Tables 1 lists the object type of an instance of each supported entity.</span></span>

<span data-ttu-id="87729-177">以下示例显示如何检索在邮件中发现的任何地址。</span><span class="sxs-lookup"><span data-stu-id="87729-177">The following example shows how to retrieve any addresses found in a message.</span></span>

```js
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a><span data-ttu-id="87729-178">根据实体的存在情况激活外接程序</span><span class="sxs-lookup"><span data-stu-id="87729-178">Activating an add-in based on the existence of an entity</span></span>

<span data-ttu-id="87729-179">使用已知实体的另一种方法是，根据当前查看的项目的主题或正文的一个或多个类型的实体的存在情况，使 Outlook 激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-179">Another way to use well-known entities is to have Outlook activate your add-in based on the existence of one or more types of entities in the subject or body of the currently viewed item.</span></span> <span data-ttu-id="87729-180">可以通过在外接程序清单`ItemHasKnownEntity`中指定规则来执行此操作。</span><span class="sxs-lookup"><span data-stu-id="87729-180">You can do so by specifying an `ItemHasKnownEntity` rule in the add-in manifest.</span></span> <span data-ttu-id="87729-181">[EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)简单类型表示`ItemHasKnownEntity`规则支持的常见实体的不同类型。</span><span class="sxs-lookup"><span data-stu-id="87729-181">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) simple type represents the different types of well-known entities supported by `ItemHasKnownEntity` rules.</span></span> <span data-ttu-id="87729-182">激活外接程序后，还可以根据需要检索此类实体的实例，如上一节" [在外接程序中检索实体](#retrieving-entities-in-your-add-in)"中所述。</span><span class="sxs-lookup"><span data-stu-id="87729-182">After your add-in is activated, you can also retrieve the instances of such entities for your purposes, as described in the previous section [Retrieving entities in your add-in](#retrieving-entities-in-your-add-in).</span></span>

<span data-ttu-id="87729-183">您可以选择在`ItemHasKnownEntity`规则中应用正则表达式，以便进一步筛选实体实例，并让 Outlook 仅在实体实例的子集上激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-183">You can optionally apply a regular expression in an `ItemHasKnownEntity` rule, so as to further filter instances of an entity and have Outlook activate an add-in only on a subset of the instances of the entity.</span></span> <span data-ttu-id="87729-184">例如，可为邮件中包含以"98"开头的华盛顿州邮政编码的街道地址实体指定筛选器。</span><span class="sxs-lookup"><span data-stu-id="87729-184">For example, you can specify a filter for the street address entity in a message that contains a Washington state zip code beginning with "98".</span></span> <span data-ttu-id="87729-185">若要对实体实例应用筛选器，请使用`RegExFilter` ItemHasKnownEntity `FilterName`类型的`Rule`元素中的和[](../reference/manifest/rule.md#itemhasknownentity-rule)属性。</span><span class="sxs-lookup"><span data-stu-id="87729-185">To apply a filter on the entity instances, use the `RegExFilter` and `FilterName` attributes in the `Rule` element of the [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) type.</span></span>

<span data-ttu-id="87729-186">类似于其他激活规则，您可以指定多个规则，为外接程序形成一个规则集合。</span><span class="sxs-lookup"><span data-stu-id="87729-186">Similar to other activation rules, you can specify multiple rules to form a rule collection for your add-in.</span></span> <span data-ttu-id="87729-187">下面的示例对2个规则应用 "AND" 操作：一个`ItemIs`规则和一个`ItemHasKnownEntity`规则。</span><span class="sxs-lookup"><span data-stu-id="87729-187">The following example applies an "AND" operation on 2 rules: an `ItemIs` rule and an `ItemHasKnownEntity` rule.</span></span> <span data-ttu-id="87729-188">只要当前项目为邮件，且 Outlook 识别该项目主题或正文中的地址时，此规则集合就将激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-188">This rule collection activates the add-in whenever the current item is a message and Outlook recognizes an address in the subject or body of that item.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<br/>

<span data-ttu-id="87729-189">下面的示例使用`getEntitiesByType`当前项将一个变量`addresses`设置为前一个规则集合的结果。</span><span class="sxs-lookup"><span data-stu-id="87729-189">The following example uses `getEntitiesByType` of the current item to set a variable `addresses` to the results of the preceding rule collection.</span></span>

```js
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

<br/>

<span data-ttu-id="87729-190">如果当前`ItemHasKnownEntity`项目的主题或正文中有一个 url，并且该 url 包含字符串 "youtube"，则以下规则示例将激活加载项，而不考虑字符串的大小写。</span><span class="sxs-lookup"><span data-stu-id="87729-190">The following `ItemHasKnownEntity` rule example activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string "youtube", regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

<br/>

<span data-ttu-id="87729-191">下面的示例使用`getFilteredEntitiesByName(name)`当前项设置一个变量`videos` ，以获取与前面`ItemHasKnownEntity`规则中的正则表达式匹配的结果数组。</span><span class="sxs-lookup"><span data-stu-id="87729-191">The following example uses `getFilteredEntitiesByName(name)` of the current item to set a variable `videos` to get an array of results that match the regular expression in the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a><span data-ttu-id="87729-192">使用已知实体的提示</span><span class="sxs-lookup"><span data-stu-id="87729-192">Tips for using well-known entities</span></span>

<span data-ttu-id="87729-193">在外接程序中使用已知实体时，应了解一些事实和限制。</span><span class="sxs-lookup"><span data-stu-id="87729-193">There are a few facts and limits you should be aware of if you use well-known entities in your add-in.</span></span> <span data-ttu-id="87729-194">只要用户读取的项包含已知实体的匹配项，并且无论您是否使用`ItemHasKnownEntity`规则，都将在下面应用以下内容：</span><span class="sxs-lookup"><span data-stu-id="87729-194">The following applies as long as your add-in is activated when the user is reading an item which contains matches of well-known entities, regardless of whether you use an `ItemHasKnownEntity` rule:</span></span>


- <span data-ttu-id="87729-195">仅当字符串为英文形式时，您才可以提取已知实体字符串。</span><span class="sxs-lookup"><span data-stu-id="87729-195">You can extract strings that are well-known entities only if the strings are in English.</span></span>
    
- <span data-ttu-id="87729-196">您可以从项目正文的前 2,000 个字符中提取已知实体，但不能超过此限制。</span><span class="sxs-lookup"><span data-stu-id="87729-196">You can extract well-known entities from the first 2,000 characters in the item body, but not beyond that limit.</span></span> <span data-ttu-id="87729-197">此大小限制有助于平衡功能和性能之间的需求，因此 Exchange Server 和 Outlook 不会因分析和确定大型邮件和约会中的已知实体实例而停滞。</span><span class="sxs-lookup"><span data-stu-id="87729-197">This size limit helps balance the need for functionality and performance, so that Exchange Server and Outlook are not bogged down by parsing and identifying instances of well-known entities in large messages and appointments.</span></span> <span data-ttu-id="87729-198">请注意，此限制与外接程序是否指定一个`ItemHasKnownEntity`规则无关。</span><span class="sxs-lookup"><span data-stu-id="87729-198">Note that this limit is independent of whether the add-in specifies an `ItemHasKnownEntity` rule.</span></span> <span data-ttu-id="87729-199">如果外接程序使用此类规则，还要注意以下项目 2 中针对 Outlook 富客户端的的规则处理限制。</span><span class="sxs-lookup"><span data-stu-id="87729-199">If the add-in does use such a rule, note also the rule processing limit in item 2 below for the Outlook rich clients.</span></span>
    
- <span data-ttu-id="87729-p123">您可以从约会（由邮箱所有者之外的人员组织的会议）中提取实体。如果日历项目不是会议或不是由邮箱所有者组织的会议，则不能从中提取实体。</span><span class="sxs-lookup"><span data-stu-id="87729-p123">You can extract entities from appointments that are meetings organized by someone other than the mailbox owner. You cannot extract entities from calendar items that are not meetings, or meetings organized by the mailbox owner.</span></span>
    
- <span data-ttu-id="87729-202">您可以仅从邮件（ `MeetingSuggestion`而非约会）中提取类型的实体。</span><span class="sxs-lookup"><span data-stu-id="87729-202">You can extract entities of the `MeetingSuggestion` type from only messages but not appointments.</span></span>
    
- <span data-ttu-id="87729-203">您可以提取项目正文中明确存在的 URL，但无法提取 HTML 项目正文中内嵌在超链接文本中的 URL。</span><span class="sxs-lookup"><span data-stu-id="87729-203">You can extract URLs that exist explicitly in the item body, but not URLs that are embedded in hyperlinked text in HTML item body.</span></span> <span data-ttu-id="87729-204">请考虑改用`ItemHasRegularExpressionMatch`规则来获取显式和嵌入的 url。</span><span class="sxs-lookup"><span data-stu-id="87729-204">Consider using an `ItemHasRegularExpressionMatch` rule instead to get both explicit and embedded URLs.</span></span> <span data-ttu-id="87729-205">指定`BodyAsHTML`为_PropertyName_，以及将 url 匹配为_RegExValue_的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="87729-205">Specify `BodyAsHTML` as the _PropertyName_, and a regular expression that matches URLs as the  _RegExValue_.</span></span>
    
- <span data-ttu-id="87729-206">不能从"已发送邮件"文件夹中的邮件提取实体。</span><span class="sxs-lookup"><span data-stu-id="87729-206">You cannot extract entities from items in the Sent Items folder.</span></span>
    
<span data-ttu-id="87729-207">此外，如果使用 [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 规则，并可能影响您希望激活外接程序的方案，则适用于以下情况：</span><span class="sxs-lookup"><span data-stu-id="87729-207">In addition, the following applies if you use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule, and may affect the scenarios where you'd otherwise expect your add-in to be activated:</span></span>

- <span data-ttu-id="87729-208">使用`ItemHasKnownEntity`规则时，预期 Outlook 仅匹配英文实体字符串，而不考虑清单中指定的默认区域设置。</span><span class="sxs-lookup"><span data-stu-id="87729-208">When using the `ItemHasKnownEntity` rule, expect Outlook to match entity strings in only English regardless of the default locale specified in the manifest.</span></span>
    
- <span data-ttu-id="87729-209">当您的外接程序在 Outlook 富客户端上运行时，预期 Outlook 会`ItemHasKnownEntity`将该规则应用于项目正文的第一个 mb，而不是在该限制范围内的其余正文中。</span><span class="sxs-lookup"><span data-stu-id="87729-209">When your add-in is running on an Outlook rich client, expect Outlook to apply the `ItemHasKnownEntity` rule to the first megabyte of the item body and not to the rest of the body over that limit.</span></span>
    
- <span data-ttu-id="87729-210">您不能使用`ItemHasKnownEntity`规则为 "已发送邮件" 文件夹中的项目激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="87729-210">You cannot use an `ItemHasKnownEntity` rule to activate an add-in for items in the Sent Items folder.</span></span>
    

## <a name="see-also"></a><span data-ttu-id="87729-211">另请参阅</span><span class="sxs-lookup"><span data-stu-id="87729-211">See also</span></span>

- [<span data-ttu-id="87729-212">创建适用于阅读窗体的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="87729-212">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="87729-213">从 Outlook 项目中提取实体字符串</span><span class="sxs-lookup"><span data-stu-id="87729-213">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)
- [<span data-ttu-id="87729-214">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="87729-214">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="87729-215">使用正则表达式激活规则显示 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="87729-215">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="87729-216">了解 Outlook 外接程序权限</span><span class="sxs-lookup"><span data-stu-id="87729-216">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
