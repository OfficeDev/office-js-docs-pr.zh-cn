---
title: Outlook 加载项的激活规则
description: 如果用户正在读取或撰写的邮件或约会符合加载项的激活规则，则 Outlook 将激活某些类型的加载项。
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: 24f17b7bb3da4665f3f05b23d34ba15bcc4ae729
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349019"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="9016b-103">上下文 Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="9016b-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="9016b-p101">如果用户正在读取或撰写的邮件或约会符合外接程序的激活规则，则 Outlook 将激活某些类型的外接程序。这一点对使用 1.1 清单架构的所有外接程序均适用。然后，用户可从 Outlook UI 选择外接程序，以开始将其用于当前项目。</span><span class="sxs-lookup"><span data-stu-id="9016b-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="9016b-107">下图显示在“阅读”窗格中的邮件的外接程序栏中激活的 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![显示已激活阅读邮件应用程序的应用程序栏。](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="9016b-109">在清单中指定激活规则</span><span class="sxs-lookup"><span data-stu-id="9016b-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="9016b-110">若要Outlook条件激活外接程序，请通过使用下列元素之一在外接程序清单中指定激活 `Rule` 规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements.</span></span>

- <span data-ttu-id="9016b-111">[Rule 元素 (MailApp complexType)](../reference/manifest/rule.md) - 指定单个规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="9016b-112">[Rule 元素 (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - 使用逻辑操作组合多个规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>


 > [!NOTE]
 > <span data-ttu-id="9016b-113">`Rule`用于指定单个规则的元素是抽象[Rule](../reference/manifest/rule.md)复杂类型。</span><span class="sxs-lookup"><span data-stu-id="9016b-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="9016b-114">以下每种类型的规则扩展了此抽象 `Rule` 复杂类型。</span><span class="sxs-lookup"><span data-stu-id="9016b-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="9016b-115">因此当你在清单中指定单个规则时，你必须使用 [xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性来进一步定义某个以下类型的规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="9016b-116">例如，以下规则定义 [ItemIs](../reference/manifest/rule.md#itemis-rule) 规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule.</span></span>
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > <span data-ttu-id="9016b-117">`FormType`属性适用于清单 v1.1 中的激活规则，但在 `VersionOverrides` v1.0 中未定义。</span><span class="sxs-lookup"><span data-stu-id="9016b-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="9016b-118">因此，当在节点中使用 [ItemIs](../reference/manifest/rule.md#itemis-rule) 时，它不能 `VersionOverrides` 使用。</span><span class="sxs-lookup"><span data-stu-id="9016b-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="9016b-p105">下表列出了可用的规则类型。你可以在表后面以及[创建适用于阅读窗体的 Outlook 外接程序](read-scenario.md)中指定的文章中查找更多信息。</span><span class="sxs-lookup"><span data-stu-id="9016b-p105">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="9016b-121">**规则名称**</span><span class="sxs-lookup"><span data-stu-id="9016b-121">**Rule name**</span></span>|<span data-ttu-id="9016b-122">**适用的窗体**</span><span class="sxs-lookup"><span data-stu-id="9016b-122">**Applicable forms**</span></span>|<span data-ttu-id="9016b-123">**说明**</span><span class="sxs-lookup"><span data-stu-id="9016b-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="9016b-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="9016b-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="9016b-125">读取，撰写</span><span class="sxs-lookup"><span data-stu-id="9016b-125">Read, Compose</span></span>|<span data-ttu-id="9016b-p106">检查当前项目是否属于指定类型（邮件或约会），另外还可以检查项目类别、窗体类型和（可选）项目邮件类别。</span><span class="sxs-lookup"><span data-stu-id="9016b-p106">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="9016b-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="9016b-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="9016b-129">读取</span><span class="sxs-lookup"><span data-stu-id="9016b-129">Read</span></span>|<span data-ttu-id="9016b-130">检查所选项是否包含附件。</span><span class="sxs-lookup"><span data-stu-id="9016b-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="9016b-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="9016b-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="9016b-132">读取</span><span class="sxs-lookup"><span data-stu-id="9016b-132">Read</span></span>|<span data-ttu-id="9016b-p107">检查所选项是否包含一个或多个已知实体。更多信息：[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。</span><span class="sxs-lookup"><span data-stu-id="9016b-p107">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="9016b-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="9016b-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="9016b-136">读取</span><span class="sxs-lookup"><span data-stu-id="9016b-136">Read</span></span>|<span data-ttu-id="9016b-137">检查发件人的电子邮件地址、所选项的主题和/或所选项的正文是否包含正则表达式的匹配项。更多信息： [使用正则表达式激活规则显示 Outlook 外接程序](use-regular-expressions-to-show-an-outlook-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="9016b-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="9016b-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="9016b-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="9016b-139">读取，撰写</span><span class="sxs-lookup"><span data-stu-id="9016b-139">Read, Compose</span></span>|<span data-ttu-id="9016b-140">组合一组规则以便形成更复杂的规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="9016b-141">ItemIs 规则</span><span class="sxs-lookup"><span data-stu-id="9016b-141">ItemIs rule</span></span>

<span data-ttu-id="9016b-142">**ItemIs** 复杂类型定义一个计算结果为 **true** 的规则（如果当前项与项类型匹配）和（可选）项邮件类别（如果在规则中指明）。</span><span class="sxs-lookup"><span data-stu-id="9016b-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="9016b-143">在 ItemIs 规则的 属性中指定 `ItemType` 以下 **项目类型之** 一。</span><span class="sxs-lookup"><span data-stu-id="9016b-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="9016b-144">可以在清单中指定多个 **ItemIs** 规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="9016b-145">ItemType simpleType 定义了支持 Outlook 加载项的 Outlook 项类型。</span><span class="sxs-lookup"><span data-stu-id="9016b-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="9016b-146">**Value**</span><span class="sxs-lookup"><span data-stu-id="9016b-146">**Value**</span></span>|<span data-ttu-id="9016b-147">**说明**</span><span class="sxs-lookup"><span data-stu-id="9016b-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9016b-148">**约会**</span><span class="sxs-lookup"><span data-stu-id="9016b-148">**Appointment**</span></span>|<span data-ttu-id="9016b-149">在 Outlook 日历中指定一个项目。</span><span class="sxs-lookup"><span data-stu-id="9016b-149">Specifies an item in an Outlook calendar.</span></span> <span data-ttu-id="9016b-150">这包括已获取响应并且具有组织者和参与者的会议项目，或者没有组织者或参与者且仅为日历上的一个项目的约会。</span><span class="sxs-lookup"><span data-stu-id="9016b-150">This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.</span></span> <span data-ttu-id="9016b-151">这与 Outlook 中的 IPM.Appointment 邮件类别相对应。</span><span class="sxs-lookup"><span data-stu-id="9016b-151">This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="9016b-152">**消息**</span><span class="sxs-lookup"><span data-stu-id="9016b-152">**Message**</span></span>|<span data-ttu-id="9016b-153">指定通常在收件箱中收到的以下项目之一。</span><span class="sxs-lookup"><span data-stu-id="9016b-153">Specifies one of the following items received in typically the Inbox.</span></span> <ul><li><p><span data-ttu-id="9016b-p110">电子邮件。这与 Outlook 中的 IPM.Note 邮件类别相对应。</span><span class="sxs-lookup"><span data-stu-id="9016b-p110">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="9016b-156">会议请求、响应或取消。</span><span class="sxs-lookup"><span data-stu-id="9016b-156">A meeting request, response, or cancellation.</span></span> <span data-ttu-id="9016b-157">这对应于 Outlook 中的以下邮件Outlook。</span><span class="sxs-lookup"><span data-stu-id="9016b-157">This corresponds to the following message classes in Outlook.</span></span></p><p><span data-ttu-id="9016b-158">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="9016b-158">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="9016b-159">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="9016b-159">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="9016b-160">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="9016b-160">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="9016b-161">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="9016b-161">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="9016b-162">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="9016b-162">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="9016b-163">`FormType`属性用于指定阅读 (撰写) 外接程序应激活的模式。</span><span class="sxs-lookup"><span data-stu-id="9016b-163">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="9016b-164">ItemIs `FormType` 属性在架构 v1.1 及更高版本中定义，但不在 `VersionOverrides` v1.0 中定义。</span><span class="sxs-lookup"><span data-stu-id="9016b-164">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="9016b-165">定义外接程序 `FormType` 命令时，请勿包含 属性。</span><span class="sxs-lookup"><span data-stu-id="9016b-165">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="9016b-166">激活外接程序后，可以使用 [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 属性获取 Outlook 中的当前所选项，以及使用 [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性获取当前项的类型。</span><span class="sxs-lookup"><span data-stu-id="9016b-166">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="9016b-167">可以选择使用 属性指定项目的邮件类，以及使用 属性指定当项目是指定类的子类时规则是否应该 `ItemClass` `IncludeSubClasses` 为true。</span><span class="sxs-lookup"><span data-stu-id="9016b-167">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="9016b-168">若要详细了解邮件类，请参阅[项类型和邮件类](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)。</span><span class="sxs-lookup"><span data-stu-id="9016b-168">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="9016b-169">下面的示例是 **一个 ItemIs** 规则，它允许用户在阅读邮件时Outlook外接程序栏中查看外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message.</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="9016b-170">下面的示例展示了 **ItemIs** 规则，可便于用户在阅读邮件或约会时在 Outlook 加载项栏中看到加载项。</span><span class="sxs-lookup"><span data-stu-id="9016b-170">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="9016b-171">ItemHasAttachment 规则</span><span class="sxs-lookup"><span data-stu-id="9016b-171">ItemHasAttachment rule</span></span>


<span data-ttu-id="9016b-172">复杂 `ItemHasAttachment` 类型定义一个规则，用于检查所选项目是否包含附件。</span><span class="sxs-lookup"><span data-stu-id="9016b-172">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="9016b-173">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="9016b-173">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="9016b-174">在项对外接程序可用之前，服务器将对其进行检查以确定主题和正文是否包含可能为某个已知实体的任何文本。</span><span class="sxs-lookup"><span data-stu-id="9016b-174">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities.</span></span> <span data-ttu-id="9016b-175">如果找到这些实体中的任意一个，则放置在使用该项的 或 方法访问的 `getEntities` 已知 `getEntitiesByType` 实体集合中。</span><span class="sxs-lookup"><span data-stu-id="9016b-175">If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="9016b-176">您可以使用 在项中出现指定类型的实体时显示外接程序来 `ItemHasKnownEntity` 指定规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-176">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item.</span></span> <span data-ttu-id="9016b-177">可以在规则的 属性中指定以下 `EntityType` 已知 `ItemHasKnownEntity` 实体。</span><span class="sxs-lookup"><span data-stu-id="9016b-177">You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule.</span></span>

- <span data-ttu-id="9016b-178">Address</span><span class="sxs-lookup"><span data-stu-id="9016b-178">Address</span></span>
- <span data-ttu-id="9016b-179">Contact</span><span class="sxs-lookup"><span data-stu-id="9016b-179">Contact</span></span>
- <span data-ttu-id="9016b-180">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="9016b-180">EmailAddress</span></span>
- <span data-ttu-id="9016b-181">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9016b-181">MeetingSuggestion</span></span>
- <span data-ttu-id="9016b-182">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9016b-182">PhoneNumber</span></span>
- <span data-ttu-id="9016b-183">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9016b-183">TaskSuggestion</span></span>
- <span data-ttu-id="9016b-184">URL</span><span class="sxs-lookup"><span data-stu-id="9016b-184">URL</span></span>

<span data-ttu-id="9016b-185">可以选择在 属性中包括正则表达式，以便仅在存在与正则表达式匹配的实体时 `RegularExpression` 显示外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-185">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present.</span></span> <span data-ttu-id="9016b-186">若要获取规则中指定的正则表达式的匹配项，可以将 或 方法用于当前选定的 `ItemHasKnownEntity` `getRegExMatches` `getFilteredEntitiesByName` Outlook项。</span><span class="sxs-lookup"><span data-stu-id="9016b-186">To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="9016b-187">以下示例显示一组元素，这些元素在邮件中出现指定的已知实体之一时 `Rule` 显示外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-187">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="9016b-188">以下示例显示一个包含属性的规则，当邮件中包含单词"contoso"的 URL 时，该规则 `ItemHasKnownEntity` `RegularExpression` 将激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-188">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="9016b-189">有关激活规则中的实体的详细信息，请参阅[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。</span><span class="sxs-lookup"><span data-stu-id="9016b-189">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="9016b-190">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="9016b-190">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="9016b-191">`ItemHasRegularExpressionMatch`复杂类型定义一个规则，该规则使用正则表达式来匹配项目的指定属性的内容。</span><span class="sxs-lookup"><span data-stu-id="9016b-191">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item.</span></span> <span data-ttu-id="9016b-192">如果在项的指定属性中发现与正则表达式匹配的文本，则 Outlook 会激活外接程序栏并显示外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-192">If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in.</span></span> <span data-ttu-id="9016b-193">可以使用表示当前选定项的对象的 或 方法 `getRegExMatches` `getRegExMatchesByName` 获取指定正则表达式的匹配项。</span><span class="sxs-lookup"><span data-stu-id="9016b-193">You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="9016b-194">以下示例显示 ，当选定项的正文包含 `ItemHasRegularExpressionMatch` "apple"、"apple"或"apple"（不区分大小写）时，将激活加载项。</span><span class="sxs-lookup"><span data-stu-id="9016b-194">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="9016b-195">有关使用该规则的信息 `ItemHasRegularExpressionMatch` ，请参阅使用[正则表达式激活规则Outlook外接程序。](use-regular-expressions-to-show-an-outlook-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="9016b-195">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="9016b-196">RuleCollection 规则</span><span class="sxs-lookup"><span data-stu-id="9016b-196">RuleCollection rule</span></span>


<span data-ttu-id="9016b-197">复杂 `RuleCollection` 类型将多个规则合并为一个规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-197">The `RuleCollection` complex type combines multiple rules into a single rule.</span></span> <span data-ttu-id="9016b-198">可以使用 属性指定集合中的规则是应结合逻辑 OR 还是逻辑 `Mode` AND。</span><span class="sxs-lookup"><span data-stu-id="9016b-198">You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="9016b-p118">指定逻辑 AND 时，项必须与集合中的所有指定规则匹配才能显示外接程序。指定逻辑 OR 时，与集合中的任何指定规则匹配的项都将显示外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-p118">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="9016b-201">您可以组合 `RuleCollection` 规则以形成复杂的规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-201">You can combine `RuleCollection` rules to form complex rules.</span></span> <span data-ttu-id="9016b-202">以下示例在用户查看约会或邮件项目（项目的主题或正文包含地址）时激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-202">The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="9016b-203">以下示例在用户撰写邮件时或查看约会（约会的标题或正文包含地址）时激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9016b-203">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="9016b-204">规则和正则表达式的限制</span><span class="sxs-lookup"><span data-stu-id="9016b-204">Limits for rules and regular expressions</span></span>


<span data-ttu-id="9016b-205">为了提供使用 Outlook 外接程序的满意体验，您应该遵守激活和 API 使用准则。</span><span class="sxs-lookup"><span data-stu-id="9016b-205">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines.</span></span> <span data-ttu-id="9016b-206">下表显示了正则表达式和规则的常规限制，但不同的应用程序有特定的规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-206">The following table shows general limits for regular expressions and rules but there are specific rules for different applications.</span></span> <span data-ttu-id="9016b-207">有关详细信息，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)和 [排查 Outlook 外接程序激活问题](troubleshoot-outlook-add-in-activation.md)。</span><span class="sxs-lookup"><span data-stu-id="9016b-207">For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="9016b-208">**外接程序元素**</span><span class="sxs-lookup"><span data-stu-id="9016b-208">**Add-in element**</span></span>|<span data-ttu-id="9016b-209">**准则**</span><span class="sxs-lookup"><span data-stu-id="9016b-209">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9016b-210">清单大小</span><span class="sxs-lookup"><span data-stu-id="9016b-210">Manifest Size</span></span>|<span data-ttu-id="9016b-211">不大于 256 KB。</span><span class="sxs-lookup"><span data-stu-id="9016b-211">No larger than 256 KB.</span></span>|
|<span data-ttu-id="9016b-212">规则</span><span class="sxs-lookup"><span data-stu-id="9016b-212">Rules</span></span>|<span data-ttu-id="9016b-213">不超过 15 条规则。</span><span class="sxs-lookup"><span data-stu-id="9016b-213">No more than 15 rules.</span></span>|
|<span data-ttu-id="9016b-214">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="9016b-214">ItemHasKnownEntity</span></span>|<span data-ttu-id="9016b-215">Outlook 富客户端将对正文的前 1 MB 内容应用规则，对正文其余部分则不应用。</span><span class="sxs-lookup"><span data-stu-id="9016b-215">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="9016b-216">正则表达式</span><span class="sxs-lookup"><span data-stu-id="9016b-216">Regular Expressions</span></span>|<span data-ttu-id="9016b-217">对于所有应用程序的所有 ItemHasKnownEntity 或 ItemHasRegularExpressionMatch Outlook规则：</span><span class="sxs-lookup"><span data-stu-id="9016b-217">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="9016b-p121">在 Outlook 加载项的激活规则中指定不超过 5 个正则表达式。如果超过该限制，则无法安装加载项。</span><span class="sxs-lookup"><span data-stu-id="9016b-p121">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="9016b-220">指定由 <b>getRegExMatches</b> 方法调用在前 50 个匹配项内返回其预期结果的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9016b-220">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="9016b-221">在正则表达式中指定向前断言，但不支持向后 `(?<=text)` 和否定向后 `(?<!text)` 断言。</span><span class="sxs-lookup"><span data-stu-id="9016b-221">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="9016b-222">指定其匹配不超过下表中的限制的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9016b-222">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="9016b-223">正则表达式匹配项的长度限制</span><span class="sxs-lookup"><span data-stu-id="9016b-223">Limit on length of a regex match</span></span></th><th><span data-ttu-id="9016b-224">Outlook 富客户端</span><span class="sxs-lookup"><span data-stu-id="9016b-224">Outlook rich clients</span></span></th><th><span data-ttu-id="9016b-225">iOS 版和 Android 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="9016b-225">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="9016b-226">项目正文采用纯文本</span><span class="sxs-lookup"><span data-stu-id="9016b-226">Item body is plain text</span></span></td><td><span data-ttu-id="9016b-227">1.5 KB</span><span class="sxs-lookup"><span data-stu-id="9016b-227">1.5 KB</span></span></td><td><span data-ttu-id="9016b-228">3 KB</span><span class="sxs-lookup"><span data-stu-id="9016b-228">3 KB</span></span></td></tr><tr><td><span data-ttu-id="9016b-229">项目正文采用 HTML</span><span class="sxs-lookup"><span data-stu-id="9016b-229">Item body it HTML</span></span></td><td><span data-ttu-id="9016b-230">3 KB</span><span class="sxs-lookup"><span data-stu-id="9016b-230">3 KB</span></span></td><td><span data-ttu-id="9016b-231">3KB</span><span class="sxs-lookup"><span data-stu-id="9016b-231">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="9016b-232">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9016b-232">See also</span></span>

- [<span data-ttu-id="9016b-233">创建适用于撰写窗体的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="9016b-233">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="9016b-234">Outlook 加载项的激活限制和 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="9016b-234">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="9016b-235">使用正则表达式激活规则显示 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="9016b-235">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="9016b-236">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="9016b-236">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
