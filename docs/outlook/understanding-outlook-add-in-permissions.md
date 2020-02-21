---
title: 了解 Outlook 加载项权限
description: Outlook 加载项在清单中指定所需的权限级别，其中包括受限、ReadItem、ReadWriteItem 或 ReadWriteMailbox。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 58d21a33034475b8c33b8449ece24c9dafc84e2b
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165880"
---
# <a name="understanding-outlook-add-in-permissions"></a><span data-ttu-id="97e77-103">了解 Outlook 加载项权限</span><span class="sxs-lookup"><span data-stu-id="97e77-103">Understanding Outlook add-in permissions</span></span>

<span data-ttu-id="97e77-p101">Outlook 外接程序在清单中指定所需的权限级别。可用级别为**Restricted**、**ReadItem**、**ReadWriteItem**或**ReadWriteMailbox**。这些权限级别可累计：“**Restricted**”是最低的级别，并且每个更高级别包括所有较低级别的权限。“**ReadWriteMailbox**”包含所有受支持的权限。</span><span class="sxs-lookup"><span data-stu-id="97e77-p101">Outlook add-ins specify the required permission level in their manifest. The available levels are **Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**. These levels of permissions are cumulative: **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.</span></span>

<span data-ttu-id="97e77-p102">在从 [AppSource](https://appsource.microsoft.com) 安装邮件加载项之前，你可以查看该邮件加载项所需的权限。你还可以在 Exchange 管理中心中查看已安装加载项所需的权限。</span><span class="sxs-lookup"><span data-stu-id="97e77-p102">You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.</span></span>

## <a name="restricted-permission"></a><span data-ttu-id="97e77-110">“Restricted”权限</span><span class="sxs-lookup"><span data-stu-id="97e77-110">Restricted permission</span></span>

<span data-ttu-id="97e77-p103">**Restricted**权限是最基本级别的权限。在清单的[权限](../reference/manifest/permissions.md)元素中指定**Restricted**可以请求获取此权限。如果外接程序不请求其清单中的将特定权限，在默认情况下，Outlook 会将此权限分配给邮件外接程序。</span><span class="sxs-lookup"><span data-stu-id="97e77-p103">The **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](../reference/manifest/permissions.md) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.</span></span>

### <a name="can-do"></a><span data-ttu-id="97e77-114">可以执行的操作</span><span class="sxs-lookup"><span data-stu-id="97e77-114">Can do</span></span>

- <span data-ttu-id="97e77-115">[仅获取项目主题或正文的特定实体](match-strings-in-an-item-as-well-known-entities.md)（电话号码、地址、URL）。</span><span class="sxs-lookup"><span data-stu-id="97e77-115">[Get only specific entities](match-strings-in-an-item-as-well-known-entities.md) (phone number, address, URL) from the item's subject or body.</span></span>

- <span data-ttu-id="97e77-116">指定[项目激活规则](activation-rules.md#itemis-rule)，此类规则需要阅读或撰写窗体中的当前项目为特定的项目类型，或与选定项目中支持的已知实体（电话号码、地址、URL）的任何较小子集匹配的 [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) 规则。</span><span class="sxs-lookup"><span data-stu-id="97e77-116">Specify an [ItemIs activation rule](activation-rules.md#itemis-rule) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.</span></span>

- <span data-ttu-id="97e77-117">访问与用户或项目具体信息**无**关的任何属性和方法。（请参阅下一部分，了解与用户或项目具体信息相关的属性和方法列表）。</span><span class="sxs-lookup"><span data-stu-id="97e77-117">Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).</span></span>

### <a name="cant-do"></a><span data-ttu-id="97e77-118">不能执行的操作</span><span class="sxs-lookup"><span data-stu-id="97e77-118">Can't do</span></span>

- <span data-ttu-id="97e77-119">在联系人、电子邮件地址、会议建议或任务建议实体上使用 [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 规则。</span><span class="sxs-lookup"><span data-stu-id="97e77-119">Use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the contact, email address, meeting suggestion, or task suggestion entitiy.</span></span>

- <span data-ttu-id="97e77-120">使用 [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) 或 [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 规则。</span><span class="sxs-lookup"><span data-stu-id="97e77-120">Use the [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule.</span></span>

- <span data-ttu-id="97e77-p104">访问以下列表中与用户或邮件具体信息相关的属性和方法。尝试访问此列表中的成员将返回 **null**，并生成指明 Outlook 要求邮件外接程序具有提升的权限的错误消息。</span><span class="sxs-lookup"><span data-stu-id="97e77-p104">Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.</span></span>

    - [<span data-ttu-id="97e77-123">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-123">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-124">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-124">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-125">item.attachments</span><span class="sxs-lookup"><span data-stu-id="97e77-125">item.attachments</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-126">item.bcc</span><span class="sxs-lookup"><span data-stu-id="97e77-126">item.bcc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-127">item.body</span><span class="sxs-lookup"><span data-stu-id="97e77-127">item.body</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-128">item.cc</span><span class="sxs-lookup"><span data-stu-id="97e77-128">item.cc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-129">item.from</span><span class="sxs-lookup"><span data-stu-id="97e77-129">item.from</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-130">item.getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="97e77-130">item.getRegExMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-131">item.getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="97e77-131">item.getRegExMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-132">item.optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="97e77-132">item.optionalAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-133">item.organizer</span><span class="sxs-lookup"><span data-stu-id="97e77-133">item.organizer</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-134">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-134">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-135">item.requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="97e77-135">item.requiredAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-136">item.sender</span><span class="sxs-lookup"><span data-stu-id="97e77-136">item.sender</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-137">item.to</span><span class="sxs-lookup"><span data-stu-id="97e77-137">item.to</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="97e77-138">mailbox.getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-138">mailbox.getCallbackTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="97e77-139">mailbox.getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-139">mailbox.getUserIdentityTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="97e77-140">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-140">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="97e77-141">mailbox.userProfile</span><span class="sxs-lookup"><span data-stu-id="97e77-141">mailbox.userProfile</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - <span data-ttu-id="97e77-142">[Body](/javascript/api/outlook/office.body) 及其所有子成员</span><span class="sxs-lookup"><span data-stu-id="97e77-142">[Body](/javascript/api/outlook/office.body) and all its child members</span></span>
    - <span data-ttu-id="97e77-143">[Location](/javascript/api/outlook/office.location) 及其所有子成员</span><span class="sxs-lookup"><span data-stu-id="97e77-143">[Location](/javascript/api/outlook/office.location) and all its child members</span></span>
    - <span data-ttu-id="97e77-144">[Recipients](/javascript/api/outlook/office.recipients) 及其所有子成员</span><span class="sxs-lookup"><span data-stu-id="97e77-144">[Recipients](/javascript/api/outlook/office.recipients) and all its child members</span></span>
    - <span data-ttu-id="97e77-145">[Subject](/javascript/api/outlook/office.subject) 及其所有子成员</span><span class="sxs-lookup"><span data-stu-id="97e77-145">[Subject](/javascript/api/outlook/office.subject) and all its child members</span></span>
    - <span data-ttu-id="97e77-146">[Time](/javascript/api/outlook/office.time) 及其所有子成员</span><span class="sxs-lookup"><span data-stu-id="97e77-146">[Time](/javascript/api/outlook/office.time) and all its child members</span></span>

## <a name="readitem-permission"></a><span data-ttu-id="97e77-147">“ReadItem”权限</span><span class="sxs-lookup"><span data-stu-id="97e77-147">ReadItem permission</span></span>

<span data-ttu-id="97e77-p105">**ReadItem**权限是权限模型中的下一级别权限。在清单的“权限”\*\*\*\* 元素中指定“ReadItem”\*\*\*\* 可以请求获取此权限。</span><span class="sxs-lookup"><span data-stu-id="97e77-p105">The **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="97e77-150">可以执行的操作</span><span class="sxs-lookup"><span data-stu-id="97e77-150">Can do</span></span>

- <span data-ttu-id="97e77-151">在读取或 [撰写窗体](item-data.md)[中读取当前项目的所有属性](get-and-set-item-data-in-a-compose-form.md)，例如阅读窗体中的 [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和撰写窗体中的 [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="97e77-151">[Read all the properties](item-data.md) of the current item in a read or [compose form](get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) in a compose form.</span></span>

- <span data-ttu-id="97e77-152">[获取回调令牌](get-attachments-of-an-outlook-item.md)，以使用 Exchange Web 服务 (EWS) 或 [Outlook REST API](use-rest-api.md) 获取邮件附件或整个邮件。</span><span class="sxs-lookup"><span data-stu-id="97e77-152">[Get a callback token to get item attachments](get-attachments-of-an-outlook-item.md) or the full item with Exchange Web Services (EWS) or [Outlook REST APIs](use-rest-api.md).</span></span>

- <span data-ttu-id="97e77-153">[编写外接程序在相应邮件上设置的自定义属性](/javascript/api/outlook/office.CustomProperties)。</span><span class="sxs-lookup"><span data-stu-id="97e77-153">[Write custom properties](/javascript/api/outlook/office.CustomProperties) set by the add-in on that item.</span></span>

- <span data-ttu-id="97e77-154">从该邮件的主题或正文中[获取所有现有已知实体](match-strings-in-an-item-as-well-known-entities.md)，而不仅仅是一个子集。</span><span class="sxs-lookup"><span data-stu-id="97e77-154">[Get all existing well-known entities](match-strings-in-an-item-as-well-known-entities.md), not just a subset, from the item's subject or body.</span></span>

- <span data-ttu-id="97e77-p106">使用 [ItemHasKnownEntity](activation-rules.md#itemhasknownentity-rule) 规则中所有的 [已知实体](../reference/manifest/rule.md#itemhasknownentity-rule)，或者 [ItemHasRegularExpressionMatch](activation-rules.md#itemhasregularexpressionmatch-rule) 规则中的 [正则表达式](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule)。以下示例遵循架构 v1.1。这说明，如果在选定邮件的主题或正文中找到一个或多个已知实体，则以下规则将激活加载项：</span><span class="sxs-lookup"><span data-stu-id="97e77-p106">Use all the [well-known entities](activation-rules.md#itemhasknownentity-rule) in [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rules, or [regular expressions](activation-rules.md#itemhasregularexpressionmatch-rule) in [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a><span data-ttu-id="97e77-158">不能执行的操作</span><span class="sxs-lookup"><span data-stu-id="97e77-158">Can't do</span></span>

- <span data-ttu-id="97e77-159">**mailbox.getCallbackTokenAsync** 提供的令牌可用于：</span><span class="sxs-lookup"><span data-stu-id="97e77-159">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="97e77-160">使用 Outlook REST API 更新或删除当前邮件，或访问用户邮箱中的其他任何邮件。</span><span class="sxs-lookup"><span data-stu-id="97e77-160">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="97e77-161">使用 Outlook REST API 获取当前日历事件项。</span><span class="sxs-lookup"><span data-stu-id="97e77-161">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="97e77-162">使用下列任一 API：</span><span class="sxs-lookup"><span data-stu-id="97e77-162">Use any of the following APIs:</span></span>
    - [<span data-ttu-id="97e77-163">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-163">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="97e77-164">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-164">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-165">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-165">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-166">item.bcc.addAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-166">item.bcc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-167">item.bcc.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-167">item.bcc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-168">item.body.prependAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-168">item.body.prependAsync</span></span>](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [<span data-ttu-id="97e77-169">item.body.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-169">item.body.setAsync</span></span>](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [<span data-ttu-id="97e77-170">item.body.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-170">item.body.setSelectedDataAsync</span></span>](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [<span data-ttu-id="97e77-171">item.cc.addAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-171">item.cc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-172">item.cc.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-172">item.cc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-173">item.end.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-173">item.end.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="97e77-174">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-174">item.location.setAsync</span></span>](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [<span data-ttu-id="97e77-175">item.optionalAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-175">item.optionalAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-176">item.optionalAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-176">item.optionalAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-177">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-177">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="97e77-178">item.requiredAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-178">item.requiredAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-179">item.requiredAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-179">item.requiredAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-180">item.start.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-180">item.start.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="97e77-181">item.subject.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-181">item.subject.setAsync</span></span>](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [<span data-ttu-id="97e77-182">item.to.addAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-182">item.to.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="97e77-183">item.to.setAsync</span><span class="sxs-lookup"><span data-stu-id="97e77-183">item.to.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a><span data-ttu-id="97e77-184">ReadWriteItem 权限</span><span class="sxs-lookup"><span data-stu-id="97e77-184">ReadWriteItem permission</span></span>

<span data-ttu-id="97e77-p107">可以在清单中的 **Permissions** 元素中指定**ReadWriteItem** 以请求此权限。在使用撰写方法（例如，**Message.to.addAsync** 或 **Message.to.setAsync**）的撰写窗体中激活的邮件加载项必须使用至少这个等级的权限。</span><span class="sxs-lookup"><span data-stu-id="97e77-p107">Specify **ReadWriteItem** in the **Permissions** element in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="97e77-187">允许事项</span><span class="sxs-lookup"><span data-stu-id="97e77-187">Can do</span></span>

- <span data-ttu-id="97e77-188">[读取和写入正在 Outlook 中查阅或撰写的邮件的所有项目级别属性](item-data.md)。</span><span class="sxs-lookup"><span data-stu-id="97e77-188">[Read and write all item-level properties](item-data.md) of the item that is being viewed or composed in Outlook.</span></span>

- <span data-ttu-id="97e77-189">[添加或删除该邮件的附件](add-and-remove-attachments-to-an-item-in-a-compose-form.md)。</span><span class="sxs-lookup"><span data-stu-id="97e77-189">[Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.</span></span>

- <span data-ttu-id="97e77-190">使用适用于 Office 的 JavaScript API 的其他所有成员，这些成员适用于邮件外接程序（**Mailbox.makeEWSRequestAsync** 除外）。</span><span class="sxs-lookup"><span data-stu-id="97e77-190">Use all other members of the JavaScript API for Office that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.</span></span>

### <a name="cant-do"></a><span data-ttu-id="97e77-191">禁止事项</span><span class="sxs-lookup"><span data-stu-id="97e77-191">Can't do</span></span>

- <span data-ttu-id="97e77-192">**mailbox.getCallbackTokenAsync** 提供的令牌可用于：</span><span class="sxs-lookup"><span data-stu-id="97e77-192">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="97e77-193">使用 Outlook REST API 更新或删除当前邮件，或访问用户邮箱中的其他任何邮件。</span><span class="sxs-lookup"><span data-stu-id="97e77-193">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="97e77-194">使用 Outlook REST API 获取当前日历事件项。</span><span class="sxs-lookup"><span data-stu-id="97e77-194">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="97e77-195">使用 **mailbox.makeEWSRequestAsync**。</span><span class="sxs-lookup"><span data-stu-id="97e77-195">Use **mailbox.makeEWSRequestAsync**.</span></span>

## <a name="readwritemailbox-permission"></a><span data-ttu-id="97e77-196">“ReadWriteMailbox”权限</span><span class="sxs-lookup"><span data-stu-id="97e77-196">ReadWriteMailbox permission</span></span>

<span data-ttu-id="97e77-p108">**ReadWriteMailbox**是最高级别权限。在清单的“权限”\*\*\*\* 元素中指定**ReadWriteMailbox**可以请求获取此权限。</span><span class="sxs-lookup"><span data-stu-id="97e77-p108">The **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.</span></span>

<span data-ttu-id="97e77-199">除了可以执行**ReadWriteItem**权限支持的操作外，还可以使用 **mailbox.getCallbackTokenAsync** 提供的令牌，通过 Exchange Web 服务 (EWS) 操作或 Outlook REST API 执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="97e77-199">In addition to what the **ReadWriteItem** permission supports, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:</span></span>

- <span data-ttu-id="97e77-200">读取和写入用户邮箱中任何邮件的所有属性。</span><span class="sxs-lookup"><span data-stu-id="97e77-200">Read and write all properties of any item in the user's mailbox.</span></span>
- <span data-ttu-id="97e77-201">创建、读取和写入该邮箱中的任何文件夹或项目。</span><span class="sxs-lookup"><span data-stu-id="97e77-201">Create, read, and write to any folder or item in that mailbox.</span></span>
- <span data-ttu-id="97e77-202">从用户邮箱发送邮件</span><span class="sxs-lookup"><span data-stu-id="97e77-202">Send an item from that mailbox</span></span>

<span data-ttu-id="97e77-203">通过 **mailbox.makeEWSRequestAsync**，可以使用以下 EWS 操作：</span><span class="sxs-lookup"><span data-stu-id="97e77-203">Through **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:</span></span>

- [<span data-ttu-id="97e77-204">CopyItem</span><span class="sxs-lookup"><span data-stu-id="97e77-204">CopyItem</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)
- [<span data-ttu-id="97e77-205">CreateFolder</span><span class="sxs-lookup"><span data-stu-id="97e77-205">CreateFolder</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)
- [<span data-ttu-id="97e77-206">CreateItem</span><span class="sxs-lookup"><span data-stu-id="97e77-206">CreateItem</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)
- [<span data-ttu-id="97e77-207">FindConversation</span><span class="sxs-lookup"><span data-stu-id="97e77-207">FindConversation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)
- [<span data-ttu-id="97e77-208">FindFolder</span><span class="sxs-lookup"><span data-stu-id="97e77-208">FindFolder</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)
- [<span data-ttu-id="97e77-209">FindItem</span><span class="sxs-lookup"><span data-stu-id="97e77-209">FindItem</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)
- [<span data-ttu-id="97e77-210">GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="97e77-210">GetConversationItems</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [<span data-ttu-id="97e77-211">GetFolder</span><span class="sxs-lookup"><span data-stu-id="97e77-211">GetFolder</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)
- [<span data-ttu-id="97e77-212">GetItem</span><span class="sxs-lookup"><span data-stu-id="97e77-212">GetItem</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)
- [<span data-ttu-id="97e77-213">MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="97e77-213">MarkAsJunk</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [<span data-ttu-id="97e77-214">MoveItem</span><span class="sxs-lookup"><span data-stu-id="97e77-214">MoveItem</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)
- [<span data-ttu-id="97e77-215">SendItem</span><span class="sxs-lookup"><span data-stu-id="97e77-215">SendItem</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)
- [<span data-ttu-id="97e77-216">UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="97e77-216">UpdateFolder</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [<span data-ttu-id="97e77-217">UpdateItem</span><span class="sxs-lookup"><span data-stu-id="97e77-217">UpdateItem</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)

<span data-ttu-id="97e77-218">尝试执行不受支持的操作会导致错误响应发生。</span><span class="sxs-lookup"><span data-stu-id="97e77-218">Attempting to use an unsupported operation will result in an error response.</span></span>

## <a name="see-also"></a><span data-ttu-id="97e77-219">另请参阅</span><span class="sxs-lookup"><span data-stu-id="97e77-219">See also</span></span>

- [<span data-ttu-id="97e77-220">Outlook 加载项的隐私、权限和安全性</span><span class="sxs-lookup"><span data-stu-id="97e77-220">Privacy, permissions, and security for Outlook add-ins</span></span>](../develop/privacy-and-security.md)
- [<span data-ttu-id="97e77-221">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="97e77-221">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
