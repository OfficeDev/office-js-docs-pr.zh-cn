---
title: "\"Context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 2ebcacb1f99df047b5f5c5ebe82c012e21e45d3c
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670137"
---
# <a name="item"></a><span data-ttu-id="c86d0-102">item</span><span class="sxs-lookup"><span data-stu-id="c86d0-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c86d0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c86d0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c86d0-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-mailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-mailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-106">Requirements</span></span>

|<span data-ttu-id="c86d0-107">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-107">Requirement</span></span>|<span data-ttu-id="c86d0-108">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-110">1.0</span></span>|
|[<span data-ttu-id="c86d0-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-112">受限</span><span class="sxs-lookup"><span data-stu-id="c86d0-112">Restricted</span></span>|
|[<span data-ttu-id="c86d0-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-114">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c86d0-115">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-115">Properties</span></span>

| <span data-ttu-id="c86d0-116">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-116">Property</span></span> | <span data-ttu-id="c86d0-117">最低</span><span class="sxs-lookup"><span data-stu-id="c86d0-117">Minimum</span></span><br><span data-ttu-id="c86d0-118">权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-118">permission level</span></span> | <span data-ttu-id="c86d0-119">型号</span><span class="sxs-lookup"><span data-stu-id="c86d0-119">Modes</span></span> | <span data-ttu-id="c86d0-120">返回类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-120">Return type</span></span> | <span data-ttu-id="c86d0-121">最低</span><span class="sxs-lookup"><span data-stu-id="c86d0-121">Minimum</span></span><br><span data-ttu-id="c86d0-122">要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-122">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="c86d0-123">attachments</span><span class="sxs-lookup"><span data-stu-id="c86d0-123">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="c86d0-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-124">ReadItem</span></span> | <span data-ttu-id="c86d0-125">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-125">Read</span></span> | <span data-ttu-id="c86d0-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span> | <span data-ttu-id="c86d0-127">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-127">1.0</span></span> |
| [<span data-ttu-id="c86d0-128">bcc</span><span class="sxs-lookup"><span data-stu-id="c86d0-128">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="c86d0-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-129">ReadItem</span></span> | <span data-ttu-id="c86d0-130">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-130">Message Compose</span></span> | [<span data-ttu-id="c86d0-131">收件人</span><span class="sxs-lookup"><span data-stu-id="c86d0-131">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c86d0-132">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-132">1.1</span></span> |
| [<span data-ttu-id="c86d0-133">body</span><span class="sxs-lookup"><span data-stu-id="c86d0-133">body</span></span>](#body-body) | <span data-ttu-id="c86d0-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-134">ReadItem</span></span> | <span data-ttu-id="c86d0-135">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-135">Compose</span></span> | [<span data-ttu-id="c86d0-136">Body</span><span class="sxs-lookup"><span data-stu-id="c86d0-136">Body</span></span>](/javascript/api/outlook/office.body) | <span data-ttu-id="c86d0-137">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-137">1.1</span></span> |
| | | <span data-ttu-id="c86d0-138">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-138">Read</span></span> | | |
| [<span data-ttu-id="c86d0-139">categories</span><span class="sxs-lookup"><span data-stu-id="c86d0-139">categories</span></span>](#categories-categories) | <span data-ttu-id="c86d0-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-140">ReadItem</span></span> | <span data-ttu-id="c86d0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-141">Compose</span></span> | [<span data-ttu-id="c86d0-142">Categories</span><span class="sxs-lookup"><span data-stu-id="c86d0-142">Categories</span></span>](/javascript/api/outlook/office.categories) | <span data-ttu-id="c86d0-143">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-143">1.8</span></span> |
| | | <span data-ttu-id="c86d0-144">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-144">Read</span></span> | | |
| [<span data-ttu-id="c86d0-145">cc</span><span class="sxs-lookup"><span data-stu-id="c86d0-145">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c86d0-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-146">ReadItem</span></span> | <span data-ttu-id="c86d0-147">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-147">Message Compose</span></span> | [<span data-ttu-id="c86d0-148">收件人</span><span class="sxs-lookup"><span data-stu-id="c86d0-148">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c86d0-149">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-149">1.0</span></span> |
| | | <span data-ttu-id="c86d0-150">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-150">Message Read</span></span> | <span data-ttu-id="c86d0-151"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-151">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="c86d0-152">conversationId</span><span class="sxs-lookup"><span data-stu-id="c86d0-152">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c86d0-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-153">ReadItem</span></span> | <span data-ttu-id="c86d0-154">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-154">Message Compose</span></span> | <span data-ttu-id="c86d0-155">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-155">String</span></span> | <span data-ttu-id="c86d0-156">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-156">1.0</span></span> |
| | | <span data-ttu-id="c86d0-157">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-157">Message Read</span></span> | | |
| [<span data-ttu-id="c86d0-158">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c86d0-158">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c86d0-159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-159">ReadItem</span></span> | <span data-ttu-id="c86d0-160">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-160">Read</span></span> | <span data-ttu-id="c86d0-161">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-161">Date</span></span> | <span data-ttu-id="c86d0-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-162">1.0</span></span> |
| [<span data-ttu-id="c86d0-163">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c86d0-163">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c86d0-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-164">ReadItem</span></span> | <span data-ttu-id="c86d0-165">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-165">Read</span></span> | <span data-ttu-id="c86d0-166">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-166">Date</span></span> | <span data-ttu-id="c86d0-167">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-167">1.0</span></span> |
| [<span data-ttu-id="c86d0-168">end</span><span class="sxs-lookup"><span data-stu-id="c86d0-168">end</span></span>](#end-datetime) | <span data-ttu-id="c86d0-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-169">ReadItem</span></span> | <span data-ttu-id="c86d0-170">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-170">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-171">Time</span><span class="sxs-lookup"><span data-stu-id="c86d0-171">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="c86d0-172">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-172">1.0</span></span> |
| | | <span data-ttu-id="c86d0-173">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-173">Appointment Attendee</span></span> | <span data-ttu-id="c86d0-174">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-174">Date</span></span> | |
| | | <span data-ttu-id="c86d0-175">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-175">Message Read</span></span><br><span data-ttu-id="c86d0-176">（会议请求）</span><span class="sxs-lookup"><span data-stu-id="c86d0-176">(Meeting Request)</span></span> | <span data-ttu-id="c86d0-177">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-177">Date</span></span> | |
| [<span data-ttu-id="c86d0-178">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c86d0-178">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="c86d0-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-179">ReadItem</span></span> | <span data-ttu-id="c86d0-180">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-180">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-181">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c86d0-181">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation) | <span data-ttu-id="c86d0-182">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-182">1.8</span></span> |
| | | <span data-ttu-id="c86d0-183">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-183">Appointment Attendee</span></span> | | |
| [<span data-ttu-id="c86d0-184">from</span><span class="sxs-lookup"><span data-stu-id="c86d0-184">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="c86d0-185">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-185">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-186">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-186">Message Compose</span></span> | [<span data-ttu-id="c86d0-187">From</span><span class="sxs-lookup"><span data-stu-id="c86d0-187">From</span></span>](/javascript/api/outlook/office.from) | <span data-ttu-id="c86d0-188">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-188">1.7</span></span> |
| | <span data-ttu-id="c86d0-189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-189">ReadItem</span></span> | <span data-ttu-id="c86d0-190">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-190">Message Read</span></span> | [<span data-ttu-id="c86d0-191">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c86d0-191">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="c86d0-192">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-192">1.0</span></span> |
| [<span data-ttu-id="c86d0-193">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="c86d0-193">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="c86d0-194">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-194">ReadItem</span></span> | <span data-ttu-id="c86d0-195">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-195">Message Compose</span></span> | [<span data-ttu-id="c86d0-196">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c86d0-196">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders) | <span data-ttu-id="c86d0-197">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-197">1.8</span></span> |
| [<span data-ttu-id="c86d0-198">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c86d0-198">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c86d0-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-199">ReadItem</span></span> | <span data-ttu-id="c86d0-200">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-200">Message Read</span></span> | <span data-ttu-id="c86d0-201">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-201">String</span></span> | <span data-ttu-id="c86d0-202">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-202">1.0</span></span> |
| [<span data-ttu-id="c86d0-203">itemClass</span><span class="sxs-lookup"><span data-stu-id="c86d0-203">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c86d0-204">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-204">ReadItem</span></span> | <span data-ttu-id="c86d0-205">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-205">Read</span></span> | <span data-ttu-id="c86d0-206">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-206">String</span></span> | <span data-ttu-id="c86d0-207">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-207">1.0</span></span> |
| [<span data-ttu-id="c86d0-208">itemId</span><span class="sxs-lookup"><span data-stu-id="c86d0-208">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c86d0-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-209">ReadItem</span></span> | <span data-ttu-id="c86d0-210">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-210">Read</span></span> | <span data-ttu-id="c86d0-211">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-211">String</span></span> | <span data-ttu-id="c86d0-212">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-212">1.0</span></span> |
| [<span data-ttu-id="c86d0-213">itemType</span><span class="sxs-lookup"><span data-stu-id="c86d0-213">itemType</span></span>](#itemtype-mailboxenumsitemtype) | <span data-ttu-id="c86d0-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-214">ReadItem</span></span> | <span data-ttu-id="c86d0-215">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-215">Compose</span></span> | [<span data-ttu-id="c86d0-216">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="c86d0-216">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype) | <span data-ttu-id="c86d0-217">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-217">1.0</span></span> |
| | | <span data-ttu-id="c86d0-218">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-218">Read</span></span> | | |
| [<span data-ttu-id="c86d0-219">location</span><span class="sxs-lookup"><span data-stu-id="c86d0-219">location</span></span>](#location-stringlocation) | <span data-ttu-id="c86d0-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-220">ReadItem</span></span> | <span data-ttu-id="c86d0-221">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-221">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-222">位置</span><span class="sxs-lookup"><span data-stu-id="c86d0-222">Location</span></span>](/javascript/api/outlook/office.location) | <span data-ttu-id="c86d0-223">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-223">1.0</span></span> |
| | | <span data-ttu-id="c86d0-224">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-224">Appointment Attendee</span></span> | <span data-ttu-id="c86d0-225">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-225">String</span></span> | |
| | | <span data-ttu-id="c86d0-226">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-226">Message Read</span></span><br><span data-ttu-id="c86d0-227">（会议请求）</span><span class="sxs-lookup"><span data-stu-id="c86d0-227">(Meeting Request)</span></span> | <span data-ttu-id="c86d0-228">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-228">String</span></span> | |
| [<span data-ttu-id="c86d0-229">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c86d0-229">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c86d0-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-230">ReadItem</span></span> | <span data-ttu-id="c86d0-231">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-231">Read</span></span> | <span data-ttu-id="c86d0-232">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-232">String</span></span> | <span data-ttu-id="c86d0-233">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-233">1.0</span></span> |
| [<span data-ttu-id="c86d0-234">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c86d0-234">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="c86d0-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-235">ReadItem</span></span> | <span data-ttu-id="c86d0-236">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-236">Message Compose</span></span> | [<span data-ttu-id="c86d0-237">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c86d0-237">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages) | <span data-ttu-id="c86d0-238">1.3</span><span class="sxs-lookup"><span data-stu-id="c86d0-238">1.3</span></span> |
| | <span data-ttu-id="c86d0-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-239">ReadItem</span></span> | <span data-ttu-id="c86d0-240">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-240">Message Read</span></span> | | |
| [<span data-ttu-id="c86d0-241">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c86d0-241">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c86d0-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-242">ReadItem</span></span> | <span data-ttu-id="c86d0-243">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-243">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-244">收件人</span><span class="sxs-lookup"><span data-stu-id="c86d0-244">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c86d0-245">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-245">1.0</span></span> |
| | | <span data-ttu-id="c86d0-246">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-246">Appointment Attendee</span></span> | <span data-ttu-id="c86d0-247"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-247">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="c86d0-248">organizer</span><span class="sxs-lookup"><span data-stu-id="c86d0-248">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="c86d0-249">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-249">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-250">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-250">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-251">Organizer</span><span class="sxs-lookup"><span data-stu-id="c86d0-251">Organizer</span></span>](/javascript/api/outlook/office.organizer) | <span data-ttu-id="c86d0-252">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-252">1.7</span></span> |
| | <span data-ttu-id="c86d0-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-253">ReadItem</span></span> | <span data-ttu-id="c86d0-254">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-254">Appointment Attendee</span></span> | [<span data-ttu-id="c86d0-255">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c86d0-255">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="c86d0-256">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-256">1.0</span></span> |
| [<span data-ttu-id="c86d0-257">recurrence</span><span class="sxs-lookup"><span data-stu-id="c86d0-257">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="c86d0-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-258">ReadItem</span></span> | <span data-ttu-id="c86d0-259">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-259">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-260">循环</span><span class="sxs-lookup"><span data-stu-id="c86d0-260">Recurrence</span></span>](/javascript/api/outlook/office.recurrence) | <span data-ttu-id="c86d0-261">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-261">1.7</span></span> |
| | | <span data-ttu-id="c86d0-262">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-262">Appointment Attendee</span></span> | | |
| | | <span data-ttu-id="c86d0-263">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-263">Message Read</span></span><br><span data-ttu-id="c86d0-264">（会议请求）</span><span class="sxs-lookup"><span data-stu-id="c86d0-264">(Meeting Request)</span></span> | | |
| [<span data-ttu-id="c86d0-265">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c86d0-265">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c86d0-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-266">ReadItem</span></span> | <span data-ttu-id="c86d0-267">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-267">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-268">收件人</span><span class="sxs-lookup"><span data-stu-id="c86d0-268">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c86d0-269">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-269">1.0</span></span> |
| | | <span data-ttu-id="c86d0-270">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-270">Appointment Attendee</span></span> | <span data-ttu-id="c86d0-271"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-271">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="c86d0-272">sender</span><span class="sxs-lookup"><span data-stu-id="c86d0-272">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="c86d0-273">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-273">ReadItem</span></span> | <span data-ttu-id="c86d0-274">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-274">Message Read</span></span> | [<span data-ttu-id="c86d0-275">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c86d0-275">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="c86d0-276">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-276">1.0</span></span> |
| [<span data-ttu-id="c86d0-277">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="c86d0-277">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c86d0-278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-278">ReadItem</span></span> | <span data-ttu-id="c86d0-279">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-279">Compose</span></span> | <span data-ttu-id="c86d0-280">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-280">String</span></span> | <span data-ttu-id="c86d0-281">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-281">1.7</span></span> |
| | | <span data-ttu-id="c86d0-282">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-282">Read</span></span> | | |
| [<span data-ttu-id="c86d0-283">start</span><span class="sxs-lookup"><span data-stu-id="c86d0-283">start</span></span>](#start-datetime) | <span data-ttu-id="c86d0-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-284">ReadItem</span></span> | <span data-ttu-id="c86d0-285">约会组织者</span><span class="sxs-lookup"><span data-stu-id="c86d0-285">Appointment Organizer</span></span> | [<span data-ttu-id="c86d0-286">Time</span><span class="sxs-lookup"><span data-stu-id="c86d0-286">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="c86d0-287">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-287">1.0</span></span> |
| | | <span data-ttu-id="c86d0-288">约会与会者</span><span class="sxs-lookup"><span data-stu-id="c86d0-288">Appointment Attendee</span></span> | <span data-ttu-id="c86d0-289">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-289">Date</span></span> | |
| | | <span data-ttu-id="c86d0-290">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-290">Message Read</span></span><br><span data-ttu-id="c86d0-291">（会议请求）</span><span class="sxs-lookup"><span data-stu-id="c86d0-291">(Meeting Request)</span></span> | <span data-ttu-id="c86d0-292">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-292">Date</span></span> | |
| [<span data-ttu-id="c86d0-293">subject</span><span class="sxs-lookup"><span data-stu-id="c86d0-293">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="c86d0-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-294">ReadItem</span></span> | <span data-ttu-id="c86d0-295">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-295">Compose</span></span> | [<span data-ttu-id="c86d0-296">Subject</span><span class="sxs-lookup"><span data-stu-id="c86d0-296">Subject</span></span>](/javascript/api/outlook/office.subject) | <span data-ttu-id="c86d0-297">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-297">1.0</span></span> |
| | | <span data-ttu-id="c86d0-298">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-298">Read</span></span> | <span data-ttu-id="c86d0-299">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-299">String</span></span> | |
| [<span data-ttu-id="c86d0-300">to</span><span class="sxs-lookup"><span data-stu-id="c86d0-300">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c86d0-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-301">ReadItem</span></span> | <span data-ttu-id="c86d0-302">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-302">Message Compose</span></span> | [<span data-ttu-id="c86d0-303">收件人</span><span class="sxs-lookup"><span data-stu-id="c86d0-303">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c86d0-304">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-304">1.0</span></span> |
| | | <span data-ttu-id="c86d0-305">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-305">Message Read</span></span> | <span data-ttu-id="c86d0-306"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-306">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |

##### <a name="methods"></a><span data-ttu-id="c86d0-307">方法</span><span class="sxs-lookup"><span data-stu-id="c86d0-307">Methods</span></span>

| <span data-ttu-id="c86d0-308">方法</span><span class="sxs-lookup"><span data-stu-id="c86d0-308">Method</span></span> | <span data-ttu-id="c86d0-309">最低</span><span class="sxs-lookup"><span data-stu-id="c86d0-309">Minimum</span></span><br><span data-ttu-id="c86d0-310">权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-310">permission level</span></span> | <span data-ttu-id="c86d0-311">型号</span><span class="sxs-lookup"><span data-stu-id="c86d0-311">Modes</span></span> | <span data-ttu-id="c86d0-312">最低</span><span class="sxs-lookup"><span data-stu-id="c86d0-312">Minimum</span></span><br><span data-ttu-id="c86d0-313">要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-313">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="c86d0-314">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-314">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c86d0-315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-315">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-316">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-316">Compose</span></span> | <span data-ttu-id="c86d0-317">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-317">1.1</span></span> |
| [<span data-ttu-id="c86d0-318">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="c86d0-318">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="c86d0-319">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-319">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-320">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-320">Compose</span></span> | <span data-ttu-id="c86d0-321">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-321">1.8</span></span> |
| [<span data-ttu-id="c86d0-322">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-322">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c86d0-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-323">ReadItem</span></span> | <span data-ttu-id="c86d0-324">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-324">Compose</span></span><br><span data-ttu-id="c86d0-325">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-325">Read</span></span> | <span data-ttu-id="c86d0-326">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-326">1.7</span></span> |
| [<span data-ttu-id="c86d0-327">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-327">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c86d0-328">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-328">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-329">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-329">Compose</span></span> | <span data-ttu-id="c86d0-330">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-330">1.1</span></span> |
| [<span data-ttu-id="c86d0-331">close</span><span class="sxs-lookup"><span data-stu-id="c86d0-331">close</span></span>](#close) | <span data-ttu-id="c86d0-332">受限</span><span class="sxs-lookup"><span data-stu-id="c86d0-332">Restricted</span></span> | <span data-ttu-id="c86d0-333">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-333">Compose</span></span> | <span data-ttu-id="c86d0-334">1.3</span><span class="sxs-lookup"><span data-stu-id="c86d0-334">1.3</span></span> |
| [<span data-ttu-id="c86d0-335">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c86d0-335">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="c86d0-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-336">ReadItem</span></span> | <span data-ttu-id="c86d0-337">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-337">Read</span></span> | <span data-ttu-id="c86d0-338">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-338">1.0</span></span> |
| [<span data-ttu-id="c86d0-339">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c86d0-339">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="c86d0-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-340">ReadItem</span></span> | <span data-ttu-id="c86d0-341">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-341">Read</span></span> | <span data-ttu-id="c86d0-342">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-342">1.0</span></span> |
| [<span data-ttu-id="c86d0-343">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-343">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="c86d0-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-344">ReadItem</span></span> | <span data-ttu-id="c86d0-345">邮件读取</span><span class="sxs-lookup"><span data-stu-id="c86d0-345">Message Read</span></span> | <span data-ttu-id="c86d0-346">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-346">1.8</span></span> |
| [<span data-ttu-id="c86d0-347">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-347">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="c86d0-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-348">ReadItem</span></span> | <span data-ttu-id="c86d0-349">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-349">Compose</span></span><br><span data-ttu-id="c86d0-350">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-350">Read</span></span> | <span data-ttu-id="c86d0-351">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-351">1.8</span></span> |
| [<span data-ttu-id="c86d0-352">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-352">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="c86d0-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-353">ReadItem</span></span> | <span data-ttu-id="c86d0-354">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-354">Compose</span></span> | <span data-ttu-id="c86d0-355">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-355">1.8</span></span> |
| [<span data-ttu-id="c86d0-356">getEntities</span><span class="sxs-lookup"><span data-stu-id="c86d0-356">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="c86d0-357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-357">ReadItem</span></span> | <span data-ttu-id="c86d0-358">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-358">Read</span></span> | <span data-ttu-id="c86d0-359">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-359">1.0</span></span> |
| [<span data-ttu-id="c86d0-360">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c86d0-360">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c86d0-361">受限</span><span class="sxs-lookup"><span data-stu-id="c86d0-361">Restricted</span></span> | <span data-ttu-id="c86d0-362">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-362">Read</span></span> | <span data-ttu-id="c86d0-363">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-363">1.0</span></span> |
| [<span data-ttu-id="c86d0-364">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c86d0-364">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c86d0-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-365">ReadItem</span></span> | <span data-ttu-id="c86d0-366">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-366">Read</span></span> | <span data-ttu-id="c86d0-367">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-367">1.0</span></span> |
| [<span data-ttu-id="c86d0-368">Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="c86d0-368">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="c86d0-369">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-369">ReadItem</span></span> | <span data-ttu-id="c86d0-370">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-370">Read</span></span> | <span data-ttu-id="c86d0-371">预览</span><span class="sxs-lookup"><span data-stu-id="c86d0-371">Preview</span></span> |
| [<span data-ttu-id="c86d0-372">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-372">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="c86d0-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-373">ReadItem</span></span> | <span data-ttu-id="c86d0-374">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-374">Compose</span></span> | <span data-ttu-id="c86d0-375">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-375">1.8</span></span> |
| [<span data-ttu-id="c86d0-376">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c86d0-376">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c86d0-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-377">ReadItem</span></span> | <span data-ttu-id="c86d0-378">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-378">Read</span></span> | <span data-ttu-id="c86d0-379">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-379">1.0</span></span> |
| [<span data-ttu-id="c86d0-380">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c86d0-380">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c86d0-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-381">ReadItem</span></span> | <span data-ttu-id="c86d0-382">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-382">Read</span></span> | <span data-ttu-id="c86d0-383">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-383">1.0</span></span> |
| [<span data-ttu-id="c86d0-384">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-384">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c86d0-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-385">ReadItem</span></span> | <span data-ttu-id="c86d0-386">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-386">Compose</span></span> | <span data-ttu-id="c86d0-387">1.2</span><span class="sxs-lookup"><span data-stu-id="c86d0-387">1.2</span></span> |
| [<span data-ttu-id="c86d0-388">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="c86d0-388">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="c86d0-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-389">ReadItem</span></span> | <span data-ttu-id="c86d0-390">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-390">Read</span></span> | <span data-ttu-id="c86d0-391">1.6</span><span class="sxs-lookup"><span data-stu-id="c86d0-391">1.6</span></span> |
| [<span data-ttu-id="c86d0-392">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="c86d0-392">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c86d0-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-393">ReadItem</span></span> | <span data-ttu-id="c86d0-394">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-394">Read</span></span> | <span data-ttu-id="c86d0-395">1.6</span><span class="sxs-lookup"><span data-stu-id="c86d0-395">1.6</span></span> |
| [<span data-ttu-id="c86d0-396">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-396">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="c86d0-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-397">ReadItem</span></span> | <span data-ttu-id="c86d0-398">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-398">Compose</span></span><br><span data-ttu-id="c86d0-399">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-399">Read</span></span> | <span data-ttu-id="c86d0-400">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-400">1.8</span></span> |
| [<span data-ttu-id="c86d0-401">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-401">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c86d0-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-402">ReadItem</span></span> | <span data-ttu-id="c86d0-403">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-403">Compose</span></span><br><span data-ttu-id="c86d0-404">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-404">Read</span></span> | <span data-ttu-id="c86d0-405">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-405">1.0</span></span> |
| [<span data-ttu-id="c86d0-406">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-406">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c86d0-407">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-407">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-408">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-408">Compose</span></span> | <span data-ttu-id="c86d0-409">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-409">1.1</span></span> |
| [<span data-ttu-id="c86d0-410">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-410">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c86d0-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-411">ReadItem</span></span> | <span data-ttu-id="c86d0-412">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-412">Compose</span></span><br><span data-ttu-id="c86d0-413">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-413">Read</span></span> | <span data-ttu-id="c86d0-414">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-414">1.7</span></span> |
| [<span data-ttu-id="c86d0-415">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-415">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c86d0-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-416">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-417">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-417">Compose</span></span> | <span data-ttu-id="c86d0-418">1.3</span><span class="sxs-lookup"><span data-stu-id="c86d0-418">1.3</span></span> |
| [<span data-ttu-id="c86d0-419">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c86d0-419">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c86d0-420">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-420">ReadWriteItem</span></span> | <span data-ttu-id="c86d0-421">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-421">Compose</span></span> | <span data-ttu-id="c86d0-422">1.2</span><span class="sxs-lookup"><span data-stu-id="c86d0-422">1.2</span></span> |

##### <a name="events"></a><span data-ttu-id="c86d0-423">活动</span><span class="sxs-lookup"><span data-stu-id="c86d0-423">Events</span></span>

<span data-ttu-id="c86d0-424">您可以分别使用[addHandlerAsync](#addhandlerasynceventtype-handler-options-callback)和[removeHandlerAsync](#removehandlerasynceventtype-options-callback)订阅和取消订阅以下事件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-424">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="c86d0-425">事件</span><span class="sxs-lookup"><span data-stu-id="c86d0-425">Event</span></span> | <span data-ttu-id="c86d0-426">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-426">Description</span></span> | <span data-ttu-id="c86d0-427">最低</span><span class="sxs-lookup"><span data-stu-id="c86d0-427">Minimum</span></span><br><span data-ttu-id="c86d0-428">要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-428">requirement set</span></span> |
|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c86d0-429">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="c86d0-429">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c86d0-430">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-430">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="c86d0-431">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-431">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="c86d0-432">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-432">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="c86d0-433">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="c86d0-433">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="c86d0-434">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-434">1.8</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c86d0-435">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="c86d0-435">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c86d0-436">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-436">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c86d0-437">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="c86d0-437">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c86d0-438">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-438">1.7</span></span> |

### <a name="example"></a><span data-ttu-id="c86d0-439">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-439">Example</span></span>

<span data-ttu-id="c86d0-440">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-440">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

## <a name="property-details"></a><span data-ttu-id="c86d0-441">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="c86d0-441">Property details</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c86d0-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c86d0-443">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-443">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c86d0-444">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-444">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-445">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="c86d0-445">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c86d0-446">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="c86d0-446">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-447">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-447">Type</span></span>

*   <span data-ttu-id="c86d0-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-449">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-449">Requirements</span></span>

|<span data-ttu-id="c86d0-450">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-450">Requirement</span></span>|<span data-ttu-id="c86d0-451">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-451">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-452">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-453">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-453">1.0</span></span>|
|[<span data-ttu-id="c86d0-454">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-455">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-456">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-457">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-457">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-458">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-458">Example</span></span>

<span data-ttu-id="c86d0-459">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="c86d0-459">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c86d0-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c86d0-461">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-461">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c86d0-462">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-462">Compose mode only.</span></span>

<span data-ttu-id="c86d0-463">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-463">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-464">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-464">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c86d0-465">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-465">Get 500 members maximum.</span></span>
- <span data-ttu-id="c86d0-466">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-466">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-467">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-467">Type</span></span>

*   [<span data-ttu-id="c86d0-468">收件人</span><span class="sxs-lookup"><span data-stu-id="c86d0-468">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c86d0-469">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-469">Requirements</span></span>

|<span data-ttu-id="c86d0-470">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-470">Requirement</span></span>|<span data-ttu-id="c86d0-471">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-472">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-473">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-473">1.1</span></span>|
|[<span data-ttu-id="c86d0-474">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-475">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-476">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-477">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-477">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-478">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-478">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="c86d0-479">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="c86d0-479">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="c86d0-480">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-480">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-481">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-481">Type</span></span>

*   [<span data-ttu-id="c86d0-482">Body</span><span class="sxs-lookup"><span data-stu-id="c86d0-482">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="c86d0-483">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-483">Requirements</span></span>

|<span data-ttu-id="c86d0-484">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-484">Requirement</span></span>|<span data-ttu-id="c86d0-485">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-486">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-487">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-487">1.1</span></span>|
|[<span data-ttu-id="c86d0-488">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-488">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-489">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-490">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-490">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-491">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-491">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-492">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-492">Example</span></span>

<span data-ttu-id="c86d0-493">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="c86d0-493">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="c86d0-494">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="c86d0-494">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="c86d0-495">类别：[类别](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="c86d0-495">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="c86d0-496">获取一个对象，该对象提供用于管理项的类别的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-496">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-497">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-497">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-498">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-498">Type</span></span>

*   [<span data-ttu-id="c86d0-499">Categories</span><span class="sxs-lookup"><span data-stu-id="c86d0-499">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="c86d0-500">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-500">Requirements</span></span>

|<span data-ttu-id="c86d0-501">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-501">Requirement</span></span>|<span data-ttu-id="c86d0-502">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-504">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-504">1.8</span></span>|
|[<span data-ttu-id="c86d0-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-506">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-508">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-508">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-509">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-509">Example</span></span>

<span data-ttu-id="c86d0-510">此示例获取项的类别。</span><span class="sxs-lookup"><span data-stu-id="c86d0-510">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c86d0-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c86d0-512">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c86d0-512">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c86d0-513">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-513">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-514">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-514">Read mode</span></span>

<span data-ttu-id="c86d0-515">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-515">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="c86d0-516">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-516">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-517">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-517">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-518">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-518">Compose mode</span></span>

<span data-ttu-id="c86d0-519">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-519">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="c86d0-520">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-520">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-521">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-521">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c86d0-522">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-522">Get 500 members maximum.</span></span>
- <span data-ttu-id="c86d0-523">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-523">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-524">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-524">Type</span></span>

*   <span data-ttu-id="c86d0-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-526">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-526">Requirements</span></span>

|<span data-ttu-id="c86d0-527">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-527">Requirement</span></span>|<span data-ttu-id="c86d0-528">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-529">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-530">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-530">1.0</span></span>|
|[<span data-ttu-id="c86d0-531">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-531">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-532">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-533">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-533">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-534">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-534">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="c86d0-535">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="c86d0-535">(nullable) conversationId: String</span></span>

<span data-ttu-id="c86d0-536">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-536">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c86d0-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c86d0-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-541">Type</span><span class="sxs-lookup"><span data-stu-id="c86d0-541">Type</span></span>

*   <span data-ttu-id="c86d0-542">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-542">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-543">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-543">Requirements</span></span>

|<span data-ttu-id="c86d0-544">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-544">Requirement</span></span>|<span data-ttu-id="c86d0-545">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-546">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-547">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-547">1.0</span></span>|
|[<span data-ttu-id="c86d0-548">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-549">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-550">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-551">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-552">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-552">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="c86d0-553">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="c86d0-553">dateTimeCreated: Date</span></span>

<span data-ttu-id="c86d0-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-556">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-556">Type</span></span>

*   <span data-ttu-id="c86d0-557">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-557">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-558">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-558">Requirements</span></span>

|<span data-ttu-id="c86d0-559">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-559">Requirement</span></span>|<span data-ttu-id="c86d0-560">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-561">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-562">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-562">1.0</span></span>|
|[<span data-ttu-id="c86d0-563">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-564">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-565">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-566">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-566">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-567">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-567">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="c86d0-568">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="c86d0-568">dateTimeModified: Date</span></span>

<span data-ttu-id="c86d0-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-571">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-571">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-572">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-572">Type</span></span>

*   <span data-ttu-id="c86d0-573">日期</span><span class="sxs-lookup"><span data-stu-id="c86d0-573">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-574">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-574">Requirements</span></span>

|<span data-ttu-id="c86d0-575">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-575">Requirement</span></span>|<span data-ttu-id="c86d0-576">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-577">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-578">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-578">1.0</span></span>|
|[<span data-ttu-id="c86d0-579">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-580">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-581">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-582">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-582">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-583">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-583">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c86d0-584">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c86d0-584">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c86d0-585">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c86d0-585">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c86d0-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-588">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-588">Read mode</span></span>

<span data-ttu-id="c86d0-589">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-589">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-590">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-590">Compose mode</span></span>

<span data-ttu-id="c86d0-591">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-591">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c86d0-592">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c86d0-592">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c86d0-593">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="c86d0-593">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="c86d0-594">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-594">Type</span></span>

*   <span data-ttu-id="c86d0-595">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c86d0-595">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-596">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-596">Requirements</span></span>

|<span data-ttu-id="c86d0-597">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-597">Requirement</span></span>|<span data-ttu-id="c86d0-598">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-599">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-600">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-600">1.0</span></span>|
|[<span data-ttu-id="c86d0-601">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-602">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-603">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-604">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-604">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="c86d0-605">enhancedLocation： [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="c86d0-605">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="c86d0-606">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="c86d0-606">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-607">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-607">Read mode</span></span>

<span data-ttu-id="c86d0-608">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，该对象允许您获取与约会关联的一组位置（每个由[LocationDetails](/javascript/api/outlook/office.locationdetails)对象表示）。</span><span class="sxs-lookup"><span data-stu-id="c86d0-608">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-609">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-609">Compose mode</span></span>

<span data-ttu-id="c86d0-610">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-610">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-611">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-611">Type</span></span>

*   [<span data-ttu-id="c86d0-612">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c86d0-612">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="c86d0-613">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-613">Requirements</span></span>

|<span data-ttu-id="c86d0-614">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-614">Requirement</span></span>|<span data-ttu-id="c86d0-615">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-616">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-617">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-617">1.8</span></span>|
|[<span data-ttu-id="c86d0-618">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-619">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-620">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-621">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-622">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-622">Example</span></span>

<span data-ttu-id="c86d0-623">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="c86d0-623">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="c86d0-624">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c86d0-624">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="c86d0-625">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c86d0-625">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c86d0-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-628">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-628">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-629">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-629">Read mode</span></span>

<span data-ttu-id="c86d0-630">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-630">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-631">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-631">Compose mode</span></span>

<span data-ttu-id="c86d0-632">`from`属性返回一个`From`对象，该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-632">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-633">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-633">Type</span></span>

*   <span data-ttu-id="c86d0-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c86d0-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-635">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-635">Requirements</span></span>

|<span data-ttu-id="c86d0-636">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-636">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c86d0-637">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-638">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-638">1.0</span></span>|<span data-ttu-id="c86d0-639">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-639">1.7</span></span>|
|[<span data-ttu-id="c86d0-640">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-640">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-641">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-641">ReadItem</span></span>|<span data-ttu-id="c86d0-642">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-642">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-643">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-643">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-644">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-644">Read</span></span>|<span data-ttu-id="c86d0-645">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-645">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="c86d0-646">internetHeaders： [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="c86d0-646">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="c86d0-647">获取或设置邮件的自定义 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="c86d0-647">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="c86d0-648">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-648">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-649">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-649">Type</span></span>

*   [<span data-ttu-id="c86d0-650">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c86d0-650">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="c86d0-651">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-651">Requirements</span></span>

|<span data-ttu-id="c86d0-652">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-652">Requirement</span></span>|<span data-ttu-id="c86d0-653">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-654">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-655">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-655">1.8</span></span>|
|[<span data-ttu-id="c86d0-656">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-657">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-658">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-659">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-659">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-660">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-660">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="c86d0-661">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="c86d0-661">internetMessageId: String</span></span>

<span data-ttu-id="c86d0-p116">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-664">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-664">Type</span></span>

*   <span data-ttu-id="c86d0-665">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-665">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-666">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-666">Requirements</span></span>

|<span data-ttu-id="c86d0-667">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-667">Requirement</span></span>|<span data-ttu-id="c86d0-668">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-669">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-670">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-670">1.0</span></span>|
|[<span data-ttu-id="c86d0-671">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-672">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-673">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-674">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-674">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-675">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-675">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="c86d0-676">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="c86d0-676">itemClass: String</span></span>

<span data-ttu-id="c86d0-p117">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c86d0-p118">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c86d0-681">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-681">Type</span></span>|<span data-ttu-id="c86d0-682">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-682">Description</span></span>|<span data-ttu-id="c86d0-683">项目类</span><span class="sxs-lookup"><span data-stu-id="c86d0-683">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c86d0-684">约会项目</span><span class="sxs-lookup"><span data-stu-id="c86d0-684">Appointment items</span></span>|<span data-ttu-id="c86d0-685">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="c86d0-685">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="c86d0-686">邮件项目</span><span class="sxs-lookup"><span data-stu-id="c86d0-686">Message items</span></span>|<span data-ttu-id="c86d0-687">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="c86d0-687">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c86d0-688">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-688">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-689">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-689">Type</span></span>

*   <span data-ttu-id="c86d0-690">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-690">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-691">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-691">Requirements</span></span>

|<span data-ttu-id="c86d0-692">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-692">Requirement</span></span>|<span data-ttu-id="c86d0-693">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-694">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-695">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-695">1.0</span></span>|
|[<span data-ttu-id="c86d0-696">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-697">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-697">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-698">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-699">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-699">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-700">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-700">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c86d0-701">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="c86d0-701">(nullable) itemId: String</span></span>

<span data-ttu-id="c86d0-p119">获取当前项目的 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-704">`itemId` 属性返回的标识符与 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="c86d0-704">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="c86d0-705">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="c86d0-705">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c86d0-706">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="c86d0-706">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c86d0-707">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="c86d0-707">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c86d0-p121">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-710">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-710">Type</span></span>

*   <span data-ttu-id="c86d0-711">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-711">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-712">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-712">Requirements</span></span>

|<span data-ttu-id="c86d0-713">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-713">Requirement</span></span>|<span data-ttu-id="c86d0-714">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-715">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-716">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-716">1.0</span></span>|
|[<span data-ttu-id="c86d0-717">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-718">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-719">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-720">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-720">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-721">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-721">Example</span></span>

<span data-ttu-id="c86d0-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-mailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="c86d0-724">itemType： [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c86d0-724">itemType: [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c86d0-725">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="c86d0-725">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c86d0-726">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="c86d0-726">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-727">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-727">Type</span></span>

*   [<span data-ttu-id="c86d0-728">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="c86d0-728">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c86d0-729">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-729">Requirements</span></span>

|<span data-ttu-id="c86d0-730">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-730">Requirement</span></span>|<span data-ttu-id="c86d0-731">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-731">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-732">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-732">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-733">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-733">1.0</span></span>|
|[<span data-ttu-id="c86d0-734">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-734">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-735">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-735">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-736">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-736">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-737">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-737">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-738">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-738">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="c86d0-739">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c86d0-739">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="c86d0-740">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="c86d0-740">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-741">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-741">Read mode</span></span>

<span data-ttu-id="c86d0-742">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="c86d0-742">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-743">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-743">Compose mode</span></span>

<span data-ttu-id="c86d0-744">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-744">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-745">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-745">Type</span></span>

*   <span data-ttu-id="c86d0-746">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c86d0-746">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-747">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-747">Requirements</span></span>

|<span data-ttu-id="c86d0-748">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-748">Requirement</span></span>|<span data-ttu-id="c86d0-749">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-750">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-751">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-751">1.0</span></span>|
|[<span data-ttu-id="c86d0-752">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-752">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-753">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-754">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-754">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-755">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-755">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c86d0-756">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="c86d0-756">normalizedSubject: String</span></span>

<span data-ttu-id="c86d0-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c86d0-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-761">Type</span><span class="sxs-lookup"><span data-stu-id="c86d0-761">Type</span></span>

*   <span data-ttu-id="c86d0-762">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-762">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-763">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-763">Requirements</span></span>

|<span data-ttu-id="c86d0-764">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-764">Requirement</span></span>|<span data-ttu-id="c86d0-765">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-765">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-766">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-766">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-767">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-767">1.0</span></span>|
|[<span data-ttu-id="c86d0-768">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-768">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-769">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-769">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-770">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-770">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-771">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-771">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-772">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-772">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="c86d0-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c86d0-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="c86d0-774">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-774">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-775">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-775">Type</span></span>

*   [<span data-ttu-id="c86d0-776">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c86d0-776">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c86d0-777">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-777">Requirements</span></span>

|<span data-ttu-id="c86d0-778">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-778">Requirement</span></span>|<span data-ttu-id="c86d0-779">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-780">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-781">1.3</span><span class="sxs-lookup"><span data-stu-id="c86d0-781">1.3</span></span>|
|[<span data-ttu-id="c86d0-782">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-782">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-783">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-783">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-784">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-784">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-785">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-785">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-786">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-786">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c86d0-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c86d0-788">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c86d0-788">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c86d0-789">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-789">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-790">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-790">Read mode</span></span>

<span data-ttu-id="c86d0-791">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-791">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="c86d0-792">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-792">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-793">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-793">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-794">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-794">Compose mode</span></span>

<span data-ttu-id="c86d0-795">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-795">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="c86d0-796">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-796">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-797">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-797">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c86d0-798">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-798">Get 500 members maximum.</span></span>
- <span data-ttu-id="c86d0-799">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-799">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-800">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-800">Type</span></span>

*   <span data-ttu-id="c86d0-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-802">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-802">Requirements</span></span>

|<span data-ttu-id="c86d0-803">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-803">Requirement</span></span>|<span data-ttu-id="c86d0-804">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-804">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-805">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-805">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-806">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-806">1.0</span></span>|
|[<span data-ttu-id="c86d0-807">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-807">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-808">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-808">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-809">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-809">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-810">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-810">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="c86d0-811">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c86d0-811">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="c86d0-812">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c86d0-812">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-813">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-813">Read mode</span></span>

<span data-ttu-id="c86d0-814">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象，该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="c86d0-814">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-815">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-815">Compose mode</span></span>

<span data-ttu-id="c86d0-816">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer)器对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-816">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="c86d0-817">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-817">Type</span></span>

*   <span data-ttu-id="c86d0-818">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c86d0-818">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-819">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-819">Requirements</span></span>

|<span data-ttu-id="c86d0-820">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-820">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c86d0-821">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-822">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-822">1.0</span></span>|<span data-ttu-id="c86d0-823">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-823">1.7</span></span>|
|[<span data-ttu-id="c86d0-824">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-824">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-825">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-825">ReadItem</span></span>|<span data-ttu-id="c86d0-826">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-826">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-827">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-827">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-828">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-828">Read</span></span>|<span data-ttu-id="c86d0-829">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-829">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="c86d0-830">（可以为 null）定期：[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="c86d0-830">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="c86d0-831">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-831">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="c86d0-832">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-832">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c86d0-833">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-833">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c86d0-834">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-834">Read mode for meeting request items.</span></span>

<span data-ttu-id="c86d0-835">如果`recurrence`项目是系列中的一个系列或一个实例，则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-835">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c86d0-836">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="c86d0-836">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c86d0-837">`undefined`对于不是会议请求的邮件，将返回。</span><span class="sxs-lookup"><span data-stu-id="c86d0-837">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c86d0-838">注意：会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="c86d0-838">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c86d0-839">注意：如果定期对象为`null`，则表示该对象是单个约会的单个约会或会议请求，而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="c86d0-839">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-840">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-840">Read mode</span></span>

<span data-ttu-id="c86d0-841">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-841">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="c86d0-842">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="c86d0-842">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-843">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-843">Compose mode</span></span>

<span data-ttu-id="c86d0-844">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence)对象，该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-844">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="c86d0-845">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="c86d0-845">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-846">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-846">Type</span></span>

* [<span data-ttu-id="c86d0-847">循环</span><span class="sxs-lookup"><span data-stu-id="c86d0-847">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="c86d0-848">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-848">Requirement</span></span>|<span data-ttu-id="c86d0-849">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-850">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-851">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-851">1.7</span></span>|
|[<span data-ttu-id="c86d0-852">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-853">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-854">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-855">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-855">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c86d0-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c86d0-857">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c86d0-857">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c86d0-858">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-858">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-859">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-859">Read mode</span></span>

<span data-ttu-id="c86d0-860">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-860">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="c86d0-861">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-861">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-862">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-862">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-863">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-863">Compose mode</span></span>

<span data-ttu-id="c86d0-864">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-864">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="c86d0-865">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-865">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-866">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-866">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c86d0-867">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-867">Get 500 members maximum.</span></span>
- <span data-ttu-id="c86d0-868">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-868">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-869">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-869">Type</span></span>

*   <span data-ttu-id="c86d0-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-871">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-871">Requirements</span></span>

|<span data-ttu-id="c86d0-872">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-872">Requirement</span></span>|<span data-ttu-id="c86d0-873">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-874">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-875">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-875">1.0</span></span>|
|[<span data-ttu-id="c86d0-876">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-876">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-877">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-877">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-878">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-878">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-879">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-879">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="c86d0-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c86d0-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="c86d0-p135">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c86d0-p136">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-885">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-885">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-886">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-886">Type</span></span>

*   [<span data-ttu-id="c86d0-887">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c86d0-887">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c86d0-888">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-888">Requirements</span></span>

|<span data-ttu-id="c86d0-889">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-889">Requirement</span></span>|<span data-ttu-id="c86d0-890">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-891">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-892">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-892">1.0</span></span>|
|[<span data-ttu-id="c86d0-893">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-894">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-894">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-895">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-896">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-896">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-897">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-897">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c86d0-898">（可以为 null） Webcasts&seriesid： String</span><span class="sxs-lookup"><span data-stu-id="c86d0-898">(nullable) seriesId: String</span></span>

<span data-ttu-id="c86d0-899">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="c86d0-899">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c86d0-900">在 web 上的 Outlook 和桌面客户端中`seriesId` ，返回此项所属的父（系列）项的 Exchange web 服务（EWS） ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-900">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c86d0-901">但是，在 iOS 和 Android 中， `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-901">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-902">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="c86d0-902">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c86d0-903">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="c86d0-903">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c86d0-904">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="c86d0-904">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c86d0-905">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="c86d0-905">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c86d0-906">对于`seriesId`不包含`null`父项（如单个约会、系列项或会议请求）的项，该属性将返回， `undefined`对于不是会议请求的任何其他项，该属性返回。</span><span class="sxs-lookup"><span data-stu-id="c86d0-906">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c86d0-907">Type</span><span class="sxs-lookup"><span data-stu-id="c86d0-907">Type</span></span>

* <span data-ttu-id="c86d0-908">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-908">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-909">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-909">Requirements</span></span>

|<span data-ttu-id="c86d0-910">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-910">Requirement</span></span>|<span data-ttu-id="c86d0-911">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-911">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-912">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-912">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-913">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-913">1.7</span></span>|
|[<span data-ttu-id="c86d0-914">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-914">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-915">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-915">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-916">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-916">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-917">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-917">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-918">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-918">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c86d0-919">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c86d0-919">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c86d0-920">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c86d0-920">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c86d0-p139">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-923">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-923">Read mode</span></span>

<span data-ttu-id="c86d0-924">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-924">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-925">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-925">Compose mode</span></span>

<span data-ttu-id="c86d0-926">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-926">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c86d0-927">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c86d0-927">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c86d0-928">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="c86d0-928">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="c86d0-929">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-929">Type</span></span>

*   <span data-ttu-id="c86d0-930">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c86d0-930">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-931">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-931">Requirements</span></span>

|<span data-ttu-id="c86d0-932">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-932">Requirement</span></span>|<span data-ttu-id="c86d0-933">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-934">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-935">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-935">1.0</span></span>|
|[<span data-ttu-id="c86d0-936">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-936">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-937">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-938">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-938">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-939">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-939">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="c86d0-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c86d0-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="c86d0-941">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="c86d0-941">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c86d0-942">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="c86d0-942">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-943">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-943">Read mode</span></span>

<span data-ttu-id="c86d0-p140">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="c86d0-946">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-946">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-947">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-947">Compose mode</span></span>
<span data-ttu-id="c86d0-948">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-948">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-949">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-949">Type</span></span>

*   <span data-ttu-id="c86d0-950">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c86d0-950">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-951">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-951">Requirements</span></span>

|<span data-ttu-id="c86d0-952">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-952">Requirement</span></span>|<span data-ttu-id="c86d0-953">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-953">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-954">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-954">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-955">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-955">1.0</span></span>|
|[<span data-ttu-id="c86d0-956">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-956">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-957">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-957">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-958">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-958">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-959">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-959">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c86d0-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c86d0-961">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c86d0-961">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c86d0-962">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-962">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c86d0-963">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-963">Read mode</span></span>

<span data-ttu-id="c86d0-964">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-964">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="c86d0-965">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-965">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-966">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-966">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="c86d0-967">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-967">Compose mode</span></span>

<span data-ttu-id="c86d0-968">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-968">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="c86d0-969">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-969">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c86d0-970">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-970">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c86d0-971">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-971">Get 500 members maximum.</span></span>
- <span data-ttu-id="c86d0-972">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="c86d0-972">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c86d0-973">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-973">Type</span></span>

*   <span data-ttu-id="c86d0-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c86d0-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-975">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-975">Requirements</span></span>

|<span data-ttu-id="c86d0-976">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-976">Requirement</span></span>|<span data-ttu-id="c86d0-977">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-977">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-978">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-978">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-979">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-979">1.0</span></span>|
|[<span data-ttu-id="c86d0-980">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-980">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-981">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-981">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-982">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-982">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-983">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-983">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="c86d0-984">方法详细信息</span><span class="sxs-lookup"><span data-stu-id="c86d0-984">Method details</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c86d0-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c86d0-986">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c86d0-986">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c86d0-987">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="c86d0-987">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c86d0-988">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-988">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-989">参数</span><span class="sxs-lookup"><span data-stu-id="c86d0-989">Parameters</span></span>
|<span data-ttu-id="c86d0-990">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-990">Name</span></span>|<span data-ttu-id="c86d0-991">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-991">Type</span></span>|<span data-ttu-id="c86d0-992">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-992">Attributes</span></span>|<span data-ttu-id="c86d0-993">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-993">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c86d0-994">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-994">String</span></span>||<span data-ttu-id="c86d0-p144">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c86d0-997">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-997">String</span></span>||<span data-ttu-id="c86d0-p145">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c86d0-1000">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1000">Object</span></span>|<span data-ttu-id="c86d0-1001">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1002">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1003">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1003">Object</span></span>|<span data-ttu-id="c86d0-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1005">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c86d0-1006">布尔值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1006">Boolean</span></span>|<span data-ttu-id="c86d0-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1008">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1008">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1009">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1009">function</span></span>|<span data-ttu-id="c86d0-1010">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1011">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1011">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c86d0-1012">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1012">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c86d0-1013">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1013">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c86d0-1014">错误</span><span class="sxs-lookup"><span data-stu-id="c86d0-1014">Errors</span></span>

|<span data-ttu-id="c86d0-1015">错误代码</span><span class="sxs-lookup"><span data-stu-id="c86d0-1015">Error code</span></span>|<span data-ttu-id="c86d0-1016">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1016">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c86d0-1017">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1017">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c86d0-1018">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1018">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c86d0-1019">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1019">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1020">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1020">Requirements</span></span>

|<span data-ttu-id="c86d0-1021">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1021">Requirement</span></span>|<span data-ttu-id="c86d0-1022">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1023">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1024">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-1024">1.1</span></span>|
|[<span data-ttu-id="c86d0-1025">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1026">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1026">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-1027">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1028">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1028">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c86d0-1029">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1029">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="c86d0-1030">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1030">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="c86d0-1031">addFileAttachmentFromBase64Async （base64File，attachmentName，[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="c86d0-1031">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c86d0-1032">将 base64 编码中的文件作为附件添加到邮件或约会中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1032">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c86d0-1033">该`addFileAttachmentFromBase64Async`方法从 base64 编码中上载文件，并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1033">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="c86d0-1034">此方法返回 AsyncResult 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1034">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="c86d0-1035">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1035">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1036">参数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1036">Parameters</span></span>

|<span data-ttu-id="c86d0-1037">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1037">Name</span></span>|<span data-ttu-id="c86d0-1038">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1038">Type</span></span>|<span data-ttu-id="c86d0-1039">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1039">Attributes</span></span>|<span data-ttu-id="c86d0-1040">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1040">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="c86d0-1041">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1041">String</span></span>||<span data-ttu-id="c86d0-1042">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1042">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="c86d0-1043">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1043">String</span></span>||<span data-ttu-id="c86d0-p147">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c86d0-1046">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1046">Object</span></span>|<span data-ttu-id="c86d0-1047">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1048">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1049">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1049">Object</span></span>|<span data-ttu-id="c86d0-1050">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1051">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c86d0-1052">布尔值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1052">Boolean</span></span>|<span data-ttu-id="c86d0-1053">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1054">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1054">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1055">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1055">function</span></span>|<span data-ttu-id="c86d0-1056">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1057">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c86d0-1058">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1058">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c86d0-1059">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1059">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c86d0-1060">错误</span><span class="sxs-lookup"><span data-stu-id="c86d0-1060">Errors</span></span>

|<span data-ttu-id="c86d0-1061">错误代码</span><span class="sxs-lookup"><span data-stu-id="c86d0-1061">Error code</span></span>|<span data-ttu-id="c86d0-1062">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1062">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c86d0-1063">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1063">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c86d0-1064">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1064">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c86d0-1065">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1065">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1066">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1066">Requirements</span></span>

|<span data-ttu-id="c86d0-1067">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1067">Requirement</span></span>|<span data-ttu-id="c86d0-1068">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1069">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1070">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-1070">1.8</span></span>|
|[<span data-ttu-id="c86d0-1071">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1072">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1072">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-1073">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1074">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1074">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c86d0-1075">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1075">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c86d0-1076">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1076">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c86d0-1077">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1077">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c86d0-1078">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1078">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1079">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1079">Parameters</span></span>

| <span data-ttu-id="c86d0-1080">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1080">Name</span></span> | <span data-ttu-id="c86d0-1081">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1081">Type</span></span> | <span data-ttu-id="c86d0-1082">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1082">Attributes</span></span> | <span data-ttu-id="c86d0-1083">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1083">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c86d0-1084">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c86d0-1084">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c86d0-1085">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1085">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c86d0-1086">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1086">Function</span></span> || <span data-ttu-id="c86d0-p148">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c86d0-1090">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1090">Object</span></span> | <span data-ttu-id="c86d0-1091">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1091">&lt;optional&gt;</span></span> | <span data-ttu-id="c86d0-1092">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1092">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c86d0-1093">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1093">Object</span></span> | <span data-ttu-id="c86d0-1094">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1094">&lt;optional&gt;</span></span> | <span data-ttu-id="c86d0-1095">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1095">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c86d0-1096">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1096">function</span></span>| <span data-ttu-id="c86d0-1097">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1098">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1099">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1099">Requirements</span></span>

|<span data-ttu-id="c86d0-1100">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1100">Requirement</span></span>| <span data-ttu-id="c86d0-1101">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1101">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1102">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1102">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c86d0-1103">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-1103">1.7</span></span> |
|[<span data-ttu-id="c86d0-1104">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1104">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c86d0-1105">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1105">ReadItem</span></span> |
|[<span data-ttu-id="c86d0-1106">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1106">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c86d0-1107">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1107">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="c86d0-1108">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1108">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c86d0-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c86d0-1110">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1110">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c86d0-p149">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c86d0-1114">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1114">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c86d0-1115">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1115">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1116">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1116">Parameters</span></span>

|<span data-ttu-id="c86d0-1117">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1117">Name</span></span>|<span data-ttu-id="c86d0-1118">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1118">Type</span></span>|<span data-ttu-id="c86d0-1119">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1119">Attributes</span></span>|<span data-ttu-id="c86d0-1120">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1120">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c86d0-1121">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1121">String</span></span>||<span data-ttu-id="c86d0-p150">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c86d0-1124">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1124">String</span></span>||<span data-ttu-id="c86d0-1125">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1125">The subject of the item to be attached.</span></span> <span data-ttu-id="c86d0-1126">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1126">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c86d0-1127">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1127">Object</span></span>|<span data-ttu-id="c86d0-1128">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1129">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1129">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1130">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1130">Object</span></span>|<span data-ttu-id="c86d0-1131">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1131">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1132">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1132">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1133">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1133">function</span></span>|<span data-ttu-id="c86d0-1134">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1134">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1135">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c86d0-1136">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1136">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c86d0-1137">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1137">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c86d0-1138">错误</span><span class="sxs-lookup"><span data-stu-id="c86d0-1138">Errors</span></span>

|<span data-ttu-id="c86d0-1139">错误代码</span><span class="sxs-lookup"><span data-stu-id="c86d0-1139">Error code</span></span>|<span data-ttu-id="c86d0-1140">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1140">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c86d0-1141">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1141">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1142">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1142">Requirements</span></span>

|<span data-ttu-id="c86d0-1143">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1143">Requirement</span></span>|<span data-ttu-id="c86d0-1144">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1144">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1146">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-1146">1.1</span></span>|
|[<span data-ttu-id="c86d0-1147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1148">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1148">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-1149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1150">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-1151">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1151">Example</span></span>

<span data-ttu-id="c86d0-1152">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1152">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="close"></a><span data-ttu-id="c86d0-1153">close()</span><span class="sxs-lookup"><span data-stu-id="c86d0-1153">close()</span></span>

<span data-ttu-id="c86d0-1154">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1154">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c86d0-p152">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1157">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1157">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c86d0-1158">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1158">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-1159">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1159">Requirements</span></span>

|<span data-ttu-id="c86d0-1160">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1160">Requirement</span></span>|<span data-ttu-id="c86d0-1161">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1163">1.3</span><span class="sxs-lookup"><span data-stu-id="c86d0-1163">1.3</span></span>|
|[<span data-ttu-id="c86d0-1164">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1165">受限</span><span class="sxs-lookup"><span data-stu-id="c86d0-1165">Restricted</span></span>|
|[<span data-ttu-id="c86d0-1166">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1167">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1167">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="c86d0-1168">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1168">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="c86d0-1169">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1169">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1170">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1170">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c86d0-1171">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1171">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c86d0-1172">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1172">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c86d0-p153">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1176">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1176">Parameters</span></span>

|<span data-ttu-id="c86d0-1177">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1177">Name</span></span>|<span data-ttu-id="c86d0-1178">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1178">Type</span></span>|<span data-ttu-id="c86d0-1179">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1179">Attributes</span></span>|<span data-ttu-id="c86d0-1180">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1180">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c86d0-1181">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1181">String &#124; Object</span></span>||<span data-ttu-id="c86d0-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c86d0-1184">**或**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1184">**OR**</span></span><br/><span data-ttu-id="c86d0-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c86d0-1187">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1187">String</span></span>|<span data-ttu-id="c86d0-1188">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1188">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c86d0-1191">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1191">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c86d0-1192">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1192">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1193">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1193">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c86d0-1194">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1194">String</span></span>||<span data-ttu-id="c86d0-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c86d0-1197">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1197">String</span></span>||<span data-ttu-id="c86d0-1198">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1198">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c86d0-1199">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1199">String</span></span>||<span data-ttu-id="c86d0-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c86d0-1202">布尔</span><span class="sxs-lookup"><span data-stu-id="c86d0-1202">Boolean</span></span>||<span data-ttu-id="c86d0-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c86d0-1205">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1205">String</span></span>||<span data-ttu-id="c86d0-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1209">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1209">function</span></span>|<span data-ttu-id="c86d0-1210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1210">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1211">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1212">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1212">Requirements</span></span>

|<span data-ttu-id="c86d0-1213">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1213">Requirement</span></span>|<span data-ttu-id="c86d0-1214">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1216">1.0</span></span>|
|[<span data-ttu-id="c86d0-1217">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1218">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1220">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1220">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c86d0-1221">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1221">Examples</span></span>

<span data-ttu-id="c86d0-1222">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1222">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c86d0-1223">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1223">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c86d0-1224">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1224">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c86d0-1225">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1225">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="c86d0-1226">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1226">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="c86d0-1227">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1227">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="c86d0-1228">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1228">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="c86d0-1229">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1229">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1230">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c86d0-1231">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1231">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c86d0-1232">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1232">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c86d0-p161">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1236">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1236">Parameters</span></span>

|<span data-ttu-id="c86d0-1237">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1237">Name</span></span>|<span data-ttu-id="c86d0-1238">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1238">Type</span></span>|<span data-ttu-id="c86d0-1239">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1239">Attributes</span></span>|<span data-ttu-id="c86d0-1240">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1240">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c86d0-1241">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1241">String &#124; Object</span></span>||<span data-ttu-id="c86d0-p162">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c86d0-1244">**或**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1244">**OR**</span></span><br/><span data-ttu-id="c86d0-p163">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c86d0-1247">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1247">String</span></span>|<span data-ttu-id="c86d0-1248">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1248">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-p164">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c86d0-1251">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1251">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c86d0-1252">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1252">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1253">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1253">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c86d0-1254">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1254">String</span></span>||<span data-ttu-id="c86d0-p165">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c86d0-1257">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1257">String</span></span>||<span data-ttu-id="c86d0-1258">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1258">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c86d0-1259">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1259">String</span></span>||<span data-ttu-id="c86d0-p166">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c86d0-1262">布尔</span><span class="sxs-lookup"><span data-stu-id="c86d0-1262">Boolean</span></span>||<span data-ttu-id="c86d0-p167">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c86d0-1265">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1265">String</span></span>||<span data-ttu-id="c86d0-p168">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1269">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1269">function</span></span>|<span data-ttu-id="c86d0-1270">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1271">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1272">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1272">Requirements</span></span>

|<span data-ttu-id="c86d0-1273">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1273">Requirement</span></span>|<span data-ttu-id="c86d0-1274">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1274">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1275">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1275">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1276">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1276">1.0</span></span>|
|[<span data-ttu-id="c86d0-1277">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1277">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1278">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1279">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1279">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1280">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1280">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c86d0-1281">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1281">Examples</span></span>

<span data-ttu-id="c86d0-1282">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1282">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c86d0-1283">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1283">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c86d0-1284">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1284">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c86d0-1285">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1285">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="c86d0-1286">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1286">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="c86d0-1287">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1287">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="c86d0-1288">getAllInternetHeadersAsync （[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="c86d0-1288">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="c86d0-1289">以字符串形式获取邮件的所有 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1289">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="c86d0-1290">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1290">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1291">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1291">Parameters</span></span>

|<span data-ttu-id="c86d0-1292">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1292">Name</span></span>|<span data-ttu-id="c86d0-1293">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1293">Type</span></span>|<span data-ttu-id="c86d0-1294">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1294">Attributes</span></span>|<span data-ttu-id="c86d0-1295">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1295">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c86d0-1296">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1296">Object</span></span>|<span data-ttu-id="c86d0-1297">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1297">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1298">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1298">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1299">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1299">Object</span></span>|<span data-ttu-id="c86d0-1300">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1301">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1301">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1302">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1302">function</span></span>|<span data-ttu-id="c86d0-1303">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1304">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1304">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="c86d0-1305">在成功的情况下，internet 标头数据在 asyncResult 属性中以字符串的形式提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1305">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="c86d0-1306">有关返回的字符串值的格式设置信息，请参阅[RFC 2183](https://tools.ietf.org/html/rfc2183) 。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1306">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="c86d0-1307">如果调用失败，asyncResult 属性将包含错误代码和失败原因。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1307">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1308">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1308">Requirements</span></span>

|<span data-ttu-id="c86d0-1309">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1309">Requirement</span></span>|<span data-ttu-id="c86d0-1310">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1312">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-1312">1.8</span></span>|
|[<span data-ttu-id="c86d0-1313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1314">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1316">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1316">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1317">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1317">Returns:</span></span>

<span data-ttu-id="c86d0-1318">作为字符串的 internet 标头数据，根据[RFC 2183](https://tools.ietf.org/html/rfc2183)格式化。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1318">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="c86d0-1319">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1319">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1320">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1320">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="c86d0-1321">getAttachmentContentAsync （attachmentId，[options]，[callback]）→ [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c86d0-1321">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="c86d0-1322">从邮件或约会中获取指定附件并将其作为`AttachmentContent`对象返回。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1322">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="c86d0-1323">该`getAttachmentContentAsync`方法从项目中获取具有指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1323">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c86d0-1324">作为一种最佳做法，您应使用标识符在与`getAttachmentsAsync` or `item.attachments`调用一起检索到会话的同一会话中检索附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1324">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="c86d0-1325">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1325">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c86d0-1326">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1326">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1327">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1327">Parameters</span></span>

|<span data-ttu-id="c86d0-1328">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1328">Name</span></span>|<span data-ttu-id="c86d0-1329">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1329">Type</span></span>|<span data-ttu-id="c86d0-1330">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1330">Attributes</span></span>|<span data-ttu-id="c86d0-1331">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1331">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c86d0-1332">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1332">String</span></span>||<span data-ttu-id="c86d0-1333">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1333">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="c86d0-1334">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1334">Object</span></span>|<span data-ttu-id="c86d0-1335">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1335">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1336">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1336">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1337">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1337">Object</span></span>|<span data-ttu-id="c86d0-1338">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1338">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1339">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1339">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1340">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1340">function</span></span>|<span data-ttu-id="c86d0-1341">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1342">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1343">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1343">Requirements</span></span>

|<span data-ttu-id="c86d0-1344">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1344">Requirement</span></span>|<span data-ttu-id="c86d0-1345">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1346">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1347">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-1347">1.8</span></span>|
|[<span data-ttu-id="c86d0-1348">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1349">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1350">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1351">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1351">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1352">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1352">Returns:</span></span>

<span data-ttu-id="c86d0-1353">类型： [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c86d0-1353">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1354">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1354">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c86d0-1355">getAttachmentsAsync （[options]，[callback]）→ Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-1355">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c86d0-1356">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1356">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c86d0-1357">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1357">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1358">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1358">Parameters</span></span>

|<span data-ttu-id="c86d0-1359">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1359">Name</span></span>|<span data-ttu-id="c86d0-1360">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1360">Type</span></span>|<span data-ttu-id="c86d0-1361">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1361">Attributes</span></span>|<span data-ttu-id="c86d0-1362">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1362">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c86d0-1363">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1363">Object</span></span>|<span data-ttu-id="c86d0-1364">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1364">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1365">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1365">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1366">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1366">Object</span></span>|<span data-ttu-id="c86d0-1367">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1367">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1368">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1368">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1369">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1369">function</span></span>|<span data-ttu-id="c86d0-1370">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1371">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1371">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1372">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1372">Requirements</span></span>

|<span data-ttu-id="c86d0-1373">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1373">Requirement</span></span>|<span data-ttu-id="c86d0-1374">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1374">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1375">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1376">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-1376">1.8</span></span>|
|[<span data-ttu-id="c86d0-1377">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1378">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1379">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1380">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1380">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1381">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1381">Returns:</span></span>

<span data-ttu-id="c86d0-1382">类型： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c86d0-1382">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1383">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1383">Example</span></span>

<span data-ttu-id="c86d0-1384">下面的示例将生成一个 HTML 字符串，其中包含当前项目上所有附件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1384">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c86d0-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c86d0-1386">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1386">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1387">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1387">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-1388">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1388">Requirements</span></span>

|<span data-ttu-id="c86d0-1389">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1389">Requirement</span></span>|<span data-ttu-id="c86d0-1390">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1390">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1391">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1392">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1392">1.0</span></span>|
|[<span data-ttu-id="c86d0-1393">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1393">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1394">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1395">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1395">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1396">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1396">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1397">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1397">Returns:</span></span>

<span data-ttu-id="c86d0-1398">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c86d0-1398">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1399">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1399">Example</span></span>

<span data-ttu-id="c86d0-1400">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1400">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c86d0-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c86d0-1402">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1402">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1403">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1403">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1404">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1404">Parameters</span></span>

|<span data-ttu-id="c86d0-1405">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1405">Name</span></span>|<span data-ttu-id="c86d0-1406">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1406">Type</span></span>|<span data-ttu-id="c86d0-1407">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1407">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c86d0-1408">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c86d0-1408">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="c86d0-1409">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1409">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1410">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1410">Requirements</span></span>

|<span data-ttu-id="c86d0-1411">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1411">Requirement</span></span>|<span data-ttu-id="c86d0-1412">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1412">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1414">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1414">1.0</span></span>|
|[<span data-ttu-id="c86d0-1415">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1416">受限</span><span class="sxs-lookup"><span data-stu-id="c86d0-1416">Restricted</span></span>|
|[<span data-ttu-id="c86d0-1417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1418">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1418">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1419">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1419">Returns:</span></span>

<span data-ttu-id="c86d0-1420">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1420">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c86d0-1421">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1421">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c86d0-1422">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1422">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c86d0-1423">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1423">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c86d0-1424">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1424">Value of `entityType`</span></span>|<span data-ttu-id="c86d0-1425">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1425">Type of objects in returned array</span></span>|<span data-ttu-id="c86d0-1426">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1426">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c86d0-1427">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1427">String</span></span>|<span data-ttu-id="c86d0-1428">**受限**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1428">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c86d0-1429">Contact</span><span class="sxs-lookup"><span data-stu-id="c86d0-1429">Contact</span></span>|<span data-ttu-id="c86d0-1430">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1430">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c86d0-1431">String</span><span class="sxs-lookup"><span data-stu-id="c86d0-1431">String</span></span>|<span data-ttu-id="c86d0-1432">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1432">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c86d0-1433">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c86d0-1433">MeetingSuggestion</span></span>|<span data-ttu-id="c86d0-1434">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1434">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c86d0-1435">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c86d0-1435">PhoneNumber</span></span>|<span data-ttu-id="c86d0-1436">**受限**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1436">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c86d0-1437">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c86d0-1437">TaskSuggestion</span></span>|<span data-ttu-id="c86d0-1438">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1438">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c86d0-1439">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1439">String</span></span>|<span data-ttu-id="c86d0-1440">**受限**</span><span class="sxs-lookup"><span data-stu-id="c86d0-1440">**Restricted**</span></span>|

<span data-ttu-id="c86d0-1441">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c86d0-1441">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1442">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1442">Example</span></span>

<span data-ttu-id="c86d0-1443">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1443">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c86d0-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c86d0-1445">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1445">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1446">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1446">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c86d0-1447">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1447">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1448">参数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1448">Parameters</span></span>

|<span data-ttu-id="c86d0-1449">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1449">Name</span></span>|<span data-ttu-id="c86d0-1450">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1450">Type</span></span>|<span data-ttu-id="c86d0-1451">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1451">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c86d0-1452">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1452">String</span></span>|<span data-ttu-id="c86d0-1453">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1453">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1454">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1454">Requirements</span></span>

|<span data-ttu-id="c86d0-1455">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1455">Requirement</span></span>|<span data-ttu-id="c86d0-1456">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1456">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1457">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1457">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1458">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1458">1.0</span></span>|
|[<span data-ttu-id="c86d0-1459">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1459">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1460">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1460">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1461">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1461">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1462">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1462">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1463">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1463">Returns:</span></span>

<span data-ttu-id="c86d0-p174">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c86d0-1466">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c86d0-1466">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="c86d0-1467">Office.context.mailbox.item.getinitializationcontextasync （[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="c86d0-1467">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="c86d0-1468">获取[通过可操作邮件激活](/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1468">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1469">仅 Outlook 2016 或更高版本（高于16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1469">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1470">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1470">Parameters</span></span>

|<span data-ttu-id="c86d0-1471">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1471">Name</span></span>|<span data-ttu-id="c86d0-1472">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1472">Type</span></span>|<span data-ttu-id="c86d0-1473">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1473">Attributes</span></span>|<span data-ttu-id="c86d0-1474">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1474">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c86d0-1475">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1475">Object</span></span>|<span data-ttu-id="c86d0-1476">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1476">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1477">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1477">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1478">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1478">Object</span></span>|<span data-ttu-id="c86d0-1479">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1479">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1480">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1480">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1481">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1481">function</span></span>|<span data-ttu-id="c86d0-1482">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1482">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1483">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1483">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c86d0-1484">如果成功，初始化数据在`asyncResult.value`属性中提供为字符串。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1484">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="c86d0-1485">如果没有初始化上下文，该`asyncResult`对象将包含其`Error` `code`属性设置为`9020`的对象及其`name`属性设置为。 `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="c86d0-1485">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1486">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1486">Requirements</span></span>

|<span data-ttu-id="c86d0-1487">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1487">Requirement</span></span>|<span data-ttu-id="c86d0-1488">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1489">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1490">预览</span><span class="sxs-lookup"><span data-stu-id="c86d0-1490">Preview</span></span>|
|[<span data-ttu-id="c86d0-1491">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1492">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1493">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1494">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1494">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-1495">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1495">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="c86d0-1496">getItemIdAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="c86d0-1496">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="c86d0-1497">异步获取已保存项的 ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1497">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="c86d0-1498">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1498">Compose mode only.</span></span>

<span data-ttu-id="c86d0-1499">调用此方法时，此方法通过回调方法返回项 ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1499">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1500">如果你的外接程序`getItemIdAsync`对撰写模式中的项（例如，要获取`itemId`使用 EWS 或 REST API 的使用）调用，请注意，当 Outlook 处于缓存模式下时，可能需要一段时间才能将项目同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1500">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="c86d0-1501">在同步项目之前，无法识别`itemId`该项目并使用它将返回错误。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1501">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1502">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1502">Parameters</span></span>

|<span data-ttu-id="c86d0-1503">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1503">Name</span></span>|<span data-ttu-id="c86d0-1504">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1504">Type</span></span>|<span data-ttu-id="c86d0-1505">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1505">Attributes</span></span>|<span data-ttu-id="c86d0-1506">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1506">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c86d0-1507">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1507">Object</span></span>|<span data-ttu-id="c86d0-1508">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1508">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1509">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1509">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1510">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1510">Object</span></span>|<span data-ttu-id="c86d0-1511">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1511">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1512">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1512">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1513">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1513">function</span></span>||<span data-ttu-id="c86d0-1514">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1514">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c86d0-1515">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1515">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c86d0-1516">错误</span><span class="sxs-lookup"><span data-stu-id="c86d0-1516">Errors</span></span>

|<span data-ttu-id="c86d0-1517">错误代码</span><span class="sxs-lookup"><span data-stu-id="c86d0-1517">Error code</span></span>|<span data-ttu-id="c86d0-1518">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1518">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="c86d0-1519">在保存项目之前，无法检索此 id。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1519">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1520">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1520">Requirements</span></span>

|<span data-ttu-id="c86d0-1521">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1521">Requirement</span></span>|<span data-ttu-id="c86d0-1522">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1522">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1523">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1524">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-1524">1.8</span></span>|
|[<span data-ttu-id="c86d0-1525">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1526">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1527">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1528">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1528">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c86d0-1529">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1529">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c86d0-1530">下面的示例演示传递给回调函数`result`的参数的结构。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1530">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="c86d0-1531">`value`属性包含项 ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1531">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="c86d0-1532">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1532">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c86d0-1533">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1533">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1534">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1534">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c86d0-p178">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c86d0-1538">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1538">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c86d0-1539">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1539">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c86d0-p179">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-1543">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1543">Requirements</span></span>

|<span data-ttu-id="c86d0-1544">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1544">Requirement</span></span>|<span data-ttu-id="c86d0-1545">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1546">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1547">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1547">1.0</span></span>|
|[<span data-ttu-id="c86d0-1548">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1549">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1550">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1551">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1551">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1552">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1552">Returns:</span></span>

<span data-ttu-id="c86d0-p180">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c86d0-1555">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="c86d0-1555">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c86d0-1556">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1556">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c86d0-1557">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1557">Example</span></span>

<span data-ttu-id="c86d0-1558">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1558">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c86d0-1559">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c86d0-1560">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1560">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1561">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1561">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c86d0-1562">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1562">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c86d0-p181">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1565">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1565">Parameters</span></span>

|<span data-ttu-id="c86d0-1566">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1566">Name</span></span>|<span data-ttu-id="c86d0-1567">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1567">Type</span></span>|<span data-ttu-id="c86d0-1568">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1568">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c86d0-1569">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1569">String</span></span>|<span data-ttu-id="c86d0-1570">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1570">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1571">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1571">Requirements</span></span>

|<span data-ttu-id="c86d0-1572">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1572">Requirement</span></span>|<span data-ttu-id="c86d0-1573">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1573">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1574">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1575">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1575">1.0</span></span>|
|[<span data-ttu-id="c86d0-1576">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1577">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1578">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1579">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1579">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1580">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1580">Returns:</span></span>

<span data-ttu-id="c86d0-1581">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1581">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="c86d0-1582">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c86d0-1582">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1583">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1583">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c86d0-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c86d0-1585">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1585">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c86d0-p182">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回空字符串。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p182">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1588">参数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1588">Parameters</span></span>

|<span data-ttu-id="c86d0-1589">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1589">Name</span></span>|<span data-ttu-id="c86d0-1590">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1590">Type</span></span>|<span data-ttu-id="c86d0-1591">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1591">Attributes</span></span>|<span data-ttu-id="c86d0-1592">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1592">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c86d0-1593">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c86d0-1593">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c86d0-p183">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c86d0-1597">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1597">Object</span></span>|<span data-ttu-id="c86d0-1598">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1599">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1599">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1600">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1600">Object</span></span>|<span data-ttu-id="c86d0-1601">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1602">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1602">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1603">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1603">function</span></span>||<span data-ttu-id="c86d0-1604">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1604">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c86d0-1605">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1605">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c86d0-1606">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1606">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1607">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1607">Requirements</span></span>

|<span data-ttu-id="c86d0-1608">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1608">Requirement</span></span>|<span data-ttu-id="c86d0-1609">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1609">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1610">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1611">1.2</span><span class="sxs-lookup"><span data-stu-id="c86d0-1611">1.2</span></span>|
|[<span data-ttu-id="c86d0-1612">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1613">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1614">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1615">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1615">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1616">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1616">Returns:</span></span>

<span data-ttu-id="c86d0-1617">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1617">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="c86d0-1618">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1618">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1619">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1619">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c86d0-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c86d0-1621">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1621">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="c86d0-1622">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1622">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1623">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1623">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-1624">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1624">Requirements</span></span>

|<span data-ttu-id="c86d0-1625">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1625">Requirement</span></span>|<span data-ttu-id="c86d0-1626">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1626">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1628">1.6</span><span class="sxs-lookup"><span data-stu-id="c86d0-1628">1.6</span></span>|
|[<span data-ttu-id="c86d0-1629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1630">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1632">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1632">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1633">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1633">Returns:</span></span>

<span data-ttu-id="c86d0-1634">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c86d0-1634">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1635">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1635">Example</span></span>

<span data-ttu-id="c86d0-1636">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1636">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c86d0-1637">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c86d0-1637">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c86d0-p186">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1640">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c86d0-p187">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c86d0-1644">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1644">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c86d0-1645">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1645">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c86d0-p188">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c86d0-1649">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1649">Requirements</span></span>

|<span data-ttu-id="c86d0-1650">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1650">Requirement</span></span>|<span data-ttu-id="c86d0-1651">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1651">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1652">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1653">1.6</span><span class="sxs-lookup"><span data-stu-id="c86d0-1653">1.6</span></span>|
|[<span data-ttu-id="c86d0-1654">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1655">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1656">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1657">阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1657">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c86d0-1658">返回：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1658">Returns:</span></span>

<span data-ttu-id="c86d0-p189">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c86d0-1661">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1661">Example</span></span>

<span data-ttu-id="c86d0-1662">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1662">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="c86d0-1663">getSharedPropertiesAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="c86d0-1663">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="c86d0-1664">获取共享文件夹、日历或邮箱中的所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1664">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1665">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1665">Parameters</span></span>

|<span data-ttu-id="c86d0-1666">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1666">Name</span></span>|<span data-ttu-id="c86d0-1667">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1667">Type</span></span>|<span data-ttu-id="c86d0-1668">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1668">Attributes</span></span>|<span data-ttu-id="c86d0-1669">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1669">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c86d0-1670">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1670">Object</span></span>|<span data-ttu-id="c86d0-1671">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1671">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1672">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1672">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1673">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1673">Object</span></span>|<span data-ttu-id="c86d0-1674">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1674">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1675">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1675">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1676">function</span><span class="sxs-lookup"><span data-stu-id="c86d0-1676">function</span></span>||<span data-ttu-id="c86d0-1677">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c86d0-1678">共享属性作为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`属性中的对象提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1678">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c86d0-1679">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1679">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1680">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1680">Requirements</span></span>

|<span data-ttu-id="c86d0-1681">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1681">Requirement</span></span>|<span data-ttu-id="c86d0-1682">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1682">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1683">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1684">1.8</span><span class="sxs-lookup"><span data-stu-id="c86d0-1684">1.8</span></span>|
|[<span data-ttu-id="c86d0-1685">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1686">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1687">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1688">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1688">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-1689">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1689">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c86d0-1690">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1690">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c86d0-1691">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1691">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c86d0-p191">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1695">参数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1695">Parameters</span></span>

|<span data-ttu-id="c86d0-1696">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1696">Name</span></span>|<span data-ttu-id="c86d0-1697">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1697">Type</span></span>|<span data-ttu-id="c86d0-1698">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1698">Attributes</span></span>|<span data-ttu-id="c86d0-1699">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1699">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c86d0-1700">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1700">function</span></span>||<span data-ttu-id="c86d0-1701">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c86d0-1702">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1702">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c86d0-1703">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1703">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c86d0-1704">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1704">Object</span></span>|<span data-ttu-id="c86d0-1705">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1705">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1706">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1706">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c86d0-1707">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1707">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1708">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1708">Requirements</span></span>

|<span data-ttu-id="c86d0-1709">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1709">Requirement</span></span>|<span data-ttu-id="c86d0-1710">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1710">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1711">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1712">1.0</span><span class="sxs-lookup"><span data-stu-id="c86d0-1712">1.0</span></span>|
|[<span data-ttu-id="c86d0-1713">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1714">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1714">ReadItem</span></span>|
|[<span data-ttu-id="c86d0-1715">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1716">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1716">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-1717">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1717">Example</span></span>

<span data-ttu-id="c86d0-p194">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c86d0-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c86d0-1722">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1722">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c86d0-1723">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1723">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c86d0-1724">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1724">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="c86d0-1725">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1725">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c86d0-1726">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1726">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1727">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1727">Parameters</span></span>

|<span data-ttu-id="c86d0-1728">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1728">Name</span></span>|<span data-ttu-id="c86d0-1729">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1729">Type</span></span>|<span data-ttu-id="c86d0-1730">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1730">Attributes</span></span>|<span data-ttu-id="c86d0-1731">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1731">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c86d0-1732">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1732">String</span></span>||<span data-ttu-id="c86d0-1733">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1733">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="c86d0-1734">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1734">Object</span></span>|<span data-ttu-id="c86d0-1735">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1735">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1736">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1736">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1737">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1737">Object</span></span>|<span data-ttu-id="c86d0-1738">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1738">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1739">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1739">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1740">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1740">function</span></span>|<span data-ttu-id="c86d0-1741">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1741">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1742">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1742">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c86d0-1743">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1743">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c86d0-1744">错误</span><span class="sxs-lookup"><span data-stu-id="c86d0-1744">Errors</span></span>

|<span data-ttu-id="c86d0-1745">错误代码</span><span class="sxs-lookup"><span data-stu-id="c86d0-1745">Error code</span></span>|<span data-ttu-id="c86d0-1746">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1746">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c86d0-1747">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1747">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1748">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1748">Requirements</span></span>

|<span data-ttu-id="c86d0-1749">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1749">Requirement</span></span>|<span data-ttu-id="c86d0-1750">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1750">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1751">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1752">1.1</span><span class="sxs-lookup"><span data-stu-id="c86d0-1752">1.1</span></span>|
|[<span data-ttu-id="c86d0-1753">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1754">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1754">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-1755">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1756">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1756">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-1757">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1757">Example</span></span>

<span data-ttu-id="c86d0-1758">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1758">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c86d0-1759">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c86d0-1759">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c86d0-1760">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1760">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c86d0-1761">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1761">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1762">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1762">Parameters</span></span>

| <span data-ttu-id="c86d0-1763">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1763">Name</span></span> | <span data-ttu-id="c86d0-1764">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1764">Type</span></span> | <span data-ttu-id="c86d0-1765">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1765">Attributes</span></span> | <span data-ttu-id="c86d0-1766">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1766">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c86d0-1767">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c86d0-1767">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c86d0-1768">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1768">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c86d0-1769">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1769">Object</span></span> | <span data-ttu-id="c86d0-1770">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1770">&lt;optional&gt;</span></span> | <span data-ttu-id="c86d0-1771">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1771">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c86d0-1772">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1772">Object</span></span> | <span data-ttu-id="c86d0-1773">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1773">&lt;optional&gt;</span></span> | <span data-ttu-id="c86d0-1774">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1774">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c86d0-1775">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1775">function</span></span>| <span data-ttu-id="c86d0-1776">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1776">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1777">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1778">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1778">Requirements</span></span>

|<span data-ttu-id="c86d0-1779">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1779">Requirement</span></span>| <span data-ttu-id="c86d0-1780">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1780">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1781">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c86d0-1782">1.7</span><span class="sxs-lookup"><span data-stu-id="c86d0-1782">1.7</span></span> |
|[<span data-ttu-id="c86d0-1783">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c86d0-1784">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1784">ReadItem</span></span> |
|[<span data-ttu-id="c86d0-1785">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c86d0-1786">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c86d0-1786">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="c86d0-1787">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c86d0-1787">saveAsync([options], callback)</span></span>

<span data-ttu-id="c86d0-1788">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1788">Asynchronously saves an item.</span></span>

<span data-ttu-id="c86d0-1789">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1789">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="c86d0-1790">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1790">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="c86d0-1791">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1791">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1792">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1792">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c86d0-1793">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1793">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c86d0-p198">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c86d0-1797">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="c86d0-1797">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c86d0-1798">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1798">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="c86d0-1799">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1799">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="c86d0-1800">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1800">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="c86d0-1801">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1801">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1802">Parameters</span><span class="sxs-lookup"><span data-stu-id="c86d0-1802">Parameters</span></span>

|<span data-ttu-id="c86d0-1803">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1803">Name</span></span>|<span data-ttu-id="c86d0-1804">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1804">Type</span></span>|<span data-ttu-id="c86d0-1805">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1805">Attributes</span></span>|<span data-ttu-id="c86d0-1806">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1806">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c86d0-1807">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1807">Object</span></span>|<span data-ttu-id="c86d0-1808">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1808">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1809">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1809">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1810">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1810">Object</span></span>|<span data-ttu-id="c86d0-1811">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1811">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1812">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1812">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1813">函数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1813">function</span></span>||<span data-ttu-id="c86d0-1814">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1814">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c86d0-1815">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1815">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1816">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1816">Requirements</span></span>

|<span data-ttu-id="c86d0-1817">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1817">Requirement</span></span>|<span data-ttu-id="c86d0-1818">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1818">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1819">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1819">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1820">1.3</span><span class="sxs-lookup"><span data-stu-id="c86d0-1820">1.3</span></span>|
|[<span data-ttu-id="c86d0-1821">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1821">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1822">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1822">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-1823">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1823">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1824">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1824">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c86d0-1825">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1825">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c86d0-p200">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c86d0-1828">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c86d0-1828">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c86d0-1829">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1829">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c86d0-p201">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c86d0-1833">参数</span><span class="sxs-lookup"><span data-stu-id="c86d0-1833">Parameters</span></span>

|<span data-ttu-id="c86d0-1834">名称</span><span class="sxs-lookup"><span data-stu-id="c86d0-1834">Name</span></span>|<span data-ttu-id="c86d0-1835">类型</span><span class="sxs-lookup"><span data-stu-id="c86d0-1835">Type</span></span>|<span data-ttu-id="c86d0-1836">属性</span><span class="sxs-lookup"><span data-stu-id="c86d0-1836">Attributes</span></span>|<span data-ttu-id="c86d0-1837">说明</span><span class="sxs-lookup"><span data-stu-id="c86d0-1837">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c86d0-1838">字符串</span><span class="sxs-lookup"><span data-stu-id="c86d0-1838">String</span></span>||<span data-ttu-id="c86d0-p202">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="c86d0-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c86d0-1842">Object</span><span class="sxs-lookup"><span data-stu-id="c86d0-1842">Object</span></span>|<span data-ttu-id="c86d0-1843">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1843">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1844">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1844">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c86d0-1845">对象</span><span class="sxs-lookup"><span data-stu-id="c86d0-1845">Object</span></span>|<span data-ttu-id="c86d0-1846">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1846">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1847">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1847">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c86d0-1848">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c86d0-1848">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c86d0-1849">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c86d0-1849">&lt;optional&gt;</span></span>|<span data-ttu-id="c86d0-1850">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1850">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="c86d0-1851">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1851">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c86d0-1852">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1852">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="c86d0-1853">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1853">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c86d0-1854">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1854">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c86d0-1855">function</span><span class="sxs-lookup"><span data-stu-id="c86d0-1855">function</span></span>||<span data-ttu-id="c86d0-1856">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c86d0-1856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c86d0-1857">Requirements</span><span class="sxs-lookup"><span data-stu-id="c86d0-1857">Requirements</span></span>

|<span data-ttu-id="c86d0-1858">要求</span><span class="sxs-lookup"><span data-stu-id="c86d0-1858">Requirement</span></span>|<span data-ttu-id="c86d0-1859">值</span><span class="sxs-lookup"><span data-stu-id="c86d0-1859">Value</span></span>|
|---|---|
|[<span data-ttu-id="c86d0-1860">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c86d0-1860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c86d0-1861">1.2</span><span class="sxs-lookup"><span data-stu-id="c86d0-1861">1.2</span></span>|
|[<span data-ttu-id="c86d0-1862">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c86d0-1862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c86d0-1863">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c86d0-1863">ReadWriteItem</span></span>|
|[<span data-ttu-id="c86d0-1864">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c86d0-1864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c86d0-1865">撰写</span><span class="sxs-lookup"><span data-stu-id="c86d0-1865">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c86d0-1866">示例</span><span class="sxs-lookup"><span data-stu-id="c86d0-1866">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
