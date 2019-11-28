---
title: "\"Context.subname\"： \"邮箱. userProfile-预览要求集\""
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 4afc64f247155576ab3f0024d1929a29a0f7dc0c
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629256"
---
# <a name="userprofile"></a><span data-ttu-id="a4fa0-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a4fa0-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a4fa0-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a4fa0-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4fa0-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="a4fa0-104">Requirements</span></span>

|<span data-ttu-id="a4fa0-105">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-105">Requirement</span></span>| <span data-ttu-id="a4fa0-106">值</span><span class="sxs-lookup"><span data-stu-id="a4fa0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4fa0-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a4fa0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4fa0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-108">1.0</span></span>|
|[<span data-ttu-id="a4fa0-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a4fa0-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4fa0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-110">ReadItem</span></span>|
|[<span data-ttu-id="a4fa0-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a4fa0-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4fa0-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a4fa0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a4fa0-113">属性</span><span class="sxs-lookup"><span data-stu-id="a4fa0-113">Properties</span></span>

| <span data-ttu-id="a4fa0-114">属性</span><span class="sxs-lookup"><span data-stu-id="a4fa0-114">Property</span></span> | <span data-ttu-id="a4fa0-115">最低</span><span class="sxs-lookup"><span data-stu-id="a4fa0-115">Minimum</span></span><br><span data-ttu-id="a4fa0-116">权限级别</span><span class="sxs-lookup"><span data-stu-id="a4fa0-116">permission level</span></span> | <span data-ttu-id="a4fa0-117">型号</span><span class="sxs-lookup"><span data-stu-id="a4fa0-117">Modes</span></span> | <span data-ttu-id="a4fa0-118">返回类型</span><span class="sxs-lookup"><span data-stu-id="a4fa0-118">Return type</span></span> | <span data-ttu-id="a4fa0-119">最低</span><span class="sxs-lookup"><span data-stu-id="a4fa0-119">Minimum</span></span><br><span data-ttu-id="a4fa0-120">要求集</span><span class="sxs-lookup"><span data-stu-id="a4fa0-120">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="a4fa0-121">accountType</span><span class="sxs-lookup"><span data-stu-id="a4fa0-121">accountType</span></span>](#accounttype-string) | <span data-ttu-id="a4fa0-122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-122">ReadItem</span></span> | <span data-ttu-id="a4fa0-123">撰写</span><span class="sxs-lookup"><span data-stu-id="a4fa0-123">Compose</span></span><br><span data-ttu-id="a4fa0-124">读取</span><span class="sxs-lookup"><span data-stu-id="a4fa0-124">Read</span></span> | <span data-ttu-id="a4fa0-125">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-125">String</span></span> | <span data-ttu-id="a4fa0-126">1.6</span><span class="sxs-lookup"><span data-stu-id="a4fa0-126">1.6</span></span> |
| [<span data-ttu-id="a4fa0-127">displayName</span><span class="sxs-lookup"><span data-stu-id="a4fa0-127">displayName</span></span>](#displayname-string) | <span data-ttu-id="a4fa0-128">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-128">ReadItem</span></span> | <span data-ttu-id="a4fa0-129">撰写</span><span class="sxs-lookup"><span data-stu-id="a4fa0-129">Compose</span></span><br><span data-ttu-id="a4fa0-130">读取</span><span class="sxs-lookup"><span data-stu-id="a4fa0-130">Read</span></span> | <span data-ttu-id="a4fa0-131">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-131">String</span></span> | <span data-ttu-id="a4fa0-132">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-132">1.0</span></span> |
| [<span data-ttu-id="a4fa0-133">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a4fa0-133">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a4fa0-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-134">ReadItem</span></span> | <span data-ttu-id="a4fa0-135">撰写</span><span class="sxs-lookup"><span data-stu-id="a4fa0-135">Compose</span></span><br><span data-ttu-id="a4fa0-136">读取</span><span class="sxs-lookup"><span data-stu-id="a4fa0-136">Read</span></span> | <span data-ttu-id="a4fa0-137">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-137">String</span></span> | <span data-ttu-id="a4fa0-138">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-138">1.0</span></span> |
| [<span data-ttu-id="a4fa0-139">timeZone</span><span class="sxs-lookup"><span data-stu-id="a4fa0-139">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a4fa0-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-140">ReadItem</span></span> | <span data-ttu-id="a4fa0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="a4fa0-141">Compose</span></span><br><span data-ttu-id="a4fa0-142">读取</span><span class="sxs-lookup"><span data-stu-id="a4fa0-142">Read</span></span> | <span data-ttu-id="a4fa0-143">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-143">String</span></span> | <span data-ttu-id="a4fa0-144">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-144">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="a4fa0-145">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="a4fa0-145">Property details</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="a4fa0-146">accountType： String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-146">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="a4fa0-147">此成员目前仅在 Outlook 2016 或更高版本（内部版本16.9.1212 或更高版本）中受支持。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-147">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="a4fa0-148">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-148">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="a4fa0-149">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-149">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="a4fa0-150">值</span><span class="sxs-lookup"><span data-stu-id="a4fa0-150">Value</span></span> | <span data-ttu-id="a4fa0-151">说明</span><span class="sxs-lookup"><span data-stu-id="a4fa0-151">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="a4fa0-152">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-152">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="a4fa0-153">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-153">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="a4fa0-154">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-154">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="a4fa0-155">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-155">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="a4fa0-156">类型</span><span class="sxs-lookup"><span data-stu-id="a4fa0-156">Type</span></span>

*   <span data-ttu-id="a4fa0-157">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4fa0-158">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-158">Requirements</span></span>

|<span data-ttu-id="a4fa0-159">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-159">Requirement</span></span>| <span data-ttu-id="a4fa0-160">值</span><span class="sxs-lookup"><span data-stu-id="a4fa0-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4fa0-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a4fa0-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4fa0-162">1.6</span><span class="sxs-lookup"><span data-stu-id="a4fa0-162">1.6</span></span> |
|[<span data-ttu-id="a4fa0-163">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a4fa0-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4fa0-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-164">ReadItem</span></span>|
|[<span data-ttu-id="a4fa0-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a4fa0-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4fa0-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a4fa0-166">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4fa0-167">示例</span><span class="sxs-lookup"><span data-stu-id="a4fa0-167">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="a4fa0-168">displayName： String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-168">displayName: String</span></span>

<span data-ttu-id="a4fa0-169">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-169">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a4fa0-170">类型</span><span class="sxs-lookup"><span data-stu-id="a4fa0-170">Type</span></span>

*   <span data-ttu-id="a4fa0-171">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-171">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4fa0-172">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-172">Requirements</span></span>

|<span data-ttu-id="a4fa0-173">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-173">Requirement</span></span>| <span data-ttu-id="a4fa0-174">值</span><span class="sxs-lookup"><span data-stu-id="a4fa0-174">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4fa0-175">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a4fa0-175">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4fa0-176">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-176">1.0</span></span>|
|[<span data-ttu-id="a4fa0-177">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a4fa0-177">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4fa0-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-178">ReadItem</span></span>|
|[<span data-ttu-id="a4fa0-179">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a4fa0-179">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4fa0-180">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a4fa0-180">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4fa0-181">示例</span><span class="sxs-lookup"><span data-stu-id="a4fa0-181">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="a4fa0-182">emailAddress： String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-182">emailAddress: String</span></span>

<span data-ttu-id="a4fa0-183">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-183">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a4fa0-184">类型</span><span class="sxs-lookup"><span data-stu-id="a4fa0-184">Type</span></span>

*   <span data-ttu-id="a4fa0-185">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4fa0-186">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-186">Requirements</span></span>

|<span data-ttu-id="a4fa0-187">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-187">Requirement</span></span>| <span data-ttu-id="a4fa0-188">值</span><span class="sxs-lookup"><span data-stu-id="a4fa0-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4fa0-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a4fa0-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4fa0-190">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-190">1.0</span></span>|
|[<span data-ttu-id="a4fa0-191">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a4fa0-191">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4fa0-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-192">ReadItem</span></span>|
|[<span data-ttu-id="a4fa0-193">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a4fa0-193">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4fa0-194">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a4fa0-194">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4fa0-195">示例</span><span class="sxs-lookup"><span data-stu-id="a4fa0-195">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="a4fa0-196">时区：字符串</span><span class="sxs-lookup"><span data-stu-id="a4fa0-196">timeZone: String</span></span>

<span data-ttu-id="a4fa0-197">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="a4fa0-197">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a4fa0-198">类型</span><span class="sxs-lookup"><span data-stu-id="a4fa0-198">Type</span></span>

*   <span data-ttu-id="a4fa0-199">String</span><span class="sxs-lookup"><span data-stu-id="a4fa0-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4fa0-200">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-200">Requirements</span></span>

|<span data-ttu-id="a4fa0-201">要求</span><span class="sxs-lookup"><span data-stu-id="a4fa0-201">Requirement</span></span>| <span data-ttu-id="a4fa0-202">值</span><span class="sxs-lookup"><span data-stu-id="a4fa0-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4fa0-203">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a4fa0-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4fa0-204">1.0</span><span class="sxs-lookup"><span data-stu-id="a4fa0-204">1.0</span></span>|
|[<span data-ttu-id="a4fa0-205">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a4fa0-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4fa0-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4fa0-206">ReadItem</span></span>|
|[<span data-ttu-id="a4fa0-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a4fa0-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4fa0-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a4fa0-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4fa0-209">示例</span><span class="sxs-lookup"><span data-stu-id="a4fa0-209">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
