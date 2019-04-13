---
title: "\"context.subname\": \"邮箱. userProfile-预览要求集\""
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 29111314f16bb9c6518b350254a3036ffa125796
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838436"
---
# <a name="userprofile"></a><span data-ttu-id="e03ee-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e03ee-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e03ee-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e03ee-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e03ee-104">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-104">Requirements</span></span>

|<span data-ttu-id="e03ee-105">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-105">Requirement</span></span>| <span data-ttu-id="e03ee-106">值</span><span class="sxs-lookup"><span data-stu-id="e03ee-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e03ee-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e03ee-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e03ee-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e03ee-108">1.0</span></span>|
|[<span data-ttu-id="e03ee-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e03ee-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e03ee-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e03ee-110">ReadItem</span></span>|
|[<span data-ttu-id="e03ee-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e03ee-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e03ee-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e03ee-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e03ee-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="e03ee-113">Members and methods</span></span>

| <span data-ttu-id="e03ee-114">成员</span><span class="sxs-lookup"><span data-stu-id="e03ee-114">Member</span></span> | <span data-ttu-id="e03ee-115">类型</span><span class="sxs-lookup"><span data-stu-id="e03ee-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e03ee-116">accountType</span><span class="sxs-lookup"><span data-stu-id="e03ee-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="e03ee-117">Member</span><span class="sxs-lookup"><span data-stu-id="e03ee-117">Member</span></span> |
| [<span data-ttu-id="e03ee-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e03ee-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e03ee-119">Member</span><span class="sxs-lookup"><span data-stu-id="e03ee-119">Member</span></span> |
| [<span data-ttu-id="e03ee-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e03ee-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e03ee-121">Member</span><span class="sxs-lookup"><span data-stu-id="e03ee-121">Member</span></span> |
| [<span data-ttu-id="e03ee-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e03ee-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e03ee-123">Member</span><span class="sxs-lookup"><span data-stu-id="e03ee-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e03ee-124">Members</span><span class="sxs-lookup"><span data-stu-id="e03ee-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="e03ee-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="e03ee-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="e03ee-126">此成员目前仅在 Outlook 2016 或更高版本 for Mac (内部版本16.9.1212 或更高版本) 中受支持。</span><span class="sxs-lookup"><span data-stu-id="e03ee-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="e03ee-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="e03ee-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e03ee-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="e03ee-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="e03ee-129">值</span><span class="sxs-lookup"><span data-stu-id="e03ee-129">Value</span></span> | <span data-ttu-id="e03ee-130">说明</span><span class="sxs-lookup"><span data-stu-id="e03ee-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e03ee-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="e03ee-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e03ee-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="e03ee-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e03ee-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="e03ee-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e03ee-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="e03ee-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e03ee-135">类型</span><span class="sxs-lookup"><span data-stu-id="e03ee-135">Type</span></span>

*   <span data-ttu-id="e03ee-136">String</span><span class="sxs-lookup"><span data-stu-id="e03ee-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e03ee-137">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-137">Requirements</span></span>

|<span data-ttu-id="e03ee-138">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-138">Requirement</span></span>| <span data-ttu-id="e03ee-139">值</span><span class="sxs-lookup"><span data-stu-id="e03ee-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e03ee-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e03ee-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e03ee-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e03ee-141">1.6</span></span> |
|[<span data-ttu-id="e03ee-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e03ee-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e03ee-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e03ee-143">ReadItem</span></span>|
|[<span data-ttu-id="e03ee-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e03ee-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e03ee-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e03ee-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e03ee-146">示例</span><span class="sxs-lookup"><span data-stu-id="e03ee-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="e03ee-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e03ee-147">displayName :String</span></span>

<span data-ttu-id="e03ee-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="e03ee-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e03ee-149">类型</span><span class="sxs-lookup"><span data-stu-id="e03ee-149">Type</span></span>

*   <span data-ttu-id="e03ee-150">String</span><span class="sxs-lookup"><span data-stu-id="e03ee-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e03ee-151">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-151">Requirements</span></span>

|<span data-ttu-id="e03ee-152">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-152">Requirement</span></span>| <span data-ttu-id="e03ee-153">值</span><span class="sxs-lookup"><span data-stu-id="e03ee-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e03ee-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e03ee-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e03ee-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e03ee-155">1.0</span></span>|
|[<span data-ttu-id="e03ee-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e03ee-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e03ee-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e03ee-157">ReadItem</span></span>|
|[<span data-ttu-id="e03ee-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e03ee-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e03ee-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e03ee-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e03ee-160">示例</span><span class="sxs-lookup"><span data-stu-id="e03ee-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="e03ee-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e03ee-161">emailAddress :String</span></span>

<span data-ttu-id="e03ee-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="e03ee-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e03ee-163">类型</span><span class="sxs-lookup"><span data-stu-id="e03ee-163">Type</span></span>

*   <span data-ttu-id="e03ee-164">String</span><span class="sxs-lookup"><span data-stu-id="e03ee-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e03ee-165">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-165">Requirements</span></span>

|<span data-ttu-id="e03ee-166">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-166">Requirement</span></span>| <span data-ttu-id="e03ee-167">值</span><span class="sxs-lookup"><span data-stu-id="e03ee-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e03ee-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e03ee-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e03ee-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e03ee-169">1.0</span></span>|
|[<span data-ttu-id="e03ee-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e03ee-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e03ee-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e03ee-171">ReadItem</span></span>|
|[<span data-ttu-id="e03ee-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e03ee-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e03ee-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e03ee-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e03ee-174">示例</span><span class="sxs-lookup"><span data-stu-id="e03ee-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="e03ee-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e03ee-175">timeZone :String</span></span>

<span data-ttu-id="e03ee-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="e03ee-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e03ee-177">类型</span><span class="sxs-lookup"><span data-stu-id="e03ee-177">Type</span></span>

*   <span data-ttu-id="e03ee-178">String</span><span class="sxs-lookup"><span data-stu-id="e03ee-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e03ee-179">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-179">Requirements</span></span>

|<span data-ttu-id="e03ee-180">要求</span><span class="sxs-lookup"><span data-stu-id="e03ee-180">Requirement</span></span>| <span data-ttu-id="e03ee-181">值</span><span class="sxs-lookup"><span data-stu-id="e03ee-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e03ee-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e03ee-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e03ee-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e03ee-183">1.0</span></span>|
|[<span data-ttu-id="e03ee-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e03ee-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e03ee-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e03ee-185">ReadItem</span></span>|
|[<span data-ttu-id="e03ee-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e03ee-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e03ee-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e03ee-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e03ee-188">示例</span><span class="sxs-lookup"><span data-stu-id="e03ee-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
