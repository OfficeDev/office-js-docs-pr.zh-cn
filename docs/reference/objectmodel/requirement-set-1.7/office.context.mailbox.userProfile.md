---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.7\""
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 8cfee874bbb5183d62cc3a9ce8b042a76617ec72
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451806"
---
# <a name="userprofile"></a><span data-ttu-id="746ad-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="746ad-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="746ad-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="746ad-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="746ad-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="746ad-104">Requirements</span></span>

|<span data-ttu-id="746ad-105">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-105">Requirement</span></span>| <span data-ttu-id="746ad-106">值</span><span class="sxs-lookup"><span data-stu-id="746ad-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="746ad-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="746ad-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="746ad-108">1.0</span><span class="sxs-lookup"><span data-stu-id="746ad-108">1.0</span></span>|
|[<span data-ttu-id="746ad-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="746ad-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="746ad-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="746ad-110">ReadItem</span></span>|
|[<span data-ttu-id="746ad-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="746ad-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="746ad-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="746ad-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="746ad-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="746ad-113">Members and methods</span></span>

| <span data-ttu-id="746ad-114">成员</span><span class="sxs-lookup"><span data-stu-id="746ad-114">Member</span></span> | <span data-ttu-id="746ad-115">类型</span><span class="sxs-lookup"><span data-stu-id="746ad-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="746ad-116">accountType</span><span class="sxs-lookup"><span data-stu-id="746ad-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="746ad-117">Member</span><span class="sxs-lookup"><span data-stu-id="746ad-117">Member</span></span> |
| [<span data-ttu-id="746ad-118">displayName</span><span class="sxs-lookup"><span data-stu-id="746ad-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="746ad-119">Member</span><span class="sxs-lookup"><span data-stu-id="746ad-119">Member</span></span> |
| [<span data-ttu-id="746ad-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="746ad-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="746ad-121">Member</span><span class="sxs-lookup"><span data-stu-id="746ad-121">Member</span></span> |
| [<span data-ttu-id="746ad-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="746ad-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="746ad-123">Member</span><span class="sxs-lookup"><span data-stu-id="746ad-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="746ad-124">Members</span><span class="sxs-lookup"><span data-stu-id="746ad-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="746ad-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="746ad-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="746ad-126">此成员目前仅支持适用于 Mac 的 Outlook 2016 (内部版本16.9.1212 或更高版本)。</span><span class="sxs-lookup"><span data-stu-id="746ad-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="746ad-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="746ad-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="746ad-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="746ad-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="746ad-129">值</span><span class="sxs-lookup"><span data-stu-id="746ad-129">Value</span></span> | <span data-ttu-id="746ad-130">说明</span><span class="sxs-lookup"><span data-stu-id="746ad-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="746ad-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="746ad-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="746ad-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="746ad-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="746ad-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="746ad-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="746ad-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="746ad-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="746ad-135">类型</span><span class="sxs-lookup"><span data-stu-id="746ad-135">Type</span></span>

*   <span data-ttu-id="746ad-136">String</span><span class="sxs-lookup"><span data-stu-id="746ad-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="746ad-137">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-137">Requirements</span></span>

|<span data-ttu-id="746ad-138">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-138">Requirement</span></span>| <span data-ttu-id="746ad-139">值</span><span class="sxs-lookup"><span data-stu-id="746ad-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="746ad-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="746ad-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="746ad-141">1.6</span><span class="sxs-lookup"><span data-stu-id="746ad-141">1.6</span></span> |
|[<span data-ttu-id="746ad-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="746ad-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="746ad-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="746ad-143">ReadItem</span></span>|
|[<span data-ttu-id="746ad-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="746ad-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="746ad-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="746ad-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="746ad-146">示例</span><span class="sxs-lookup"><span data-stu-id="746ad-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="746ad-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="746ad-147">displayName :String</span></span>

<span data-ttu-id="746ad-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="746ad-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="746ad-149">类型</span><span class="sxs-lookup"><span data-stu-id="746ad-149">Type</span></span>

*   <span data-ttu-id="746ad-150">String</span><span class="sxs-lookup"><span data-stu-id="746ad-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="746ad-151">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-151">Requirements</span></span>

|<span data-ttu-id="746ad-152">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-152">Requirement</span></span>| <span data-ttu-id="746ad-153">值</span><span class="sxs-lookup"><span data-stu-id="746ad-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="746ad-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="746ad-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="746ad-155">1.0</span><span class="sxs-lookup"><span data-stu-id="746ad-155">1.0</span></span>|
|[<span data-ttu-id="746ad-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="746ad-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="746ad-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="746ad-157">ReadItem</span></span>|
|[<span data-ttu-id="746ad-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="746ad-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="746ad-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="746ad-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="746ad-160">示例</span><span class="sxs-lookup"><span data-stu-id="746ad-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="746ad-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="746ad-161">emailAddress :String</span></span>

<span data-ttu-id="746ad-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="746ad-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="746ad-163">类型</span><span class="sxs-lookup"><span data-stu-id="746ad-163">Type</span></span>

*   <span data-ttu-id="746ad-164">String</span><span class="sxs-lookup"><span data-stu-id="746ad-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="746ad-165">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-165">Requirements</span></span>

|<span data-ttu-id="746ad-166">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-166">Requirement</span></span>| <span data-ttu-id="746ad-167">值</span><span class="sxs-lookup"><span data-stu-id="746ad-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="746ad-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="746ad-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="746ad-169">1.0</span><span class="sxs-lookup"><span data-stu-id="746ad-169">1.0</span></span>|
|[<span data-ttu-id="746ad-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="746ad-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="746ad-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="746ad-171">ReadItem</span></span>|
|[<span data-ttu-id="746ad-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="746ad-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="746ad-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="746ad-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="746ad-174">示例</span><span class="sxs-lookup"><span data-stu-id="746ad-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="746ad-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="746ad-175">timeZone :String</span></span>

<span data-ttu-id="746ad-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="746ad-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="746ad-177">类型</span><span class="sxs-lookup"><span data-stu-id="746ad-177">Type</span></span>

*   <span data-ttu-id="746ad-178">String</span><span class="sxs-lookup"><span data-stu-id="746ad-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="746ad-179">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-179">Requirements</span></span>

|<span data-ttu-id="746ad-180">要求</span><span class="sxs-lookup"><span data-stu-id="746ad-180">Requirement</span></span>| <span data-ttu-id="746ad-181">值</span><span class="sxs-lookup"><span data-stu-id="746ad-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="746ad-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="746ad-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="746ad-183">1.0</span><span class="sxs-lookup"><span data-stu-id="746ad-183">1.0</span></span>|
|[<span data-ttu-id="746ad-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="746ad-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="746ad-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="746ad-185">ReadItem</span></span>|
|[<span data-ttu-id="746ad-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="746ad-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="746ad-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="746ad-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="746ad-188">示例</span><span class="sxs-lookup"><span data-stu-id="746ad-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
