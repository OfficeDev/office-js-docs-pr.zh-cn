---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.7\""
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 036f18e4cb98cfe510a19d85a5a79f393ca8bd17
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353291"
---
# <a name="userprofile"></a><span data-ttu-id="b8ffd-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="b8ffd-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="b8ffd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="b8ffd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8ffd-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="b8ffd-104">Requirements</span></span>

|<span data-ttu-id="b8ffd-105">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-105">Requirement</span></span>| <span data-ttu-id="b8ffd-106">值</span><span class="sxs-lookup"><span data-stu-id="b8ffd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8ffd-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8ffd-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8ffd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b8ffd-108">1.0</span></span>|
|[<span data-ttu-id="b8ffd-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b8ffd-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8ffd-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8ffd-110">ReadItem</span></span>|
|[<span data-ttu-id="b8ffd-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8ffd-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8ffd-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8ffd-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b8ffd-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="b8ffd-113">Members and methods</span></span>

| <span data-ttu-id="b8ffd-114">成员</span><span class="sxs-lookup"><span data-stu-id="b8ffd-114">Member</span></span> | <span data-ttu-id="b8ffd-115">类型</span><span class="sxs-lookup"><span data-stu-id="b8ffd-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b8ffd-116">accountType</span><span class="sxs-lookup"><span data-stu-id="b8ffd-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="b8ffd-117">Member</span><span class="sxs-lookup"><span data-stu-id="b8ffd-117">Member</span></span> |
| [<span data-ttu-id="b8ffd-118">displayName</span><span class="sxs-lookup"><span data-stu-id="b8ffd-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="b8ffd-119">Member</span><span class="sxs-lookup"><span data-stu-id="b8ffd-119">Member</span></span> |
| [<span data-ttu-id="b8ffd-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="b8ffd-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="b8ffd-121">Member</span><span class="sxs-lookup"><span data-stu-id="b8ffd-121">Member</span></span> |
| [<span data-ttu-id="b8ffd-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="b8ffd-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="b8ffd-123">Member</span><span class="sxs-lookup"><span data-stu-id="b8ffd-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="b8ffd-124">Members</span><span class="sxs-lookup"><span data-stu-id="b8ffd-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="b8ffd-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="b8ffd-126">此成员目前仅支持适用于 Mac 的 Outlook 2016 (内部版本16.9.1212 或更高版本)。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="b8ffd-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="b8ffd-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="b8ffd-129">值</span><span class="sxs-lookup"><span data-stu-id="b8ffd-129">Value</span></span> | <span data-ttu-id="b8ffd-130">说明</span><span class="sxs-lookup"><span data-stu-id="b8ffd-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="b8ffd-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="b8ffd-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="b8ffd-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="b8ffd-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="b8ffd-135">类型</span><span class="sxs-lookup"><span data-stu-id="b8ffd-135">Type</span></span>

*   <span data-ttu-id="b8ffd-136">String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8ffd-137">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-137">Requirements</span></span>

|<span data-ttu-id="b8ffd-138">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-138">Requirement</span></span>| <span data-ttu-id="b8ffd-139">值</span><span class="sxs-lookup"><span data-stu-id="b8ffd-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8ffd-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8ffd-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8ffd-141">1.6</span><span class="sxs-lookup"><span data-stu-id="b8ffd-141">1.6</span></span> |
|[<span data-ttu-id="b8ffd-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b8ffd-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8ffd-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8ffd-143">ReadItem</span></span>|
|[<span data-ttu-id="b8ffd-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8ffd-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8ffd-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8ffd-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b8ffd-146">示例</span><span class="sxs-lookup"><span data-stu-id="b8ffd-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="b8ffd-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-147">displayName: String</span></span>

<span data-ttu-id="b8ffd-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b8ffd-149">类型</span><span class="sxs-lookup"><span data-stu-id="b8ffd-149">Type</span></span>

*   <span data-ttu-id="b8ffd-150">String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8ffd-151">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-151">Requirements</span></span>

|<span data-ttu-id="b8ffd-152">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-152">Requirement</span></span>| <span data-ttu-id="b8ffd-153">值</span><span class="sxs-lookup"><span data-stu-id="b8ffd-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8ffd-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8ffd-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8ffd-155">1.0</span><span class="sxs-lookup"><span data-stu-id="b8ffd-155">1.0</span></span>|
|[<span data-ttu-id="b8ffd-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b8ffd-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8ffd-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8ffd-157">ReadItem</span></span>|
|[<span data-ttu-id="b8ffd-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8ffd-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8ffd-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8ffd-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b8ffd-160">示例</span><span class="sxs-lookup"><span data-stu-id="b8ffd-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="b8ffd-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-161">emailAddress: String</span></span>

<span data-ttu-id="b8ffd-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b8ffd-163">类型</span><span class="sxs-lookup"><span data-stu-id="b8ffd-163">Type</span></span>

*   <span data-ttu-id="b8ffd-164">String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8ffd-165">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-165">Requirements</span></span>

|<span data-ttu-id="b8ffd-166">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-166">Requirement</span></span>| <span data-ttu-id="b8ffd-167">值</span><span class="sxs-lookup"><span data-stu-id="b8ffd-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8ffd-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8ffd-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8ffd-169">1.0</span><span class="sxs-lookup"><span data-stu-id="b8ffd-169">1.0</span></span>|
|[<span data-ttu-id="b8ffd-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b8ffd-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8ffd-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8ffd-171">ReadItem</span></span>|
|[<span data-ttu-id="b8ffd-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8ffd-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8ffd-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8ffd-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b8ffd-174">示例</span><span class="sxs-lookup"><span data-stu-id="b8ffd-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="b8ffd-175">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="b8ffd-175">timeZone: String</span></span>

<span data-ttu-id="b8ffd-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="b8ffd-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b8ffd-177">类型</span><span class="sxs-lookup"><span data-stu-id="b8ffd-177">Type</span></span>

*   <span data-ttu-id="b8ffd-178">String</span><span class="sxs-lookup"><span data-stu-id="b8ffd-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8ffd-179">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-179">Requirements</span></span>

|<span data-ttu-id="b8ffd-180">要求</span><span class="sxs-lookup"><span data-stu-id="b8ffd-180">Requirement</span></span>| <span data-ttu-id="b8ffd-181">值</span><span class="sxs-lookup"><span data-stu-id="b8ffd-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8ffd-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8ffd-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8ffd-183">1.0</span><span class="sxs-lookup"><span data-stu-id="b8ffd-183">1.0</span></span>|
|[<span data-ttu-id="b8ffd-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b8ffd-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8ffd-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8ffd-185">ReadItem</span></span>|
|[<span data-ttu-id="b8ffd-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8ffd-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8ffd-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8ffd-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b8ffd-188">示例</span><span class="sxs-lookup"><span data-stu-id="b8ffd-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
