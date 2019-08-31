---
title: "\"Context.subname\": \"邮箱. userProfile-预览要求集\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5941c4e1276535091a3ffcf5b2fb6aa972ed8c4d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696468"
---
# <a name="userprofile"></a><span data-ttu-id="3efa4-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="3efa4-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="3efa4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="3efa4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3efa4-104">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-104">Requirements</span></span>

|<span data-ttu-id="3efa4-105">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-105">Requirement</span></span>| <span data-ttu-id="3efa4-106">值</span><span class="sxs-lookup"><span data-stu-id="3efa4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3efa4-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3efa4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3efa4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3efa4-108">1.0</span></span>|
|[<span data-ttu-id="3efa4-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3efa4-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3efa4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3efa4-110">ReadItem</span></span>|
|[<span data-ttu-id="3efa4-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3efa4-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3efa4-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3efa4-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3efa4-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="3efa4-113">Members and methods</span></span>

| <span data-ttu-id="3efa4-114">成员</span><span class="sxs-lookup"><span data-stu-id="3efa4-114">Member</span></span> | <span data-ttu-id="3efa4-115">类型</span><span class="sxs-lookup"><span data-stu-id="3efa4-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3efa4-116">accountType</span><span class="sxs-lookup"><span data-stu-id="3efa4-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="3efa4-117">Member</span><span class="sxs-lookup"><span data-stu-id="3efa4-117">Member</span></span> |
| [<span data-ttu-id="3efa4-118">displayName</span><span class="sxs-lookup"><span data-stu-id="3efa4-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="3efa4-119">Member</span><span class="sxs-lookup"><span data-stu-id="3efa4-119">Member</span></span> |
| [<span data-ttu-id="3efa4-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="3efa4-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="3efa4-121">Member</span><span class="sxs-lookup"><span data-stu-id="3efa4-121">Member</span></span> |
| [<span data-ttu-id="3efa4-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="3efa4-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="3efa4-123">Member</span><span class="sxs-lookup"><span data-stu-id="3efa4-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="3efa4-124">Members</span><span class="sxs-lookup"><span data-stu-id="3efa4-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="3efa4-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="3efa4-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="3efa4-126">此成员目前仅在 Outlook 2016 或更高版本 (内部版本16.9.1212 或更高版本) 中受支持。</span><span class="sxs-lookup"><span data-stu-id="3efa4-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="3efa4-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="3efa4-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="3efa4-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="3efa4-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="3efa4-129">值</span><span class="sxs-lookup"><span data-stu-id="3efa4-129">Value</span></span> | <span data-ttu-id="3efa4-130">说明</span><span class="sxs-lookup"><span data-stu-id="3efa4-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="3efa4-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="3efa4-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="3efa4-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="3efa4-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="3efa4-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="3efa4-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="3efa4-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="3efa4-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="3efa4-135">类型</span><span class="sxs-lookup"><span data-stu-id="3efa4-135">Type</span></span>

*   <span data-ttu-id="3efa4-136">String</span><span class="sxs-lookup"><span data-stu-id="3efa4-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3efa4-137">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-137">Requirements</span></span>

|<span data-ttu-id="3efa4-138">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-138">Requirement</span></span>| <span data-ttu-id="3efa4-139">值</span><span class="sxs-lookup"><span data-stu-id="3efa4-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="3efa4-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3efa4-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3efa4-141">1.6</span><span class="sxs-lookup"><span data-stu-id="3efa4-141">1.6</span></span> |
|[<span data-ttu-id="3efa4-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3efa4-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3efa4-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3efa4-143">ReadItem</span></span>|
|[<span data-ttu-id="3efa4-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3efa4-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3efa4-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3efa4-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3efa4-146">示例</span><span class="sxs-lookup"><span data-stu-id="3efa4-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="3efa4-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="3efa4-147">displayName: String</span></span>

<span data-ttu-id="3efa4-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="3efa4-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3efa4-149">类型</span><span class="sxs-lookup"><span data-stu-id="3efa4-149">Type</span></span>

*   <span data-ttu-id="3efa4-150">String</span><span class="sxs-lookup"><span data-stu-id="3efa4-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3efa4-151">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-151">Requirements</span></span>

|<span data-ttu-id="3efa4-152">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-152">Requirement</span></span>| <span data-ttu-id="3efa4-153">值</span><span class="sxs-lookup"><span data-stu-id="3efa4-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="3efa4-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3efa4-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3efa4-155">1.0</span><span class="sxs-lookup"><span data-stu-id="3efa4-155">1.0</span></span>|
|[<span data-ttu-id="3efa4-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3efa4-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3efa4-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3efa4-157">ReadItem</span></span>|
|[<span data-ttu-id="3efa4-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3efa4-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3efa4-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3efa4-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3efa4-160">示例</span><span class="sxs-lookup"><span data-stu-id="3efa4-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="3efa4-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="3efa4-161">emailAddress: String</span></span>

<span data-ttu-id="3efa4-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="3efa4-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3efa4-163">类型</span><span class="sxs-lookup"><span data-stu-id="3efa4-163">Type</span></span>

*   <span data-ttu-id="3efa4-164">String</span><span class="sxs-lookup"><span data-stu-id="3efa4-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3efa4-165">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-165">Requirements</span></span>

|<span data-ttu-id="3efa4-166">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-166">Requirement</span></span>| <span data-ttu-id="3efa4-167">值</span><span class="sxs-lookup"><span data-stu-id="3efa4-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3efa4-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3efa4-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3efa4-169">1.0</span><span class="sxs-lookup"><span data-stu-id="3efa4-169">1.0</span></span>|
|[<span data-ttu-id="3efa4-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3efa4-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3efa4-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3efa4-171">ReadItem</span></span>|
|[<span data-ttu-id="3efa4-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3efa4-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3efa4-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3efa4-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3efa4-174">示例</span><span class="sxs-lookup"><span data-stu-id="3efa4-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="3efa4-175">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="3efa4-175">timeZone: String</span></span>

<span data-ttu-id="3efa4-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="3efa4-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3efa4-177">类型</span><span class="sxs-lookup"><span data-stu-id="3efa4-177">Type</span></span>

*   <span data-ttu-id="3efa4-178">String</span><span class="sxs-lookup"><span data-stu-id="3efa4-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3efa4-179">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-179">Requirements</span></span>

|<span data-ttu-id="3efa4-180">要求</span><span class="sxs-lookup"><span data-stu-id="3efa4-180">Requirement</span></span>| <span data-ttu-id="3efa4-181">值</span><span class="sxs-lookup"><span data-stu-id="3efa4-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="3efa4-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3efa4-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3efa4-183">1.0</span><span class="sxs-lookup"><span data-stu-id="3efa4-183">1.0</span></span>|
|[<span data-ttu-id="3efa4-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3efa4-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3efa4-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3efa4-185">ReadItem</span></span>|
|[<span data-ttu-id="3efa4-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3efa4-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3efa4-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3efa4-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3efa4-188">示例</span><span class="sxs-lookup"><span data-stu-id="3efa4-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
