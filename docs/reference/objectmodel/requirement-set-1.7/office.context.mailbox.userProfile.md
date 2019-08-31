---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.7\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 8e33d40bec9b561c642ad6e0da73ae13a18378b6
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695901"
---
# <a name="userprofile"></a><span data-ttu-id="f4e6e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="f4e6e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="f4e6e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="f4e6e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4e6e-104">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-104">Requirements</span></span>

|<span data-ttu-id="f4e6e-105">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-105">Requirement</span></span>| <span data-ttu-id="f4e6e-106">值</span><span class="sxs-lookup"><span data-stu-id="f4e6e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4e6e-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f4e6e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4e6e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f4e6e-108">1.0</span></span>|
|[<span data-ttu-id="f4e6e-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f4e6e-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4e6e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4e6e-110">ReadItem</span></span>|
|[<span data-ttu-id="f4e6e-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f4e6e-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4e6e-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f4e6e-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f4e6e-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="f4e6e-113">Members and methods</span></span>

| <span data-ttu-id="f4e6e-114">成员</span><span class="sxs-lookup"><span data-stu-id="f4e6e-114">Member</span></span> | <span data-ttu-id="f4e6e-115">类型</span><span class="sxs-lookup"><span data-stu-id="f4e6e-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f4e6e-116">accountType</span><span class="sxs-lookup"><span data-stu-id="f4e6e-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="f4e6e-117">Member</span><span class="sxs-lookup"><span data-stu-id="f4e6e-117">Member</span></span> |
| [<span data-ttu-id="f4e6e-118">displayName</span><span class="sxs-lookup"><span data-stu-id="f4e6e-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="f4e6e-119">Member</span><span class="sxs-lookup"><span data-stu-id="f4e6e-119">Member</span></span> |
| [<span data-ttu-id="f4e6e-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f4e6e-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="f4e6e-121">Member</span><span class="sxs-lookup"><span data-stu-id="f4e6e-121">Member</span></span> |
| [<span data-ttu-id="f4e6e-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="f4e6e-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="f4e6e-123">Member</span><span class="sxs-lookup"><span data-stu-id="f4e6e-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f4e6e-124">Members</span><span class="sxs-lookup"><span data-stu-id="f4e6e-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="f4e6e-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="f4e6e-126">此成员目前仅受 Outlook 2016 或更高版本的 Mac (内部版本16.9.1212 或更高版本) 支持。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-126">This member is currently only supported by Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="f4e6e-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="f4e6e-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="f4e6e-129">值</span><span class="sxs-lookup"><span data-stu-id="f4e6e-129">Value</span></span> | <span data-ttu-id="f4e6e-130">说明</span><span class="sxs-lookup"><span data-stu-id="f4e6e-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="f4e6e-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="f4e6e-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="f4e6e-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="f4e6e-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="f4e6e-135">类型</span><span class="sxs-lookup"><span data-stu-id="f4e6e-135">Type</span></span>

*   <span data-ttu-id="f4e6e-136">String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4e6e-137">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-137">Requirements</span></span>

|<span data-ttu-id="f4e6e-138">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-138">Requirement</span></span>| <span data-ttu-id="f4e6e-139">值</span><span class="sxs-lookup"><span data-stu-id="f4e6e-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4e6e-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f4e6e-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4e6e-141">1.6</span><span class="sxs-lookup"><span data-stu-id="f4e6e-141">1.6</span></span> |
|[<span data-ttu-id="f4e6e-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f4e6e-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4e6e-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4e6e-143">ReadItem</span></span>|
|[<span data-ttu-id="f4e6e-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f4e6e-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4e6e-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f4e6e-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4e6e-146">示例</span><span class="sxs-lookup"><span data-stu-id="f4e6e-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="f4e6e-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-147">displayName: String</span></span>

<span data-ttu-id="f4e6e-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="f4e6e-149">类型</span><span class="sxs-lookup"><span data-stu-id="f4e6e-149">Type</span></span>

*   <span data-ttu-id="f4e6e-150">String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4e6e-151">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-151">Requirements</span></span>

|<span data-ttu-id="f4e6e-152">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-152">Requirement</span></span>| <span data-ttu-id="f4e6e-153">值</span><span class="sxs-lookup"><span data-stu-id="f4e6e-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4e6e-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f4e6e-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4e6e-155">1.0</span><span class="sxs-lookup"><span data-stu-id="f4e6e-155">1.0</span></span>|
|[<span data-ttu-id="f4e6e-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f4e6e-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4e6e-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4e6e-157">ReadItem</span></span>|
|[<span data-ttu-id="f4e6e-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f4e6e-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4e6e-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f4e6e-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4e6e-160">示例</span><span class="sxs-lookup"><span data-stu-id="f4e6e-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="f4e6e-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-161">emailAddress: String</span></span>

<span data-ttu-id="f4e6e-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="f4e6e-163">类型</span><span class="sxs-lookup"><span data-stu-id="f4e6e-163">Type</span></span>

*   <span data-ttu-id="f4e6e-164">String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4e6e-165">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-165">Requirements</span></span>

|<span data-ttu-id="f4e6e-166">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-166">Requirement</span></span>| <span data-ttu-id="f4e6e-167">值</span><span class="sxs-lookup"><span data-stu-id="f4e6e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4e6e-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f4e6e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4e6e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="f4e6e-169">1.0</span></span>|
|[<span data-ttu-id="f4e6e-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f4e6e-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4e6e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4e6e-171">ReadItem</span></span>|
|[<span data-ttu-id="f4e6e-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f4e6e-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4e6e-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f4e6e-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4e6e-174">示例</span><span class="sxs-lookup"><span data-stu-id="f4e6e-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="f4e6e-175">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="f4e6e-175">timeZone: String</span></span>

<span data-ttu-id="f4e6e-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="f4e6e-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="f4e6e-177">类型</span><span class="sxs-lookup"><span data-stu-id="f4e6e-177">Type</span></span>

*   <span data-ttu-id="f4e6e-178">String</span><span class="sxs-lookup"><span data-stu-id="f4e6e-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4e6e-179">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-179">Requirements</span></span>

|<span data-ttu-id="f4e6e-180">要求</span><span class="sxs-lookup"><span data-stu-id="f4e6e-180">Requirement</span></span>| <span data-ttu-id="f4e6e-181">值</span><span class="sxs-lookup"><span data-stu-id="f4e6e-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4e6e-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f4e6e-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4e6e-183">1.0</span><span class="sxs-lookup"><span data-stu-id="f4e6e-183">1.0</span></span>|
|[<span data-ttu-id="f4e6e-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f4e6e-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4e6e-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4e6e-185">ReadItem</span></span>|
|[<span data-ttu-id="f4e6e-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f4e6e-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4e6e-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f4e6e-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4e6e-188">示例</span><span class="sxs-lookup"><span data-stu-id="f4e6e-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
