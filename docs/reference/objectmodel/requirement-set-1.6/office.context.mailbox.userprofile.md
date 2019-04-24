---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.6\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9bb4335690236bdbbf2004f04f9af924747366d4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450217"
---
# <a name="userprofile"></a><span data-ttu-id="79014-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="79014-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="79014-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="79014-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="79014-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="79014-104">Requirements</span></span>

|<span data-ttu-id="79014-105">要求</span><span class="sxs-lookup"><span data-stu-id="79014-105">Requirement</span></span>| <span data-ttu-id="79014-106">值</span><span class="sxs-lookup"><span data-stu-id="79014-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="79014-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="79014-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79014-108">1.0</span><span class="sxs-lookup"><span data-stu-id="79014-108">1.0</span></span>|
|[<span data-ttu-id="79014-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="79014-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79014-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79014-110">ReadItem</span></span>|
|[<span data-ttu-id="79014-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="79014-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="79014-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="79014-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="79014-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="79014-113">Members and methods</span></span>

| <span data-ttu-id="79014-114">成员</span><span class="sxs-lookup"><span data-stu-id="79014-114">Member</span></span> | <span data-ttu-id="79014-115">类型</span><span class="sxs-lookup"><span data-stu-id="79014-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="79014-116">accountType</span><span class="sxs-lookup"><span data-stu-id="79014-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="79014-117">Member</span><span class="sxs-lookup"><span data-stu-id="79014-117">Member</span></span> |
| [<span data-ttu-id="79014-118">displayName</span><span class="sxs-lookup"><span data-stu-id="79014-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="79014-119">Member</span><span class="sxs-lookup"><span data-stu-id="79014-119">Member</span></span> |
| [<span data-ttu-id="79014-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="79014-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="79014-121">Member</span><span class="sxs-lookup"><span data-stu-id="79014-121">Member</span></span> |
| [<span data-ttu-id="79014-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="79014-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="79014-123">Member</span><span class="sxs-lookup"><span data-stu-id="79014-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="79014-124">Members</span><span class="sxs-lookup"><span data-stu-id="79014-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="79014-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="79014-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="79014-126">此成员目前仅在 Outlook 2016 或更高版本 for Mac (内部版本16.9.1212 或更高版本) 中受支持。</span><span class="sxs-lookup"><span data-stu-id="79014-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="79014-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="79014-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="79014-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="79014-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="79014-129">值</span><span class="sxs-lookup"><span data-stu-id="79014-129">Value</span></span> | <span data-ttu-id="79014-130">说明</span><span class="sxs-lookup"><span data-stu-id="79014-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="79014-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="79014-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="79014-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="79014-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="79014-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="79014-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="79014-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="79014-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="79014-135">类型</span><span class="sxs-lookup"><span data-stu-id="79014-135">Type</span></span>

*   <span data-ttu-id="79014-136">String</span><span class="sxs-lookup"><span data-stu-id="79014-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79014-137">要求</span><span class="sxs-lookup"><span data-stu-id="79014-137">Requirements</span></span>

|<span data-ttu-id="79014-138">要求</span><span class="sxs-lookup"><span data-stu-id="79014-138">Requirement</span></span>| <span data-ttu-id="79014-139">值</span><span class="sxs-lookup"><span data-stu-id="79014-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="79014-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="79014-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79014-141">1.6</span><span class="sxs-lookup"><span data-stu-id="79014-141">1.6</span></span> |
|[<span data-ttu-id="79014-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="79014-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79014-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79014-143">ReadItem</span></span>|
|[<span data-ttu-id="79014-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="79014-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="79014-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="79014-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79014-146">示例</span><span class="sxs-lookup"><span data-stu-id="79014-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="79014-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="79014-147">displayName :String</span></span>

<span data-ttu-id="79014-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="79014-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="79014-149">类型</span><span class="sxs-lookup"><span data-stu-id="79014-149">Type</span></span>

*   <span data-ttu-id="79014-150">String</span><span class="sxs-lookup"><span data-stu-id="79014-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79014-151">要求</span><span class="sxs-lookup"><span data-stu-id="79014-151">Requirements</span></span>

|<span data-ttu-id="79014-152">要求</span><span class="sxs-lookup"><span data-stu-id="79014-152">Requirement</span></span>| <span data-ttu-id="79014-153">值</span><span class="sxs-lookup"><span data-stu-id="79014-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="79014-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="79014-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79014-155">1.0</span><span class="sxs-lookup"><span data-stu-id="79014-155">1.0</span></span>|
|[<span data-ttu-id="79014-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="79014-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79014-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79014-157">ReadItem</span></span>|
|[<span data-ttu-id="79014-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="79014-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="79014-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="79014-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79014-160">示例</span><span class="sxs-lookup"><span data-stu-id="79014-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="79014-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="79014-161">emailAddress :String</span></span>

<span data-ttu-id="79014-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="79014-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="79014-163">类型</span><span class="sxs-lookup"><span data-stu-id="79014-163">Type</span></span>

*   <span data-ttu-id="79014-164">String</span><span class="sxs-lookup"><span data-stu-id="79014-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79014-165">要求</span><span class="sxs-lookup"><span data-stu-id="79014-165">Requirements</span></span>

|<span data-ttu-id="79014-166">要求</span><span class="sxs-lookup"><span data-stu-id="79014-166">Requirement</span></span>| <span data-ttu-id="79014-167">值</span><span class="sxs-lookup"><span data-stu-id="79014-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="79014-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="79014-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79014-169">1.0</span><span class="sxs-lookup"><span data-stu-id="79014-169">1.0</span></span>|
|[<span data-ttu-id="79014-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="79014-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79014-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79014-171">ReadItem</span></span>|
|[<span data-ttu-id="79014-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="79014-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="79014-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="79014-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79014-174">示例</span><span class="sxs-lookup"><span data-stu-id="79014-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="79014-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="79014-175">timeZone :String</span></span>

<span data-ttu-id="79014-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="79014-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="79014-177">类型</span><span class="sxs-lookup"><span data-stu-id="79014-177">Type</span></span>

*   <span data-ttu-id="79014-178">String</span><span class="sxs-lookup"><span data-stu-id="79014-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79014-179">要求</span><span class="sxs-lookup"><span data-stu-id="79014-179">Requirements</span></span>

|<span data-ttu-id="79014-180">要求</span><span class="sxs-lookup"><span data-stu-id="79014-180">Requirement</span></span>| <span data-ttu-id="79014-181">值</span><span class="sxs-lookup"><span data-stu-id="79014-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="79014-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="79014-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79014-183">1.0</span><span class="sxs-lookup"><span data-stu-id="79014-183">1.0</span></span>|
|[<span data-ttu-id="79014-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="79014-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79014-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79014-185">ReadItem</span></span>|
|[<span data-ttu-id="79014-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="79014-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="79014-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="79014-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="79014-188">示例</span><span class="sxs-lookup"><span data-stu-id="79014-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
