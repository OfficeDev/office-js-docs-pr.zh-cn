---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.6\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9bb4335690236bdbbf2004f04f9af924747366d4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871568"
---
# <a name="userprofile"></a><span data-ttu-id="60208-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="60208-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="60208-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="60208-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="60208-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="60208-104">Requirements</span></span>

|<span data-ttu-id="60208-105">要求</span><span class="sxs-lookup"><span data-stu-id="60208-105">Requirement</span></span>| <span data-ttu-id="60208-106">值</span><span class="sxs-lookup"><span data-stu-id="60208-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="60208-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60208-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60208-108">1.0</span><span class="sxs-lookup"><span data-stu-id="60208-108">1.0</span></span>|
|[<span data-ttu-id="60208-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60208-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60208-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60208-110">ReadItem</span></span>|
|[<span data-ttu-id="60208-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60208-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60208-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60208-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="60208-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="60208-113">Members and methods</span></span>

| <span data-ttu-id="60208-114">成员</span><span class="sxs-lookup"><span data-stu-id="60208-114">Member</span></span> | <span data-ttu-id="60208-115">类型</span><span class="sxs-lookup"><span data-stu-id="60208-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="60208-116">accountType</span><span class="sxs-lookup"><span data-stu-id="60208-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="60208-117">Member</span><span class="sxs-lookup"><span data-stu-id="60208-117">Member</span></span> |
| [<span data-ttu-id="60208-118">displayName</span><span class="sxs-lookup"><span data-stu-id="60208-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="60208-119">Member</span><span class="sxs-lookup"><span data-stu-id="60208-119">Member</span></span> |
| [<span data-ttu-id="60208-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="60208-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="60208-121">Member</span><span class="sxs-lookup"><span data-stu-id="60208-121">Member</span></span> |
| [<span data-ttu-id="60208-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="60208-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="60208-123">Member</span><span class="sxs-lookup"><span data-stu-id="60208-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="60208-124">Members</span><span class="sxs-lookup"><span data-stu-id="60208-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="60208-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="60208-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="60208-126">此成员目前仅在 Outlook 2016 或更高版本 for Mac (内部版本16.9.1212 或更高版本) 中受支持。</span><span class="sxs-lookup"><span data-stu-id="60208-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="60208-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="60208-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="60208-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="60208-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="60208-129">值</span><span class="sxs-lookup"><span data-stu-id="60208-129">Value</span></span> | <span data-ttu-id="60208-130">说明</span><span class="sxs-lookup"><span data-stu-id="60208-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="60208-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="60208-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="60208-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="60208-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="60208-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="60208-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="60208-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="60208-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="60208-135">类型</span><span class="sxs-lookup"><span data-stu-id="60208-135">Type</span></span>

*   <span data-ttu-id="60208-136">String</span><span class="sxs-lookup"><span data-stu-id="60208-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60208-137">要求</span><span class="sxs-lookup"><span data-stu-id="60208-137">Requirements</span></span>

|<span data-ttu-id="60208-138">要求</span><span class="sxs-lookup"><span data-stu-id="60208-138">Requirement</span></span>| <span data-ttu-id="60208-139">值</span><span class="sxs-lookup"><span data-stu-id="60208-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="60208-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60208-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60208-141">1.6</span><span class="sxs-lookup"><span data-stu-id="60208-141">1.6</span></span> |
|[<span data-ttu-id="60208-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60208-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60208-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60208-143">ReadItem</span></span>|
|[<span data-ttu-id="60208-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60208-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60208-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60208-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60208-146">示例</span><span class="sxs-lookup"><span data-stu-id="60208-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="60208-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="60208-147">displayName :String</span></span>

<span data-ttu-id="60208-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="60208-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="60208-149">类型</span><span class="sxs-lookup"><span data-stu-id="60208-149">Type</span></span>

*   <span data-ttu-id="60208-150">String</span><span class="sxs-lookup"><span data-stu-id="60208-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60208-151">要求</span><span class="sxs-lookup"><span data-stu-id="60208-151">Requirements</span></span>

|<span data-ttu-id="60208-152">要求</span><span class="sxs-lookup"><span data-stu-id="60208-152">Requirement</span></span>| <span data-ttu-id="60208-153">值</span><span class="sxs-lookup"><span data-stu-id="60208-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="60208-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60208-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60208-155">1.0</span><span class="sxs-lookup"><span data-stu-id="60208-155">1.0</span></span>|
|[<span data-ttu-id="60208-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60208-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60208-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60208-157">ReadItem</span></span>|
|[<span data-ttu-id="60208-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60208-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60208-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60208-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60208-160">示例</span><span class="sxs-lookup"><span data-stu-id="60208-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="60208-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="60208-161">emailAddress :String</span></span>

<span data-ttu-id="60208-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="60208-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="60208-163">类型</span><span class="sxs-lookup"><span data-stu-id="60208-163">Type</span></span>

*   <span data-ttu-id="60208-164">String</span><span class="sxs-lookup"><span data-stu-id="60208-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60208-165">要求</span><span class="sxs-lookup"><span data-stu-id="60208-165">Requirements</span></span>

|<span data-ttu-id="60208-166">要求</span><span class="sxs-lookup"><span data-stu-id="60208-166">Requirement</span></span>| <span data-ttu-id="60208-167">值</span><span class="sxs-lookup"><span data-stu-id="60208-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="60208-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60208-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60208-169">1.0</span><span class="sxs-lookup"><span data-stu-id="60208-169">1.0</span></span>|
|[<span data-ttu-id="60208-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60208-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60208-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60208-171">ReadItem</span></span>|
|[<span data-ttu-id="60208-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60208-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60208-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60208-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60208-174">示例</span><span class="sxs-lookup"><span data-stu-id="60208-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="60208-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="60208-175">timeZone :String</span></span>

<span data-ttu-id="60208-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="60208-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="60208-177">类型</span><span class="sxs-lookup"><span data-stu-id="60208-177">Type</span></span>

*   <span data-ttu-id="60208-178">String</span><span class="sxs-lookup"><span data-stu-id="60208-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60208-179">要求</span><span class="sxs-lookup"><span data-stu-id="60208-179">Requirements</span></span>

|<span data-ttu-id="60208-180">要求</span><span class="sxs-lookup"><span data-stu-id="60208-180">Requirement</span></span>| <span data-ttu-id="60208-181">值</span><span class="sxs-lookup"><span data-stu-id="60208-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="60208-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="60208-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60208-183">1.0</span><span class="sxs-lookup"><span data-stu-id="60208-183">1.0</span></span>|
|[<span data-ttu-id="60208-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="60208-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60208-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60208-185">ReadItem</span></span>|
|[<span data-ttu-id="60208-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="60208-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60208-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="60208-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60208-188">示例</span><span class="sxs-lookup"><span data-stu-id="60208-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
