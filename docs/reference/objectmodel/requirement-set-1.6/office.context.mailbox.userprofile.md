---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.6\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 3ca06925dcd37d8e68f086daf4705b10fb936623
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127203"
---
# <a name="userprofile"></a><span data-ttu-id="1b806-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="1b806-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="1b806-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="1b806-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b806-104">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-104">Requirements</span></span>

|<span data-ttu-id="1b806-105">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-105">Requirement</span></span>| <span data-ttu-id="1b806-106">值</span><span class="sxs-lookup"><span data-stu-id="1b806-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b806-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1b806-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b806-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1b806-108">1.0</span></span>|
|[<span data-ttu-id="1b806-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1b806-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b806-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b806-110">ReadItem</span></span>|
|[<span data-ttu-id="1b806-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1b806-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b806-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1b806-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1b806-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="1b806-113">Members and methods</span></span>

| <span data-ttu-id="1b806-114">成员</span><span class="sxs-lookup"><span data-stu-id="1b806-114">Member</span></span> | <span data-ttu-id="1b806-115">类型</span><span class="sxs-lookup"><span data-stu-id="1b806-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1b806-116">accountType</span><span class="sxs-lookup"><span data-stu-id="1b806-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="1b806-117">Member</span><span class="sxs-lookup"><span data-stu-id="1b806-117">Member</span></span> |
| [<span data-ttu-id="1b806-118">displayName</span><span class="sxs-lookup"><span data-stu-id="1b806-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="1b806-119">Member</span><span class="sxs-lookup"><span data-stu-id="1b806-119">Member</span></span> |
| [<span data-ttu-id="1b806-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="1b806-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="1b806-121">Member</span><span class="sxs-lookup"><span data-stu-id="1b806-121">Member</span></span> |
| [<span data-ttu-id="1b806-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="1b806-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="1b806-123">Member</span><span class="sxs-lookup"><span data-stu-id="1b806-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1b806-124">Members</span><span class="sxs-lookup"><span data-stu-id="1b806-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="1b806-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="1b806-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="1b806-126">此成员目前仅在 Outlook 2016 或更高版本 (内部版本16.9.1212 或更高版本) 中受支持。</span><span class="sxs-lookup"><span data-stu-id="1b806-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="1b806-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="1b806-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="1b806-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="1b806-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="1b806-129">值</span><span class="sxs-lookup"><span data-stu-id="1b806-129">Value</span></span> | <span data-ttu-id="1b806-130">说明</span><span class="sxs-lookup"><span data-stu-id="1b806-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="1b806-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="1b806-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="1b806-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="1b806-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="1b806-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="1b806-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="1b806-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="1b806-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="1b806-135">类型</span><span class="sxs-lookup"><span data-stu-id="1b806-135">Type</span></span>

*   <span data-ttu-id="1b806-136">String</span><span class="sxs-lookup"><span data-stu-id="1b806-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b806-137">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-137">Requirements</span></span>

|<span data-ttu-id="1b806-138">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-138">Requirement</span></span>| <span data-ttu-id="1b806-139">值</span><span class="sxs-lookup"><span data-stu-id="1b806-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b806-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1b806-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b806-141">1.6</span><span class="sxs-lookup"><span data-stu-id="1b806-141">1.6</span></span> |
|[<span data-ttu-id="1b806-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1b806-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b806-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b806-143">ReadItem</span></span>|
|[<span data-ttu-id="1b806-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1b806-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b806-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1b806-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b806-146">示例</span><span class="sxs-lookup"><span data-stu-id="1b806-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

#### <a name="displayname-string"></a><span data-ttu-id="1b806-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="1b806-147">displayName: String</span></span>

<span data-ttu-id="1b806-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="1b806-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="1b806-149">类型</span><span class="sxs-lookup"><span data-stu-id="1b806-149">Type</span></span>

*   <span data-ttu-id="1b806-150">String</span><span class="sxs-lookup"><span data-stu-id="1b806-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b806-151">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-151">Requirements</span></span>

|<span data-ttu-id="1b806-152">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-152">Requirement</span></span>| <span data-ttu-id="1b806-153">值</span><span class="sxs-lookup"><span data-stu-id="1b806-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b806-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1b806-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b806-155">1.0</span><span class="sxs-lookup"><span data-stu-id="1b806-155">1.0</span></span>|
|[<span data-ttu-id="1b806-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1b806-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b806-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b806-157">ReadItem</span></span>|
|[<span data-ttu-id="1b806-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1b806-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b806-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1b806-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b806-160">示例</span><span class="sxs-lookup"><span data-stu-id="1b806-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="1b806-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="1b806-161">emailAddress: String</span></span>

<span data-ttu-id="1b806-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="1b806-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="1b806-163">类型</span><span class="sxs-lookup"><span data-stu-id="1b806-163">Type</span></span>

*   <span data-ttu-id="1b806-164">String</span><span class="sxs-lookup"><span data-stu-id="1b806-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b806-165">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-165">Requirements</span></span>

|<span data-ttu-id="1b806-166">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-166">Requirement</span></span>| <span data-ttu-id="1b806-167">值</span><span class="sxs-lookup"><span data-stu-id="1b806-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b806-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1b806-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b806-169">1.0</span><span class="sxs-lookup"><span data-stu-id="1b806-169">1.0</span></span>|
|[<span data-ttu-id="1b806-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1b806-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b806-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b806-171">ReadItem</span></span>|
|[<span data-ttu-id="1b806-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1b806-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b806-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1b806-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b806-174">示例</span><span class="sxs-lookup"><span data-stu-id="1b806-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="1b806-175">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="1b806-175">timeZone: String</span></span>

<span data-ttu-id="1b806-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="1b806-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="1b806-177">类型</span><span class="sxs-lookup"><span data-stu-id="1b806-177">Type</span></span>

*   <span data-ttu-id="1b806-178">String</span><span class="sxs-lookup"><span data-stu-id="1b806-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b806-179">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-179">Requirements</span></span>

|<span data-ttu-id="1b806-180">要求</span><span class="sxs-lookup"><span data-stu-id="1b806-180">Requirement</span></span>| <span data-ttu-id="1b806-181">值</span><span class="sxs-lookup"><span data-stu-id="1b806-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b806-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1b806-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b806-183">1.0</span><span class="sxs-lookup"><span data-stu-id="1b806-183">1.0</span></span>|
|[<span data-ttu-id="1b806-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1b806-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b806-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b806-185">ReadItem</span></span>|
|[<span data-ttu-id="1b806-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1b806-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b806-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1b806-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b806-188">示例</span><span class="sxs-lookup"><span data-stu-id="1b806-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
