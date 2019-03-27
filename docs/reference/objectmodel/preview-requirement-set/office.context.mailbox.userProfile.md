---
title: "\"context.subname\": \"邮箱. userProfile-预览要求集\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 204097497c958c26a6e67fc01d6dbd5142d8dced
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871869"
---
# <a name="userprofile"></a><span data-ttu-id="a029f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a029f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a029f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a029f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a029f-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="a029f-104">Requirements</span></span>

|<span data-ttu-id="a029f-105">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-105">Requirement</span></span>| <span data-ttu-id="a029f-106">值</span><span class="sxs-lookup"><span data-stu-id="a029f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a029f-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a029f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a029f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a029f-108">1.0</span></span>|
|[<span data-ttu-id="a029f-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a029f-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a029f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a029f-110">ReadItem</span></span>|
|[<span data-ttu-id="a029f-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a029f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a029f-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a029f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a029f-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="a029f-113">Members and methods</span></span>

| <span data-ttu-id="a029f-114">成员</span><span class="sxs-lookup"><span data-stu-id="a029f-114">Member</span></span> | <span data-ttu-id="a029f-115">类型</span><span class="sxs-lookup"><span data-stu-id="a029f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a029f-116">accountType</span><span class="sxs-lookup"><span data-stu-id="a029f-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="a029f-117">Member</span><span class="sxs-lookup"><span data-stu-id="a029f-117">Member</span></span> |
| [<span data-ttu-id="a029f-118">displayName</span><span class="sxs-lookup"><span data-stu-id="a029f-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="a029f-119">Member</span><span class="sxs-lookup"><span data-stu-id="a029f-119">Member</span></span> |
| [<span data-ttu-id="a029f-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a029f-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a029f-121">Member</span><span class="sxs-lookup"><span data-stu-id="a029f-121">Member</span></span> |
| [<span data-ttu-id="a029f-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="a029f-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a029f-123">Member</span><span class="sxs-lookup"><span data-stu-id="a029f-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a029f-124">Members</span><span class="sxs-lookup"><span data-stu-id="a029f-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="a029f-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="a029f-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="a029f-126">此成员目前仅在 Outlook 2016 或更高版本 for Mac (内部版本16.9.1212 或更高版本) 中受支持。</span><span class="sxs-lookup"><span data-stu-id="a029f-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="a029f-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="a029f-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="a029f-128">下表中列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="a029f-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="a029f-129">值</span><span class="sxs-lookup"><span data-stu-id="a029f-129">Value</span></span> | <span data-ttu-id="a029f-130">说明</span><span class="sxs-lookup"><span data-stu-id="a029f-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="a029f-131">邮箱位于本地 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="a029f-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="a029f-132">邮箱与 Gmail 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="a029f-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="a029f-133">邮箱与 Office 365 工作或学校帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="a029f-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="a029f-134">邮箱与个人 Outlook.com 帐户相关联。</span><span class="sxs-lookup"><span data-stu-id="a029f-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="a029f-135">类型</span><span class="sxs-lookup"><span data-stu-id="a029f-135">Type</span></span>

*   <span data-ttu-id="a029f-136">String</span><span class="sxs-lookup"><span data-stu-id="a029f-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a029f-137">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-137">Requirements</span></span>

|<span data-ttu-id="a029f-138">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-138">Requirement</span></span>| <span data-ttu-id="a029f-139">值</span><span class="sxs-lookup"><span data-stu-id="a029f-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="a029f-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a029f-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a029f-141">1.6</span><span class="sxs-lookup"><span data-stu-id="a029f-141">1.6</span></span> |
|[<span data-ttu-id="a029f-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a029f-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a029f-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a029f-143">ReadItem</span></span>|
|[<span data-ttu-id="a029f-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a029f-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a029f-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a029f-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a029f-146">示例</span><span class="sxs-lookup"><span data-stu-id="a029f-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="a029f-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a029f-147">displayName :String</span></span>

<span data-ttu-id="a029f-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="a029f-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a029f-149">类型</span><span class="sxs-lookup"><span data-stu-id="a029f-149">Type</span></span>

*   <span data-ttu-id="a029f-150">String</span><span class="sxs-lookup"><span data-stu-id="a029f-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a029f-151">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-151">Requirements</span></span>

|<span data-ttu-id="a029f-152">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-152">Requirement</span></span>| <span data-ttu-id="a029f-153">值</span><span class="sxs-lookup"><span data-stu-id="a029f-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="a029f-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a029f-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a029f-155">1.0</span><span class="sxs-lookup"><span data-stu-id="a029f-155">1.0</span></span>|
|[<span data-ttu-id="a029f-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a029f-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a029f-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a029f-157">ReadItem</span></span>|
|[<span data-ttu-id="a029f-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a029f-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a029f-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a029f-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a029f-160">示例</span><span class="sxs-lookup"><span data-stu-id="a029f-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a029f-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a029f-161">emailAddress :String</span></span>

<span data-ttu-id="a029f-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a029f-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a029f-163">类型</span><span class="sxs-lookup"><span data-stu-id="a029f-163">Type</span></span>

*   <span data-ttu-id="a029f-164">String</span><span class="sxs-lookup"><span data-stu-id="a029f-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a029f-165">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-165">Requirements</span></span>

|<span data-ttu-id="a029f-166">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-166">Requirement</span></span>| <span data-ttu-id="a029f-167">值</span><span class="sxs-lookup"><span data-stu-id="a029f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a029f-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a029f-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a029f-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a029f-169">1.0</span></span>|
|[<span data-ttu-id="a029f-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a029f-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a029f-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a029f-171">ReadItem</span></span>|
|[<span data-ttu-id="a029f-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a029f-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a029f-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a029f-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a029f-174">示例</span><span class="sxs-lookup"><span data-stu-id="a029f-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a029f-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a029f-175">timeZone :String</span></span>

<span data-ttu-id="a029f-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="a029f-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a029f-177">类型</span><span class="sxs-lookup"><span data-stu-id="a029f-177">Type</span></span>

*   <span data-ttu-id="a029f-178">String</span><span class="sxs-lookup"><span data-stu-id="a029f-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a029f-179">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-179">Requirements</span></span>

|<span data-ttu-id="a029f-180">要求</span><span class="sxs-lookup"><span data-stu-id="a029f-180">Requirement</span></span>| <span data-ttu-id="a029f-181">值</span><span class="sxs-lookup"><span data-stu-id="a029f-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a029f-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a029f-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a029f-183">1.0</span><span class="sxs-lookup"><span data-stu-id="a029f-183">1.0</span></span>|
|[<span data-ttu-id="a029f-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a029f-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a029f-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a029f-185">ReadItem</span></span>|
|[<span data-ttu-id="a029f-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a029f-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a029f-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a029f-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a029f-188">示例</span><span class="sxs-lookup"><span data-stu-id="a029f-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
