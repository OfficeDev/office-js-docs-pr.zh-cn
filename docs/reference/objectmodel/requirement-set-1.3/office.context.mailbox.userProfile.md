---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.3\""
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8924d8b0dfa5bb43be8867cbd0e83ee01ff788cb
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268388"
---
# <a name="userprofile"></a><span data-ttu-id="6b78a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="6b78a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="6b78a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="6b78a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b78a-104">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-104">Requirements</span></span>

|<span data-ttu-id="6b78a-105">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-105">Requirement</span></span>| <span data-ttu-id="6b78a-106">值</span><span class="sxs-lookup"><span data-stu-id="6b78a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b78a-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6b78a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b78a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6b78a-108">1.0</span></span>|
|[<span data-ttu-id="6b78a-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6b78a-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b78a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b78a-110">ReadItem</span></span>|
|[<span data-ttu-id="6b78a-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6b78a-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b78a-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6b78a-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6b78a-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="6b78a-113">Members and methods</span></span>

| <span data-ttu-id="6b78a-114">成员</span><span class="sxs-lookup"><span data-stu-id="6b78a-114">Member</span></span> | <span data-ttu-id="6b78a-115">类型</span><span class="sxs-lookup"><span data-stu-id="6b78a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6b78a-116">displayName</span><span class="sxs-lookup"><span data-stu-id="6b78a-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="6b78a-117">Member</span><span class="sxs-lookup"><span data-stu-id="6b78a-117">Member</span></span> |
| [<span data-ttu-id="6b78a-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="6b78a-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="6b78a-119">Member</span><span class="sxs-lookup"><span data-stu-id="6b78a-119">Member</span></span> |
| [<span data-ttu-id="6b78a-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="6b78a-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="6b78a-121">Member</span><span class="sxs-lookup"><span data-stu-id="6b78a-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="6b78a-122">Members</span><span class="sxs-lookup"><span data-stu-id="6b78a-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="6b78a-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="6b78a-123">displayName: String</span></span>

<span data-ttu-id="6b78a-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="6b78a-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6b78a-125">类型</span><span class="sxs-lookup"><span data-stu-id="6b78a-125">Type</span></span>

*   <span data-ttu-id="6b78a-126">String</span><span class="sxs-lookup"><span data-stu-id="6b78a-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b78a-127">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-127">Requirements</span></span>

|<span data-ttu-id="6b78a-128">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-128">Requirement</span></span>| <span data-ttu-id="6b78a-129">值</span><span class="sxs-lookup"><span data-stu-id="6b78a-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b78a-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6b78a-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b78a-131">1.0</span><span class="sxs-lookup"><span data-stu-id="6b78a-131">1.0</span></span>|
|[<span data-ttu-id="6b78a-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6b78a-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b78a-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b78a-133">ReadItem</span></span>|
|[<span data-ttu-id="6b78a-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6b78a-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b78a-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6b78a-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b78a-136">示例</span><span class="sxs-lookup"><span data-stu-id="6b78a-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="6b78a-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="6b78a-137">emailAddress: String</span></span>

<span data-ttu-id="6b78a-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="6b78a-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6b78a-139">类型</span><span class="sxs-lookup"><span data-stu-id="6b78a-139">Type</span></span>

*   <span data-ttu-id="6b78a-140">String</span><span class="sxs-lookup"><span data-stu-id="6b78a-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b78a-141">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-141">Requirements</span></span>

|<span data-ttu-id="6b78a-142">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-142">Requirement</span></span>| <span data-ttu-id="6b78a-143">值</span><span class="sxs-lookup"><span data-stu-id="6b78a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b78a-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6b78a-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b78a-145">1.0</span><span class="sxs-lookup"><span data-stu-id="6b78a-145">1.0</span></span>|
|[<span data-ttu-id="6b78a-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6b78a-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b78a-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b78a-147">ReadItem</span></span>|
|[<span data-ttu-id="6b78a-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6b78a-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b78a-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6b78a-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b78a-150">示例</span><span class="sxs-lookup"><span data-stu-id="6b78a-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="6b78a-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="6b78a-151">timeZone: String</span></span>

<span data-ttu-id="6b78a-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="6b78a-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6b78a-153">类型</span><span class="sxs-lookup"><span data-stu-id="6b78a-153">Type</span></span>

*   <span data-ttu-id="6b78a-154">String</span><span class="sxs-lookup"><span data-stu-id="6b78a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b78a-155">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-155">Requirements</span></span>

|<span data-ttu-id="6b78a-156">要求</span><span class="sxs-lookup"><span data-stu-id="6b78a-156">Requirement</span></span>| <span data-ttu-id="6b78a-157">值</span><span class="sxs-lookup"><span data-stu-id="6b78a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b78a-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6b78a-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b78a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="6b78a-159">1.0</span></span>|
|[<span data-ttu-id="6b78a-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6b78a-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b78a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b78a-161">ReadItem</span></span>|
|[<span data-ttu-id="6b78a-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6b78a-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b78a-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6b78a-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b78a-164">示例</span><span class="sxs-lookup"><span data-stu-id="6b78a-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
