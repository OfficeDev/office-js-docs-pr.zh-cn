---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.1\""
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: af9a7f790f56124a86af08567690452b7f497408
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268486"
---
# <a name="userprofile"></a><span data-ttu-id="5378d-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="5378d-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="5378d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="5378d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5378d-104">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-104">Requirements</span></span>

|<span data-ttu-id="5378d-105">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-105">Requirement</span></span>| <span data-ttu-id="5378d-106">值</span><span class="sxs-lookup"><span data-stu-id="5378d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5378d-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5378d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5378d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5378d-108">1.0</span></span>|
|[<span data-ttu-id="5378d-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5378d-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5378d-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5378d-110">ReadItem</span></span>|
|[<span data-ttu-id="5378d-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5378d-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5378d-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5378d-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5378d-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="5378d-113">Members and methods</span></span>

| <span data-ttu-id="5378d-114">成员</span><span class="sxs-lookup"><span data-stu-id="5378d-114">Member</span></span> | <span data-ttu-id="5378d-115">类型</span><span class="sxs-lookup"><span data-stu-id="5378d-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5378d-116">displayName</span><span class="sxs-lookup"><span data-stu-id="5378d-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="5378d-117">Member</span><span class="sxs-lookup"><span data-stu-id="5378d-117">Member</span></span> |
| [<span data-ttu-id="5378d-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5378d-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="5378d-119">Member</span><span class="sxs-lookup"><span data-stu-id="5378d-119">Member</span></span> |
| [<span data-ttu-id="5378d-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="5378d-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="5378d-121">Member</span><span class="sxs-lookup"><span data-stu-id="5378d-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="5378d-122">Members</span><span class="sxs-lookup"><span data-stu-id="5378d-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="5378d-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="5378d-123">displayName: String</span></span>

<span data-ttu-id="5378d-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="5378d-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5378d-125">类型</span><span class="sxs-lookup"><span data-stu-id="5378d-125">Type</span></span>

*   <span data-ttu-id="5378d-126">String</span><span class="sxs-lookup"><span data-stu-id="5378d-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5378d-127">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-127">Requirements</span></span>

|<span data-ttu-id="5378d-128">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-128">Requirement</span></span>| <span data-ttu-id="5378d-129">值</span><span class="sxs-lookup"><span data-stu-id="5378d-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="5378d-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5378d-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5378d-131">1.0</span><span class="sxs-lookup"><span data-stu-id="5378d-131">1.0</span></span>|
|[<span data-ttu-id="5378d-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5378d-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5378d-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5378d-133">ReadItem</span></span>|
|[<span data-ttu-id="5378d-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5378d-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5378d-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5378d-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5378d-136">示例</span><span class="sxs-lookup"><span data-stu-id="5378d-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="5378d-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="5378d-137">emailAddress: String</span></span>

<span data-ttu-id="5378d-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="5378d-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5378d-139">类型</span><span class="sxs-lookup"><span data-stu-id="5378d-139">Type</span></span>

*   <span data-ttu-id="5378d-140">String</span><span class="sxs-lookup"><span data-stu-id="5378d-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5378d-141">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-141">Requirements</span></span>

|<span data-ttu-id="5378d-142">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-142">Requirement</span></span>| <span data-ttu-id="5378d-143">值</span><span class="sxs-lookup"><span data-stu-id="5378d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="5378d-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5378d-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5378d-145">1.0</span><span class="sxs-lookup"><span data-stu-id="5378d-145">1.0</span></span>|
|[<span data-ttu-id="5378d-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5378d-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5378d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5378d-147">ReadItem</span></span>|
|[<span data-ttu-id="5378d-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5378d-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5378d-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5378d-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5378d-150">示例</span><span class="sxs-lookup"><span data-stu-id="5378d-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="5378d-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="5378d-151">timeZone: String</span></span>

<span data-ttu-id="5378d-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="5378d-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5378d-153">类型</span><span class="sxs-lookup"><span data-stu-id="5378d-153">Type</span></span>

*   <span data-ttu-id="5378d-154">String</span><span class="sxs-lookup"><span data-stu-id="5378d-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5378d-155">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-155">Requirements</span></span>

|<span data-ttu-id="5378d-156">要求</span><span class="sxs-lookup"><span data-stu-id="5378d-156">Requirement</span></span>| <span data-ttu-id="5378d-157">值</span><span class="sxs-lookup"><span data-stu-id="5378d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="5378d-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5378d-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5378d-159">1.0</span><span class="sxs-lookup"><span data-stu-id="5378d-159">1.0</span></span>|
|[<span data-ttu-id="5378d-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5378d-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5378d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5378d-161">ReadItem</span></span>|
|[<span data-ttu-id="5378d-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5378d-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5378d-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5378d-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5378d-164">示例</span><span class="sxs-lookup"><span data-stu-id="5378d-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
