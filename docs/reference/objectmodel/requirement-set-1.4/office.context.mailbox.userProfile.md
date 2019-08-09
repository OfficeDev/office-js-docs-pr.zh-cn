---
title: Office.context.mailbox.userProfile - 要求集 1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7a728ebbec0136e0b2eddfb4402e45abe3f02ad4
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268633"
---
# <a name="userprofile"></a><span data-ttu-id="8a5ed-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8a5ed-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8a5ed-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8a5ed-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a5ed-104">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-104">Requirements</span></span>

|<span data-ttu-id="8a5ed-105">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-105">Requirement</span></span>| <span data-ttu-id="8a5ed-106">值</span><span class="sxs-lookup"><span data-stu-id="8a5ed-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a5ed-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8a5ed-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a5ed-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8a5ed-108">1.0</span></span>|
|[<span data-ttu-id="8a5ed-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8a5ed-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a5ed-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a5ed-110">ReadItem</span></span>|
|[<span data-ttu-id="8a5ed-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8a5ed-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a5ed-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8a5ed-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8a5ed-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="8a5ed-113">Members and methods</span></span>

| <span data-ttu-id="8a5ed-114">成员</span><span class="sxs-lookup"><span data-stu-id="8a5ed-114">Member</span></span> | <span data-ttu-id="8a5ed-115">类型</span><span class="sxs-lookup"><span data-stu-id="8a5ed-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8a5ed-116">displayName</span><span class="sxs-lookup"><span data-stu-id="8a5ed-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="8a5ed-117">Member</span><span class="sxs-lookup"><span data-stu-id="8a5ed-117">Member</span></span> |
| [<span data-ttu-id="8a5ed-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="8a5ed-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="8a5ed-119">Member</span><span class="sxs-lookup"><span data-stu-id="8a5ed-119">Member</span></span> |
| [<span data-ttu-id="8a5ed-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="8a5ed-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="8a5ed-121">Member</span><span class="sxs-lookup"><span data-stu-id="8a5ed-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8a5ed-122">Members</span><span class="sxs-lookup"><span data-stu-id="8a5ed-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="8a5ed-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="8a5ed-123">displayName: String</span></span>

<span data-ttu-id="8a5ed-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="8a5ed-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8a5ed-125">类型</span><span class="sxs-lookup"><span data-stu-id="8a5ed-125">Type</span></span>

*   <span data-ttu-id="8a5ed-126">String</span><span class="sxs-lookup"><span data-stu-id="8a5ed-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a5ed-127">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-127">Requirements</span></span>

|<span data-ttu-id="8a5ed-128">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-128">Requirement</span></span>| <span data-ttu-id="8a5ed-129">值</span><span class="sxs-lookup"><span data-stu-id="8a5ed-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a5ed-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8a5ed-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a5ed-131">1.0</span><span class="sxs-lookup"><span data-stu-id="8a5ed-131">1.0</span></span>|
|[<span data-ttu-id="8a5ed-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8a5ed-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a5ed-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a5ed-133">ReadItem</span></span>|
|[<span data-ttu-id="8a5ed-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8a5ed-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a5ed-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8a5ed-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a5ed-136">示例</span><span class="sxs-lookup"><span data-stu-id="8a5ed-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="8a5ed-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="8a5ed-137">emailAddress: String</span></span>

<span data-ttu-id="8a5ed-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="8a5ed-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8a5ed-139">类型</span><span class="sxs-lookup"><span data-stu-id="8a5ed-139">Type</span></span>

*   <span data-ttu-id="8a5ed-140">String</span><span class="sxs-lookup"><span data-stu-id="8a5ed-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a5ed-141">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-141">Requirements</span></span>

|<span data-ttu-id="8a5ed-142">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-142">Requirement</span></span>| <span data-ttu-id="8a5ed-143">值</span><span class="sxs-lookup"><span data-stu-id="8a5ed-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a5ed-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8a5ed-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a5ed-145">1.0</span><span class="sxs-lookup"><span data-stu-id="8a5ed-145">1.0</span></span>|
|[<span data-ttu-id="8a5ed-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8a5ed-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a5ed-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a5ed-147">ReadItem</span></span>|
|[<span data-ttu-id="8a5ed-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8a5ed-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a5ed-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8a5ed-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a5ed-150">示例</span><span class="sxs-lookup"><span data-stu-id="8a5ed-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="8a5ed-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="8a5ed-151">timeZone: String</span></span>

<span data-ttu-id="8a5ed-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="8a5ed-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8a5ed-153">类型</span><span class="sxs-lookup"><span data-stu-id="8a5ed-153">Type</span></span>

*   <span data-ttu-id="8a5ed-154">String</span><span class="sxs-lookup"><span data-stu-id="8a5ed-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a5ed-155">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-155">Requirements</span></span>

|<span data-ttu-id="8a5ed-156">要求</span><span class="sxs-lookup"><span data-stu-id="8a5ed-156">Requirement</span></span>| <span data-ttu-id="8a5ed-157">值</span><span class="sxs-lookup"><span data-stu-id="8a5ed-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a5ed-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8a5ed-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a5ed-159">1.0</span><span class="sxs-lookup"><span data-stu-id="8a5ed-159">1.0</span></span>|
|[<span data-ttu-id="8a5ed-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8a5ed-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a5ed-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a5ed-161">ReadItem</span></span>|
|[<span data-ttu-id="8a5ed-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8a5ed-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a5ed-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8a5ed-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a5ed-164">示例</span><span class="sxs-lookup"><span data-stu-id="8a5ed-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
