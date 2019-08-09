---
title: "\"Context.subname\": \"邮箱. userProfile-要求集 1.2\""
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7258195e7ec0ef2432723d0f32f3d9ef1a3acf2b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268682"
---
# <a name="userprofile"></a><span data-ttu-id="c0837-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c0837-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c0837-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c0837-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0837-104">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-104">Requirements</span></span>

|<span data-ttu-id="c0837-105">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-105">Requirement</span></span>| <span data-ttu-id="c0837-106">值</span><span class="sxs-lookup"><span data-stu-id="c0837-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0837-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0837-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0837-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c0837-108">1.0</span></span>|
|[<span data-ttu-id="c0837-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0837-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0837-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0837-110">ReadItem</span></span>|
|[<span data-ttu-id="c0837-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0837-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0837-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0837-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c0837-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c0837-113">Members and methods</span></span>

| <span data-ttu-id="c0837-114">成员</span><span class="sxs-lookup"><span data-stu-id="c0837-114">Member</span></span> | <span data-ttu-id="c0837-115">类型</span><span class="sxs-lookup"><span data-stu-id="c0837-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c0837-116">displayName</span><span class="sxs-lookup"><span data-stu-id="c0837-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="c0837-117">Member</span><span class="sxs-lookup"><span data-stu-id="c0837-117">Member</span></span> |
| [<span data-ttu-id="c0837-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c0837-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c0837-119">Member</span><span class="sxs-lookup"><span data-stu-id="c0837-119">Member</span></span> |
| [<span data-ttu-id="c0837-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="c0837-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c0837-121">Member</span><span class="sxs-lookup"><span data-stu-id="c0837-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c0837-122">Members</span><span class="sxs-lookup"><span data-stu-id="c0837-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="c0837-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="c0837-123">displayName: String</span></span>

<span data-ttu-id="c0837-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="c0837-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c0837-125">类型</span><span class="sxs-lookup"><span data-stu-id="c0837-125">Type</span></span>

*   <span data-ttu-id="c0837-126">String</span><span class="sxs-lookup"><span data-stu-id="c0837-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0837-127">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-127">Requirements</span></span>

|<span data-ttu-id="c0837-128">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-128">Requirement</span></span>| <span data-ttu-id="c0837-129">值</span><span class="sxs-lookup"><span data-stu-id="c0837-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0837-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0837-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0837-131">1.0</span><span class="sxs-lookup"><span data-stu-id="c0837-131">1.0</span></span>|
|[<span data-ttu-id="c0837-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0837-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0837-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0837-133">ReadItem</span></span>|
|[<span data-ttu-id="c0837-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0837-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0837-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0837-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0837-136">示例</span><span class="sxs-lookup"><span data-stu-id="c0837-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="c0837-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="c0837-137">emailAddress: String</span></span>

<span data-ttu-id="c0837-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c0837-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c0837-139">类型</span><span class="sxs-lookup"><span data-stu-id="c0837-139">Type</span></span>

*   <span data-ttu-id="c0837-140">String</span><span class="sxs-lookup"><span data-stu-id="c0837-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0837-141">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-141">Requirements</span></span>

|<span data-ttu-id="c0837-142">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-142">Requirement</span></span>| <span data-ttu-id="c0837-143">值</span><span class="sxs-lookup"><span data-stu-id="c0837-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0837-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0837-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0837-145">1.0</span><span class="sxs-lookup"><span data-stu-id="c0837-145">1.0</span></span>|
|[<span data-ttu-id="c0837-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0837-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0837-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0837-147">ReadItem</span></span>|
|[<span data-ttu-id="c0837-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0837-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0837-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0837-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0837-150">示例</span><span class="sxs-lookup"><span data-stu-id="c0837-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="c0837-151">时区: 字符串</span><span class="sxs-lookup"><span data-stu-id="c0837-151">timeZone: String</span></span>

<span data-ttu-id="c0837-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="c0837-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c0837-153">类型</span><span class="sxs-lookup"><span data-stu-id="c0837-153">Type</span></span>

*   <span data-ttu-id="c0837-154">String</span><span class="sxs-lookup"><span data-stu-id="c0837-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0837-155">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-155">Requirements</span></span>

|<span data-ttu-id="c0837-156">要求</span><span class="sxs-lookup"><span data-stu-id="c0837-156">Requirement</span></span>| <span data-ttu-id="c0837-157">值</span><span class="sxs-lookup"><span data-stu-id="c0837-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0837-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0837-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0837-159">1.0</span><span class="sxs-lookup"><span data-stu-id="c0837-159">1.0</span></span>|
|[<span data-ttu-id="c0837-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c0837-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0837-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0837-161">ReadItem</span></span>|
|[<span data-ttu-id="c0837-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0837-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0837-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0837-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0837-164">示例</span><span class="sxs-lookup"><span data-stu-id="c0837-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
