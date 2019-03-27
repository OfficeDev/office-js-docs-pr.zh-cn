---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.5\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871015"
---
# <a name="userprofile"></a><span data-ttu-id="38147-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="38147-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="38147-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="38147-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="38147-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="38147-104">Requirements</span></span>

|<span data-ttu-id="38147-105">要求</span><span class="sxs-lookup"><span data-stu-id="38147-105">Requirement</span></span>| <span data-ttu-id="38147-106">值</span><span class="sxs-lookup"><span data-stu-id="38147-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="38147-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="38147-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38147-108">1.0</span><span class="sxs-lookup"><span data-stu-id="38147-108">1.0</span></span>|
|[<span data-ttu-id="38147-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="38147-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38147-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38147-110">ReadItem</span></span>|
|[<span data-ttu-id="38147-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="38147-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="38147-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="38147-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="38147-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="38147-113">Members and methods</span></span>

| <span data-ttu-id="38147-114">成员</span><span class="sxs-lookup"><span data-stu-id="38147-114">Member</span></span> | <span data-ttu-id="38147-115">类型</span><span class="sxs-lookup"><span data-stu-id="38147-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="38147-116">displayName</span><span class="sxs-lookup"><span data-stu-id="38147-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="38147-117">Member</span><span class="sxs-lookup"><span data-stu-id="38147-117">Member</span></span> |
| [<span data-ttu-id="38147-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="38147-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="38147-119">Member</span><span class="sxs-lookup"><span data-stu-id="38147-119">Member</span></span> |
| [<span data-ttu-id="38147-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="38147-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="38147-121">Member</span><span class="sxs-lookup"><span data-stu-id="38147-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="38147-122">成员</span><span class="sxs-lookup"><span data-stu-id="38147-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="38147-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="38147-123">displayName :String</span></span>

<span data-ttu-id="38147-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="38147-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="38147-125">类型</span><span class="sxs-lookup"><span data-stu-id="38147-125">Type</span></span>

*   <span data-ttu-id="38147-126">String</span><span class="sxs-lookup"><span data-stu-id="38147-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="38147-127">要求</span><span class="sxs-lookup"><span data-stu-id="38147-127">Requirements</span></span>

|<span data-ttu-id="38147-128">要求</span><span class="sxs-lookup"><span data-stu-id="38147-128">Requirement</span></span>| <span data-ttu-id="38147-129">值</span><span class="sxs-lookup"><span data-stu-id="38147-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="38147-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="38147-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38147-131">1.0</span><span class="sxs-lookup"><span data-stu-id="38147-131">1.0</span></span>|
|[<span data-ttu-id="38147-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="38147-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38147-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38147-133">ReadItem</span></span>|
|[<span data-ttu-id="38147-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="38147-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="38147-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="38147-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38147-136">示例</span><span class="sxs-lookup"><span data-stu-id="38147-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="38147-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="38147-137">emailAddress :String</span></span>

<span data-ttu-id="38147-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="38147-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="38147-139">类型</span><span class="sxs-lookup"><span data-stu-id="38147-139">Type</span></span>

*   <span data-ttu-id="38147-140">String</span><span class="sxs-lookup"><span data-stu-id="38147-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="38147-141">要求</span><span class="sxs-lookup"><span data-stu-id="38147-141">Requirements</span></span>

|<span data-ttu-id="38147-142">要求</span><span class="sxs-lookup"><span data-stu-id="38147-142">Requirement</span></span>| <span data-ttu-id="38147-143">值</span><span class="sxs-lookup"><span data-stu-id="38147-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="38147-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="38147-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38147-145">1.0</span><span class="sxs-lookup"><span data-stu-id="38147-145">1.0</span></span>|
|[<span data-ttu-id="38147-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="38147-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38147-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38147-147">ReadItem</span></span>|
|[<span data-ttu-id="38147-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="38147-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="38147-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="38147-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38147-150">示例</span><span class="sxs-lookup"><span data-stu-id="38147-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="38147-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="38147-151">timeZone :String</span></span>

<span data-ttu-id="38147-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="38147-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="38147-153">类型</span><span class="sxs-lookup"><span data-stu-id="38147-153">Type</span></span>

*   <span data-ttu-id="38147-154">String</span><span class="sxs-lookup"><span data-stu-id="38147-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="38147-155">要求</span><span class="sxs-lookup"><span data-stu-id="38147-155">Requirements</span></span>

|<span data-ttu-id="38147-156">要求</span><span class="sxs-lookup"><span data-stu-id="38147-156">Requirement</span></span>| <span data-ttu-id="38147-157">值</span><span class="sxs-lookup"><span data-stu-id="38147-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="38147-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="38147-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38147-159">1.0</span><span class="sxs-lookup"><span data-stu-id="38147-159">1.0</span></span>|
|[<span data-ttu-id="38147-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="38147-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38147-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38147-161">ReadItem</span></span>|
|[<span data-ttu-id="38147-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="38147-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="38147-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="38147-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38147-164">示例</span><span class="sxs-lookup"><span data-stu-id="38147-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
