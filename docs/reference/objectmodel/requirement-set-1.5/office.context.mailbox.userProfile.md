---
title: Office.context.mailbox.userProfile - 要求集 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e98e88cde184db121e69fdd267dff4e39d887b1f
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067825"
---
# <a name="userprofile"></a><span data-ttu-id="a6606-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a6606-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a6606-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a6606-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6606-104">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-104">Requirements</span></span>

|<span data-ttu-id="a6606-105">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-105">Requirement</span></span>| <span data-ttu-id="a6606-106">值</span><span class="sxs-lookup"><span data-stu-id="a6606-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6606-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a6606-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6606-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a6606-108">1.0</span></span>|
|[<span data-ttu-id="a6606-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a6606-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6606-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6606-110">ReadItem</span></span>|
|[<span data-ttu-id="a6606-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a6606-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6606-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a6606-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a6606-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="a6606-113">Members and methods</span></span>

| <span data-ttu-id="a6606-114">成员</span><span class="sxs-lookup"><span data-stu-id="a6606-114">Member</span></span> | <span data-ttu-id="a6606-115">类型</span><span class="sxs-lookup"><span data-stu-id="a6606-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a6606-116">displayName</span><span class="sxs-lookup"><span data-stu-id="a6606-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="a6606-117">成员</span><span class="sxs-lookup"><span data-stu-id="a6606-117">Member</span></span> |
| [<span data-ttu-id="a6606-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a6606-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a6606-119">成员</span><span class="sxs-lookup"><span data-stu-id="a6606-119">Member</span></span> |
| [<span data-ttu-id="a6606-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="a6606-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a6606-121">成员</span><span class="sxs-lookup"><span data-stu-id="a6606-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a6606-122">成员</span><span class="sxs-lookup"><span data-stu-id="a6606-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a6606-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a6606-123">displayName :String</span></span>

<span data-ttu-id="a6606-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="a6606-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a6606-125">Type</span><span class="sxs-lookup"><span data-stu-id="a6606-125">Type</span></span>

*   <span data-ttu-id="a6606-126">String</span><span class="sxs-lookup"><span data-stu-id="a6606-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6606-127">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-127">Requirements</span></span>

|<span data-ttu-id="a6606-128">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-128">Requirement</span></span>| <span data-ttu-id="a6606-129">值</span><span class="sxs-lookup"><span data-stu-id="a6606-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6606-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a6606-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6606-131">1.0</span><span class="sxs-lookup"><span data-stu-id="a6606-131">1.0</span></span>|
|[<span data-ttu-id="a6606-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a6606-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6606-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6606-133">ReadItem</span></span>|
|[<span data-ttu-id="a6606-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a6606-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6606-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a6606-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6606-136">示例</span><span class="sxs-lookup"><span data-stu-id="a6606-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a6606-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a6606-137">emailAddress :String</span></span>

<span data-ttu-id="a6606-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a6606-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a6606-139">Type</span><span class="sxs-lookup"><span data-stu-id="a6606-139">Type</span></span>

*   <span data-ttu-id="a6606-140">String</span><span class="sxs-lookup"><span data-stu-id="a6606-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6606-141">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-141">Requirements</span></span>

|<span data-ttu-id="a6606-142">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-142">Requirement</span></span>| <span data-ttu-id="a6606-143">值</span><span class="sxs-lookup"><span data-stu-id="a6606-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6606-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a6606-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6606-145">1.0</span><span class="sxs-lookup"><span data-stu-id="a6606-145">1.0</span></span>|
|[<span data-ttu-id="a6606-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a6606-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6606-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6606-147">ReadItem</span></span>|
|[<span data-ttu-id="a6606-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a6606-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6606-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a6606-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6606-150">示例</span><span class="sxs-lookup"><span data-stu-id="a6606-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a6606-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a6606-151">timeZone :String</span></span>

<span data-ttu-id="a6606-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="a6606-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a6606-153">Type</span><span class="sxs-lookup"><span data-stu-id="a6606-153">Type</span></span>

*   <span data-ttu-id="a6606-154">String</span><span class="sxs-lookup"><span data-stu-id="a6606-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6606-155">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-155">Requirements</span></span>

|<span data-ttu-id="a6606-156">要求</span><span class="sxs-lookup"><span data-stu-id="a6606-156">Requirement</span></span>| <span data-ttu-id="a6606-157">值</span><span class="sxs-lookup"><span data-stu-id="a6606-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6606-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a6606-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6606-159">1.0</span><span class="sxs-lookup"><span data-stu-id="a6606-159">1.0</span></span>|
|[<span data-ttu-id="a6606-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a6606-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6606-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6606-161">ReadItem</span></span>|
|[<span data-ttu-id="a6606-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a6606-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6606-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a6606-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6606-164">示例</span><span class="sxs-lookup"><span data-stu-id="a6606-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
