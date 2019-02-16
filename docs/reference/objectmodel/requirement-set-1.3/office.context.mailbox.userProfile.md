---
title: Office.context.mailbox.userProfile - 要求集 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e29cf90d1c5d4c288417ef98f6e9d22eaf908b67
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067928"
---
# <a name="userprofile"></a><span data-ttu-id="80e83-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="80e83-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="80e83-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="80e83-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="80e83-104">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-104">Requirements</span></span>

|<span data-ttu-id="80e83-105">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-105">Requirement</span></span>| <span data-ttu-id="80e83-106">值</span><span class="sxs-lookup"><span data-stu-id="80e83-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="80e83-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="80e83-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80e83-108">1.0</span><span class="sxs-lookup"><span data-stu-id="80e83-108">1.0</span></span>|
|[<span data-ttu-id="80e83-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="80e83-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80e83-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80e83-110">ReadItem</span></span>|
|[<span data-ttu-id="80e83-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="80e83-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80e83-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="80e83-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="80e83-113">成员</span><span class="sxs-lookup"><span data-stu-id="80e83-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="80e83-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="80e83-114">displayName :String</span></span>

<span data-ttu-id="80e83-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="80e83-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="80e83-116">Type</span><span class="sxs-lookup"><span data-stu-id="80e83-116">Type</span></span>

*   <span data-ttu-id="80e83-117">String</span><span class="sxs-lookup"><span data-stu-id="80e83-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80e83-118">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-118">Requirements</span></span>

|<span data-ttu-id="80e83-119">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-119">Requirement</span></span>| <span data-ttu-id="80e83-120">值</span><span class="sxs-lookup"><span data-stu-id="80e83-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="80e83-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="80e83-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80e83-122">1.0</span><span class="sxs-lookup"><span data-stu-id="80e83-122">1.0</span></span>|
|[<span data-ttu-id="80e83-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="80e83-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80e83-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80e83-124">ReadItem</span></span>|
|[<span data-ttu-id="80e83-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="80e83-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80e83-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="80e83-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80e83-127">示例</span><span class="sxs-lookup"><span data-stu-id="80e83-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="80e83-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="80e83-128">emailAddress :String</span></span>

<span data-ttu-id="80e83-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="80e83-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="80e83-130">Type</span><span class="sxs-lookup"><span data-stu-id="80e83-130">Type</span></span>

*   <span data-ttu-id="80e83-131">String</span><span class="sxs-lookup"><span data-stu-id="80e83-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80e83-132">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-132">Requirements</span></span>

|<span data-ttu-id="80e83-133">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-133">Requirement</span></span>| <span data-ttu-id="80e83-134">值</span><span class="sxs-lookup"><span data-stu-id="80e83-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="80e83-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="80e83-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80e83-136">1.0</span><span class="sxs-lookup"><span data-stu-id="80e83-136">1.0</span></span>|
|[<span data-ttu-id="80e83-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="80e83-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80e83-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80e83-138">ReadItem</span></span>|
|[<span data-ttu-id="80e83-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="80e83-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80e83-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="80e83-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80e83-141">示例</span><span class="sxs-lookup"><span data-stu-id="80e83-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="80e83-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="80e83-142">timeZone :String</span></span>

<span data-ttu-id="80e83-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="80e83-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="80e83-144">Type</span><span class="sxs-lookup"><span data-stu-id="80e83-144">Type</span></span>

*   <span data-ttu-id="80e83-145">String</span><span class="sxs-lookup"><span data-stu-id="80e83-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80e83-146">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-146">Requirements</span></span>

|<span data-ttu-id="80e83-147">要求</span><span class="sxs-lookup"><span data-stu-id="80e83-147">Requirement</span></span>| <span data-ttu-id="80e83-148">值</span><span class="sxs-lookup"><span data-stu-id="80e83-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="80e83-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="80e83-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80e83-150">1.0</span><span class="sxs-lookup"><span data-stu-id="80e83-150">1.0</span></span>|
|[<span data-ttu-id="80e83-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="80e83-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80e83-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="80e83-152">ReadItem</span></span>|
|[<span data-ttu-id="80e83-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="80e83-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80e83-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="80e83-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80e83-155">示例</span><span class="sxs-lookup"><span data-stu-id="80e83-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
