---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.2\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 496a59f4ef02f03cda95fde0bf14634b1db13f77
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450336"
---
# <a name="userprofile"></a><span data-ttu-id="9bc9b-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="9bc9b-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="9bc9b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="9bc9b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="9bc9b-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="9bc9b-104">Requirements</span></span>

|<span data-ttu-id="9bc9b-105">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-105">Requirement</span></span>| <span data-ttu-id="9bc9b-106">值</span><span class="sxs-lookup"><span data-stu-id="9bc9b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bc9b-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9bc9b-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bc9b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9bc9b-108">1.0</span></span>|
|[<span data-ttu-id="9bc9b-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9bc9b-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9bc9b-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9bc9b-110">ReadItem</span></span>|
|[<span data-ttu-id="9bc9b-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9bc9b-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9bc9b-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9bc9b-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="9bc9b-113">成员</span><span class="sxs-lookup"><span data-stu-id="9bc9b-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="9bc9b-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="9bc9b-114">displayName :String</span></span>

<span data-ttu-id="9bc9b-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="9bc9b-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="9bc9b-116">类型</span><span class="sxs-lookup"><span data-stu-id="9bc9b-116">Type</span></span>

*   <span data-ttu-id="9bc9b-117">String</span><span class="sxs-lookup"><span data-stu-id="9bc9b-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9bc9b-118">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-118">Requirements</span></span>

|<span data-ttu-id="9bc9b-119">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-119">Requirement</span></span>| <span data-ttu-id="9bc9b-120">值</span><span class="sxs-lookup"><span data-stu-id="9bc9b-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bc9b-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9bc9b-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bc9b-122">1.0</span><span class="sxs-lookup"><span data-stu-id="9bc9b-122">1.0</span></span>|
|[<span data-ttu-id="9bc9b-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9bc9b-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9bc9b-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9bc9b-124">ReadItem</span></span>|
|[<span data-ttu-id="9bc9b-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9bc9b-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9bc9b-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9bc9b-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9bc9b-127">示例</span><span class="sxs-lookup"><span data-stu-id="9bc9b-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="9bc9b-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="9bc9b-128">emailAddress :String</span></span>

<span data-ttu-id="9bc9b-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="9bc9b-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="9bc9b-130">类型</span><span class="sxs-lookup"><span data-stu-id="9bc9b-130">Type</span></span>

*   <span data-ttu-id="9bc9b-131">String</span><span class="sxs-lookup"><span data-stu-id="9bc9b-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9bc9b-132">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-132">Requirements</span></span>

|<span data-ttu-id="9bc9b-133">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-133">Requirement</span></span>| <span data-ttu-id="9bc9b-134">值</span><span class="sxs-lookup"><span data-stu-id="9bc9b-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bc9b-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9bc9b-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bc9b-136">1.0</span><span class="sxs-lookup"><span data-stu-id="9bc9b-136">1.0</span></span>|
|[<span data-ttu-id="9bc9b-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9bc9b-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9bc9b-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9bc9b-138">ReadItem</span></span>|
|[<span data-ttu-id="9bc9b-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9bc9b-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9bc9b-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9bc9b-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9bc9b-141">示例</span><span class="sxs-lookup"><span data-stu-id="9bc9b-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="9bc9b-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="9bc9b-142">timeZone :String</span></span>

<span data-ttu-id="9bc9b-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="9bc9b-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="9bc9b-144">类型</span><span class="sxs-lookup"><span data-stu-id="9bc9b-144">Type</span></span>

*   <span data-ttu-id="9bc9b-145">String</span><span class="sxs-lookup"><span data-stu-id="9bc9b-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9bc9b-146">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-146">Requirements</span></span>

|<span data-ttu-id="9bc9b-147">要求</span><span class="sxs-lookup"><span data-stu-id="9bc9b-147">Requirement</span></span>| <span data-ttu-id="9bc9b-148">值</span><span class="sxs-lookup"><span data-stu-id="9bc9b-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bc9b-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9bc9b-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bc9b-150">1.0</span><span class="sxs-lookup"><span data-stu-id="9bc9b-150">1.0</span></span>|
|[<span data-ttu-id="9bc9b-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9bc9b-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9bc9b-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9bc9b-152">ReadItem</span></span>|
|[<span data-ttu-id="9bc9b-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9bc9b-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9bc9b-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9bc9b-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9bc9b-155">示例</span><span class="sxs-lookup"><span data-stu-id="9bc9b-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
