---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.3\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 03cdc13845bff0fbd3855f29f43298cd770e5ad9
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30869909"
---
# <a name="userprofile"></a><span data-ttu-id="85380-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="85380-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="85380-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="85380-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="85380-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="85380-104">Requirements</span></span>

|<span data-ttu-id="85380-105">要求</span><span class="sxs-lookup"><span data-stu-id="85380-105">Requirement</span></span>| <span data-ttu-id="85380-106">值</span><span class="sxs-lookup"><span data-stu-id="85380-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="85380-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85380-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85380-108">1.0</span><span class="sxs-lookup"><span data-stu-id="85380-108">1.0</span></span>|
|[<span data-ttu-id="85380-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85380-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85380-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85380-110">ReadItem</span></span>|
|[<span data-ttu-id="85380-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85380-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="85380-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85380-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="85380-113">成员</span><span class="sxs-lookup"><span data-stu-id="85380-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="85380-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="85380-114">displayName :String</span></span>

<span data-ttu-id="85380-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="85380-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="85380-116">类型</span><span class="sxs-lookup"><span data-stu-id="85380-116">Type</span></span>

*   <span data-ttu-id="85380-117">String</span><span class="sxs-lookup"><span data-stu-id="85380-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85380-118">要求</span><span class="sxs-lookup"><span data-stu-id="85380-118">Requirements</span></span>

|<span data-ttu-id="85380-119">要求</span><span class="sxs-lookup"><span data-stu-id="85380-119">Requirement</span></span>| <span data-ttu-id="85380-120">值</span><span class="sxs-lookup"><span data-stu-id="85380-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="85380-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85380-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85380-122">1.0</span><span class="sxs-lookup"><span data-stu-id="85380-122">1.0</span></span>|
|[<span data-ttu-id="85380-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85380-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85380-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85380-124">ReadItem</span></span>|
|[<span data-ttu-id="85380-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85380-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="85380-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85380-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85380-127">示例</span><span class="sxs-lookup"><span data-stu-id="85380-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="85380-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="85380-128">emailAddress :String</span></span>

<span data-ttu-id="85380-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="85380-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="85380-130">类型</span><span class="sxs-lookup"><span data-stu-id="85380-130">Type</span></span>

*   <span data-ttu-id="85380-131">String</span><span class="sxs-lookup"><span data-stu-id="85380-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85380-132">要求</span><span class="sxs-lookup"><span data-stu-id="85380-132">Requirements</span></span>

|<span data-ttu-id="85380-133">要求</span><span class="sxs-lookup"><span data-stu-id="85380-133">Requirement</span></span>| <span data-ttu-id="85380-134">值</span><span class="sxs-lookup"><span data-stu-id="85380-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="85380-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85380-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85380-136">1.0</span><span class="sxs-lookup"><span data-stu-id="85380-136">1.0</span></span>|
|[<span data-ttu-id="85380-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85380-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85380-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85380-138">ReadItem</span></span>|
|[<span data-ttu-id="85380-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85380-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="85380-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85380-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85380-141">示例</span><span class="sxs-lookup"><span data-stu-id="85380-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="85380-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="85380-142">timeZone :String</span></span>

<span data-ttu-id="85380-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="85380-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="85380-144">类型</span><span class="sxs-lookup"><span data-stu-id="85380-144">Type</span></span>

*   <span data-ttu-id="85380-145">String</span><span class="sxs-lookup"><span data-stu-id="85380-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85380-146">要求</span><span class="sxs-lookup"><span data-stu-id="85380-146">Requirements</span></span>

|<span data-ttu-id="85380-147">要求</span><span class="sxs-lookup"><span data-stu-id="85380-147">Requirement</span></span>| <span data-ttu-id="85380-148">值</span><span class="sxs-lookup"><span data-stu-id="85380-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="85380-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85380-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85380-150">1.0</span><span class="sxs-lookup"><span data-stu-id="85380-150">1.0</span></span>|
|[<span data-ttu-id="85380-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85380-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85380-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85380-152">ReadItem</span></span>|
|[<span data-ttu-id="85380-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85380-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="85380-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85380-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85380-155">示例</span><span class="sxs-lookup"><span data-stu-id="85380-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
