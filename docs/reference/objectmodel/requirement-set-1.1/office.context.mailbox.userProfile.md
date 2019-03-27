---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.1\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870189"
---
# <a name="userprofile"></a><span data-ttu-id="b9ff9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="b9ff9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="b9ff9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="b9ff9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9ff9-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="b9ff9-104">Requirements</span></span>

|<span data-ttu-id="b9ff9-105">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-105">Requirement</span></span>| <span data-ttu-id="b9ff9-106">值</span><span class="sxs-lookup"><span data-stu-id="b9ff9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9ff9-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b9ff9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9ff9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b9ff9-108">1.0</span></span>|
|[<span data-ttu-id="b9ff9-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b9ff9-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9ff9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9ff9-110">ReadItem</span></span>|
|[<span data-ttu-id="b9ff9-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b9ff9-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9ff9-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b9ff9-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="b9ff9-113">成员</span><span class="sxs-lookup"><span data-stu-id="b9ff9-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b9ff9-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b9ff9-114">displayName :String</span></span>

<span data-ttu-id="b9ff9-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="b9ff9-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b9ff9-116">类型</span><span class="sxs-lookup"><span data-stu-id="b9ff9-116">Type</span></span>

*   <span data-ttu-id="b9ff9-117">String</span><span class="sxs-lookup"><span data-stu-id="b9ff9-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9ff9-118">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-118">Requirements</span></span>

|<span data-ttu-id="b9ff9-119">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-119">Requirement</span></span>| <span data-ttu-id="b9ff9-120">值</span><span class="sxs-lookup"><span data-stu-id="b9ff9-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9ff9-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b9ff9-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9ff9-122">1.0</span><span class="sxs-lookup"><span data-stu-id="b9ff9-122">1.0</span></span>|
|[<span data-ttu-id="b9ff9-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b9ff9-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9ff9-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9ff9-124">ReadItem</span></span>|
|[<span data-ttu-id="b9ff9-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b9ff9-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9ff9-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b9ff9-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9ff9-127">示例</span><span class="sxs-lookup"><span data-stu-id="b9ff9-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b9ff9-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b9ff9-128">emailAddress :String</span></span>

<span data-ttu-id="b9ff9-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="b9ff9-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b9ff9-130">类型</span><span class="sxs-lookup"><span data-stu-id="b9ff9-130">Type</span></span>

*   <span data-ttu-id="b9ff9-131">String</span><span class="sxs-lookup"><span data-stu-id="b9ff9-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9ff9-132">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-132">Requirements</span></span>

|<span data-ttu-id="b9ff9-133">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-133">Requirement</span></span>| <span data-ttu-id="b9ff9-134">值</span><span class="sxs-lookup"><span data-stu-id="b9ff9-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9ff9-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b9ff9-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9ff9-136">1.0</span><span class="sxs-lookup"><span data-stu-id="b9ff9-136">1.0</span></span>|
|[<span data-ttu-id="b9ff9-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b9ff9-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9ff9-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9ff9-138">ReadItem</span></span>|
|[<span data-ttu-id="b9ff9-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b9ff9-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9ff9-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b9ff9-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9ff9-141">示例</span><span class="sxs-lookup"><span data-stu-id="b9ff9-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b9ff9-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b9ff9-142">timeZone :String</span></span>

<span data-ttu-id="b9ff9-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="b9ff9-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b9ff9-144">类型</span><span class="sxs-lookup"><span data-stu-id="b9ff9-144">Type</span></span>

*   <span data-ttu-id="b9ff9-145">String</span><span class="sxs-lookup"><span data-stu-id="b9ff9-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9ff9-146">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-146">Requirements</span></span>

|<span data-ttu-id="b9ff9-147">要求</span><span class="sxs-lookup"><span data-stu-id="b9ff9-147">Requirement</span></span>| <span data-ttu-id="b9ff9-148">值</span><span class="sxs-lookup"><span data-stu-id="b9ff9-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9ff9-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b9ff9-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9ff9-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b9ff9-150">1.0</span></span>|
|[<span data-ttu-id="b9ff9-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b9ff9-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9ff9-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9ff9-152">ReadItem</span></span>|
|[<span data-ttu-id="b9ff9-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b9ff9-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9ff9-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b9ff9-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9ff9-155">示例</span><span class="sxs-lookup"><span data-stu-id="b9ff9-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
