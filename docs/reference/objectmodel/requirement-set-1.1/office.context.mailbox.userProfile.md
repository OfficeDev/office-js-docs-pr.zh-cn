---
title: "\"context.subname\": \"邮箱. userProfile-要求集 1.1\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451918"
---
# <a name="userprofile"></a><span data-ttu-id="44b68-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="44b68-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="44b68-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="44b68-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="44b68-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="44b68-104">Requirements</span></span>

|<span data-ttu-id="44b68-105">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-105">Requirement</span></span>| <span data-ttu-id="44b68-106">值</span><span class="sxs-lookup"><span data-stu-id="44b68-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="44b68-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44b68-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44b68-108">1.0</span><span class="sxs-lookup"><span data-stu-id="44b68-108">1.0</span></span>|
|[<span data-ttu-id="44b68-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44b68-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44b68-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44b68-110">ReadItem</span></span>|
|[<span data-ttu-id="44b68-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44b68-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44b68-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44b68-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="44b68-113">成员</span><span class="sxs-lookup"><span data-stu-id="44b68-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="44b68-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="44b68-114">displayName :String</span></span>

<span data-ttu-id="44b68-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="44b68-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="44b68-116">类型</span><span class="sxs-lookup"><span data-stu-id="44b68-116">Type</span></span>

*   <span data-ttu-id="44b68-117">String</span><span class="sxs-lookup"><span data-stu-id="44b68-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44b68-118">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-118">Requirements</span></span>

|<span data-ttu-id="44b68-119">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-119">Requirement</span></span>| <span data-ttu-id="44b68-120">值</span><span class="sxs-lookup"><span data-stu-id="44b68-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="44b68-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44b68-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44b68-122">1.0</span><span class="sxs-lookup"><span data-stu-id="44b68-122">1.0</span></span>|
|[<span data-ttu-id="44b68-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44b68-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44b68-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44b68-124">ReadItem</span></span>|
|[<span data-ttu-id="44b68-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44b68-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44b68-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44b68-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44b68-127">示例</span><span class="sxs-lookup"><span data-stu-id="44b68-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="44b68-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="44b68-128">emailAddress :String</span></span>

<span data-ttu-id="44b68-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="44b68-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="44b68-130">类型</span><span class="sxs-lookup"><span data-stu-id="44b68-130">Type</span></span>

*   <span data-ttu-id="44b68-131">String</span><span class="sxs-lookup"><span data-stu-id="44b68-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44b68-132">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-132">Requirements</span></span>

|<span data-ttu-id="44b68-133">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-133">Requirement</span></span>| <span data-ttu-id="44b68-134">值</span><span class="sxs-lookup"><span data-stu-id="44b68-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="44b68-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44b68-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44b68-136">1.0</span><span class="sxs-lookup"><span data-stu-id="44b68-136">1.0</span></span>|
|[<span data-ttu-id="44b68-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44b68-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44b68-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44b68-138">ReadItem</span></span>|
|[<span data-ttu-id="44b68-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44b68-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44b68-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44b68-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44b68-141">示例</span><span class="sxs-lookup"><span data-stu-id="44b68-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="44b68-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="44b68-142">timeZone :String</span></span>

<span data-ttu-id="44b68-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="44b68-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="44b68-144">类型</span><span class="sxs-lookup"><span data-stu-id="44b68-144">Type</span></span>

*   <span data-ttu-id="44b68-145">String</span><span class="sxs-lookup"><span data-stu-id="44b68-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44b68-146">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-146">Requirements</span></span>

|<span data-ttu-id="44b68-147">要求</span><span class="sxs-lookup"><span data-stu-id="44b68-147">Requirement</span></span>| <span data-ttu-id="44b68-148">值</span><span class="sxs-lookup"><span data-stu-id="44b68-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="44b68-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44b68-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44b68-150">1.0</span><span class="sxs-lookup"><span data-stu-id="44b68-150">1.0</span></span>|
|[<span data-ttu-id="44b68-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44b68-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44b68-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44b68-152">ReadItem</span></span>|
|[<span data-ttu-id="44b68-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44b68-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44b68-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44b68-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44b68-155">示例</span><span class="sxs-lookup"><span data-stu-id="44b68-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
