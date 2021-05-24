---
title: Office.context - 要求集 1.9
description: Office。使用邮箱 API 要求集 1.9 Outlook外接程序可用的上下文对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: f45eec7ce638f4bbb97ad4be9f2ba089905c631d
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590517"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="2f92d-103">context (Mailbox requirement set 1.9) </span><span class="sxs-lookup"><span data-stu-id="2f92d-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="2f92d-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="2f92d-104">[Office](office.md).context</span></span>

<span data-ttu-id="2f92d-105">Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。</span><span class="sxs-lookup"><span data-stu-id="2f92d-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="2f92d-106">此列表仅记录外接程序Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="2f92d-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f92d-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-107">Requirements</span></span>

|<span data-ttu-id="2f92d-108">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-108">Requirement</span></span>| <span data-ttu-id="2f92d-109">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-111">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-111">1.1</span></span>|
|[<span data-ttu-id="2f92d-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="2f92d-114">属性</span><span class="sxs-lookup"><span data-stu-id="2f92d-114">Properties</span></span>

| <span data-ttu-id="2f92d-115">属性</span><span class="sxs-lookup"><span data-stu-id="2f92d-115">Property</span></span> | <span data-ttu-id="2f92d-116">模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-116">Modes</span></span> | <span data-ttu-id="2f92d-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-117">Return type</span></span> | <span data-ttu-id="2f92d-118">最小值</span><span class="sxs-lookup"><span data-stu-id="2f92d-118">Minimum</span></span><br><span data-ttu-id="2f92d-119">要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2f92d-120">auth</span><span class="sxs-lookup"><span data-stu-id="2f92d-120">auth</span></span>](#auth-auth) | <span data-ttu-id="2f92d-121">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-121">Compose</span></span><br><span data-ttu-id="2f92d-122">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-122">Read</span></span> | [<span data-ttu-id="2f92d-123">Auth</span><span class="sxs-lookup"><span data-stu-id="2f92d-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="2f92d-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="2f92d-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="2f92d-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="2f92d-126">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-126">Compose</span></span><br><span data-ttu-id="2f92d-127">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-127">Read</span></span> | <span data-ttu-id="2f92d-128">字符串</span><span class="sxs-lookup"><span data-stu-id="2f92d-128">String</span></span> | [<span data-ttu-id="2f92d-129">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f92d-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="2f92d-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="2f92d-131">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-131">Compose</span></span><br><span data-ttu-id="2f92d-132">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-132">Read</span></span> | [<span data-ttu-id="2f92d-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="2f92d-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f92d-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="2f92d-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="2f92d-136">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-136">Compose</span></span><br><span data-ttu-id="2f92d-137">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-137">Read</span></span> | <span data-ttu-id="2f92d-138">字符串</span><span class="sxs-lookup"><span data-stu-id="2f92d-138">String</span></span> | [<span data-ttu-id="2f92d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f92d-140">host</span><span class="sxs-lookup"><span data-stu-id="2f92d-140">host</span></span>](#host-hosttype) | <span data-ttu-id="2f92d-141">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-141">Compose</span></span><br><span data-ttu-id="2f92d-142">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-142">Read</span></span> | [<span data-ttu-id="2f92d-143">HostType</span><span class="sxs-lookup"><span data-stu-id="2f92d-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-144">1.5</span><span class="sxs-lookup"><span data-stu-id="2f92d-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2f92d-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="2f92d-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="2f92d-146">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-146">Compose</span></span><br><span data-ttu-id="2f92d-147">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-147">Read</span></span> | [<span data-ttu-id="2f92d-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="2f92d-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f92d-150">平台</span><span class="sxs-lookup"><span data-stu-id="2f92d-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="2f92d-151">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-151">Compose</span></span><br><span data-ttu-id="2f92d-152">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-152">Read</span></span> | [<span data-ttu-id="2f92d-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="2f92d-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-154">1.5</span><span class="sxs-lookup"><span data-stu-id="2f92d-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2f92d-155">requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="2f92d-156">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-156">Compose</span></span><br><span data-ttu-id="2f92d-157">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-157">Read</span></span> | [<span data-ttu-id="2f92d-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="2f92d-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f92d-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="2f92d-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="2f92d-161">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-161">Compose</span></span><br><span data-ttu-id="2f92d-162">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-162">Read</span></span> | [<span data-ttu-id="2f92d-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2f92d-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-164">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f92d-165">ui</span><span class="sxs-lookup"><span data-stu-id="2f92d-165">ui</span></span>](#ui-ui) | <span data-ttu-id="2f92d-166">撰写</span><span class="sxs-lookup"><span data-stu-id="2f92d-166">Compose</span></span><br><span data-ttu-id="2f92d-167">阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-167">Read</span></span> | [<span data-ttu-id="2f92d-168">UI</span><span class="sxs-lookup"><span data-stu-id="2f92d-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="2f92d-169">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="2f92d-170">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="2f92d-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="2f92d-171">身份验证 [：Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="2f92d-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="2f92d-172">通过提供允许 Office 应用程序获取对外接程序 Web 应用程序的访问令牌的方法 ([SSO](../../../outlook/authenticate-a-user-with-an-sso-token.md)) 支持单一登录。</span><span class="sxs-lookup"><span data-stu-id="2f92d-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="2f92d-173">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="2f92d-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="2f92d-174">请参阅 [IdentityAPI 1.3 要求集](../../requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="2f92d-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-175">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-175">Type</span></span>

*   [<span data-ttu-id="2f92d-176">Auth</span><span class="sxs-lookup"><span data-stu-id="2f92d-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="2f92d-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-177">Requirements</span></span>

|<span data-ttu-id="2f92d-178">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-178">Requirement</span></span>| <span data-ttu-id="2f92d-179">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-180">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-181">无</span><span class="sxs-lookup"><span data-stu-id="2f92d-181">N/A</span></span>|
|[<span data-ttu-id="2f92d-182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-183">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-184">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-184">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="2f92d-185">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="2f92d-185">contentLanguage: String</span></span>

<span data-ttu-id="2f92d-186">获取用户 (编辑) 的语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="2f92d-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="2f92d-187">该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office设置。</span><span class="sxs-lookup"><span data-stu-id="2f92d-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-188">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-188">Type</span></span>

*   <span data-ttu-id="2f92d-189">String</span><span class="sxs-lookup"><span data-stu-id="2f92d-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f92d-190">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-190">Requirements</span></span>

|<span data-ttu-id="2f92d-191">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-191">Requirement</span></span>| <span data-ttu-id="2f92d-192">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-193">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-194">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-194">1.1</span></span>|
|[<span data-ttu-id="2f92d-195">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-196">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-197">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-197">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="2f92d-198">diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="2f92d-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="2f92d-199">获取加载项运行环境的信息。</span><span class="sxs-lookup"><span data-stu-id="2f92d-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-200">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-200">Type</span></span>

*   [<span data-ttu-id="2f92d-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="2f92d-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="2f92d-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-202">Requirements</span></span>

|<span data-ttu-id="2f92d-203">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-203">Requirement</span></span>| <span data-ttu-id="2f92d-204">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-206">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-206">1.1</span></span>|
|[<span data-ttu-id="2f92d-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-209">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="2f92d-210">displayLanguage：String</span><span class="sxs-lookup"><span data-stu-id="2f92d-210">displayLanguage: String</span></span>

<span data-ttu-id="2f92d-211">获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。</span><span class="sxs-lookup"><span data-stu-id="2f92d-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="2f92d-212">该值反映当前显示语言设置，该设置由 > `displayLanguage` **客户端** 应用程序中>选项Office语言。 </span><span class="sxs-lookup"><span data-stu-id="2f92d-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-213">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-213">Type</span></span>

*   <span data-ttu-id="2f92d-214">String</span><span class="sxs-lookup"><span data-stu-id="2f92d-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f92d-215">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-215">Requirements</span></span>

|<span data-ttu-id="2f92d-216">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-216">Requirement</span></span>| <span data-ttu-id="2f92d-217">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-219">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-219">1.1</span></span>|
|[<span data-ttu-id="2f92d-220">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-221">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-222">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-222">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="2f92d-223">host： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="2f92d-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="2f92d-224">获取Office加载项的加载项应用程序。</span><span class="sxs-lookup"><span data-stu-id="2f92d-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2f92d-225">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="2f92d-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-226">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-226">Type</span></span>

*   [<span data-ttu-id="2f92d-227">HostType</span><span class="sxs-lookup"><span data-stu-id="2f92d-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="2f92d-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-228">Requirements</span></span>

|<span data-ttu-id="2f92d-229">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-229">Requirement</span></span>| <span data-ttu-id="2f92d-230">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-231">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-232">1.5</span><span class="sxs-lookup"><span data-stu-id="2f92d-232">1.5</span></span>|
|[<span data-ttu-id="2f92d-233">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-234">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-235">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="2f92d-236">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="2f92d-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="2f92d-237">提供运行加载项的平台。</span><span class="sxs-lookup"><span data-stu-id="2f92d-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="2f92d-238">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="2f92d-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-239">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-239">Type</span></span>

*   [<span data-ttu-id="2f92d-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="2f92d-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="2f92d-241">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-241">Requirements</span></span>

|<span data-ttu-id="2f92d-242">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-242">Requirement</span></span>| <span data-ttu-id="2f92d-243">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-244">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-245">1.5</span><span class="sxs-lookup"><span data-stu-id="2f92d-245">1.5</span></span>|
|[<span data-ttu-id="2f92d-246">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-247">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-248">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="2f92d-249">requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="2f92d-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="2f92d-250">提供用于确定当前应用程序和平台上支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="2f92d-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-251">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-251">Type</span></span>

*   [<span data-ttu-id="2f92d-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="2f92d-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="2f92d-253">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-253">Requirements</span></span>

|<span data-ttu-id="2f92d-254">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-254">Requirement</span></span>| <span data-ttu-id="2f92d-255">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-256">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-257">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-257">1.1</span></span>|
|[<span data-ttu-id="2f92d-258">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-259">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2f92d-260">示例</span><span class="sxs-lookup"><span data-stu-id="2f92d-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="2f92d-261">[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="2f92d-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="2f92d-262">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="2f92d-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="2f92d-263">该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可供该外接程序使用 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="2f92d-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-264">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-264">Type</span></span>

*   [<span data-ttu-id="2f92d-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2f92d-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="2f92d-266">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-266">Requirements</span></span>

|<span data-ttu-id="2f92d-267">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-267">Requirement</span></span>| <span data-ttu-id="2f92d-268">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-269">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-270">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-270">1.1</span></span>|
|[<span data-ttu-id="2f92d-271">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2f92d-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="2f92d-272">受限</span><span class="sxs-lookup"><span data-stu-id="2f92d-272">Restricted</span></span>|
|[<span data-ttu-id="2f92d-273">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-274">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="2f92d-275">[ui：UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="2f92d-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="2f92d-276">提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。</span><span class="sxs-lookup"><span data-stu-id="2f92d-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="2f92d-277">类型</span><span class="sxs-lookup"><span data-stu-id="2f92d-277">Type</span></span>

*   [<span data-ttu-id="2f92d-278">UI</span><span class="sxs-lookup"><span data-stu-id="2f92d-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="2f92d-279">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f92d-279">Requirements</span></span>

|<span data-ttu-id="2f92d-280">要求</span><span class="sxs-lookup"><span data-stu-id="2f92d-280">Requirement</span></span>| <span data-ttu-id="2f92d-281">值</span><span class="sxs-lookup"><span data-stu-id="2f92d-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f92d-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2f92d-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f92d-283">1.1</span><span class="sxs-lookup"><span data-stu-id="2f92d-283">1.1</span></span>|
|[<span data-ttu-id="2f92d-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2f92d-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f92d-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2f92d-285">Compose or Read</span></span>|
