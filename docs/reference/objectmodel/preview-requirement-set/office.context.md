---
title: Office.context - 预览要求集
description: Office。适用于使用邮箱 API Outlook要求集的外接程序的上下文对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 59b1cce579afe69384e41a6f31cc70c8cec25bea
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591070"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="1dd6a-103">上下文 (邮箱预览要求集) </span><span class="sxs-lookup"><span data-stu-id="1dd6a-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="1dd6a-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1dd6a-104">[Office](office.md).context</span></span>

<span data-ttu-id="1dd6a-105">Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1dd6a-106">此列表仅记录外接程序Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd6a-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-107">Requirements</span></span>

|<span data-ttu-id="1dd6a-108">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-108">Requirement</span></span>| <span data-ttu-id="1dd6a-109">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-111">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-111">1.1</span></span>|
|[<span data-ttu-id="1dd6a-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="1dd6a-114">属性</span><span class="sxs-lookup"><span data-stu-id="1dd6a-114">Properties</span></span>

| <span data-ttu-id="1dd6a-115">属性</span><span class="sxs-lookup"><span data-stu-id="1dd6a-115">Property</span></span> | <span data-ttu-id="1dd6a-116">模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-116">Modes</span></span> | <span data-ttu-id="1dd6a-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-117">Return type</span></span> | <span data-ttu-id="1dd6a-118">最小值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-118">Minimum</span></span><br><span data-ttu-id="1dd6a-119">要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1dd6a-120">auth</span><span class="sxs-lookup"><span data-stu-id="1dd6a-120">auth</span></span>](#auth-auth) | <span data-ttu-id="1dd6a-121">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-121">Compose</span></span><br><span data-ttu-id="1dd6a-122">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-122">Read</span></span> | [<span data-ttu-id="1dd6a-123">Auth</span><span class="sxs-lookup"><span data-stu-id="1dd6a-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="1dd6a-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="1dd6a-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1dd6a-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1dd6a-126">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-126">Compose</span></span><br><span data-ttu-id="1dd6a-127">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-127">Read</span></span> | <span data-ttu-id="1dd6a-128">字符串</span><span class="sxs-lookup"><span data-stu-id="1dd6a-128">String</span></span> | [<span data-ttu-id="1dd6a-129">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1dd6a-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="1dd6a-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1dd6a-131">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-131">Compose</span></span><br><span data-ttu-id="1dd6a-132">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-132">Read</span></span> | [<span data-ttu-id="1dd6a-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1dd6a-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1dd6a-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1dd6a-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1dd6a-136">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-136">Compose</span></span><br><span data-ttu-id="1dd6a-137">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-137">Read</span></span> | <span data-ttu-id="1dd6a-138">字符串</span><span class="sxs-lookup"><span data-stu-id="1dd6a-138">String</span></span> | [<span data-ttu-id="1dd6a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1dd6a-140">host</span><span class="sxs-lookup"><span data-stu-id="1dd6a-140">host</span></span>](#host-hosttype) | <span data-ttu-id="1dd6a-141">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-141">Compose</span></span><br><span data-ttu-id="1dd6a-142">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-142">Read</span></span> | [<span data-ttu-id="1dd6a-143">HostType</span><span class="sxs-lookup"><span data-stu-id="1dd6a-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-144">1.5</span><span class="sxs-lookup"><span data-stu-id="1dd6a-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1dd6a-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="1dd6a-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1dd6a-146">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-146">Compose</span></span><br><span data-ttu-id="1dd6a-147">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-147">Read</span></span> | [<span data-ttu-id="1dd6a-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="1dd6a-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1dd6a-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="1dd6a-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="1dd6a-151">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-151">Compose</span></span><br><span data-ttu-id="1dd6a-152">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-152">Read</span></span> | [<span data-ttu-id="1dd6a-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="1dd6a-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-154">预览</span><span class="sxs-lookup"><span data-stu-id="1dd6a-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="1dd6a-155">平台</span><span class="sxs-lookup"><span data-stu-id="1dd6a-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="1dd6a-156">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-156">Compose</span></span><br><span data-ttu-id="1dd6a-157">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-157">Read</span></span> | [<span data-ttu-id="1dd6a-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1dd6a-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-159">1.5</span><span class="sxs-lookup"><span data-stu-id="1dd6a-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1dd6a-160">requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1dd6a-161">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-161">Compose</span></span><br><span data-ttu-id="1dd6a-162">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-162">Read</span></span> | [<span data-ttu-id="1dd6a-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1dd6a-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-164">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1dd6a-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1dd6a-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1dd6a-166">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-166">Compose</span></span><br><span data-ttu-id="1dd6a-167">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-167">Read</span></span> | [<span data-ttu-id="1dd6a-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1dd6a-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-169">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1dd6a-170">ui</span><span class="sxs-lookup"><span data-stu-id="1dd6a-170">ui</span></span>](#ui-ui) | <span data-ttu-id="1dd6a-171">撰写</span><span class="sxs-lookup"><span data-stu-id="1dd6a-171">Compose</span></span><br><span data-ttu-id="1dd6a-172">阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-172">Read</span></span> | [<span data-ttu-id="1dd6a-173">UI</span><span class="sxs-lookup"><span data-stu-id="1dd6a-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="1dd6a-174">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1dd6a-175">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="1dd6a-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="1dd6a-176">身份验证 [：Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="1dd6a-177">通过提供允许 Office 应用程序获取对外接程序 Web 应用程序的访问令牌的方法 ([SSO](../../../outlook/authenticate-a-user-with-an-sso-token.md)) 支持单一登录。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="1dd6a-178">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-179">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-179">Type</span></span>

*   [<span data-ttu-id="1dd6a-180">Auth</span><span class="sxs-lookup"><span data-stu-id="1dd6a-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-181">Requirements</span></span>

|<span data-ttu-id="1dd6a-182">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-182">Requirement</span></span>| <span data-ttu-id="1dd6a-183">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-185">预览</span><span class="sxs-lookup"><span data-stu-id="1dd6a-185">Preview</span></span>|
|[<span data-ttu-id="1dd6a-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-188">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="1dd6a-189">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="1dd6a-189">contentLanguage: String</span></span>

<span data-ttu-id="1dd6a-190">获取用户 (编辑) 的语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1dd6a-191">该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office设置。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-192">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-192">Type</span></span>

*   <span data-ttu-id="1dd6a-193">String</span><span class="sxs-lookup"><span data-stu-id="1dd6a-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd6a-194">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-194">Requirements</span></span>

|<span data-ttu-id="1dd6a-195">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-195">Requirement</span></span>| <span data-ttu-id="1dd6a-196">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-198">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-198">1.1</span></span>|
|[<span data-ttu-id="1dd6a-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-201">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1dd6a-202">diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1dd6a-203">获取加载项运行环境的信息。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-204">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-204">Type</span></span>

*   [<span data-ttu-id="1dd6a-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1dd6a-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-206">Requirements</span></span>

|<span data-ttu-id="1dd6a-207">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-207">Requirement</span></span>| <span data-ttu-id="1dd6a-208">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-210">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-210">1.1</span></span>|
|[<span data-ttu-id="1dd6a-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-212">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-213">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1dd6a-214">displayLanguage：String</span><span class="sxs-lookup"><span data-stu-id="1dd6a-214">displayLanguage: String</span></span>

<span data-ttu-id="1dd6a-215">获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="1dd6a-216">该值反映当前显示语言设置，该设置由 > `displayLanguage` **客户端** 应用程序中>选项Office语言。 </span><span class="sxs-lookup"><span data-stu-id="1dd6a-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-217">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-217">Type</span></span>

*   <span data-ttu-id="1dd6a-218">String</span><span class="sxs-lookup"><span data-stu-id="1dd6a-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dd6a-219">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-219">Requirements</span></span>

|<span data-ttu-id="1dd6a-220">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-220">Requirement</span></span>| <span data-ttu-id="1dd6a-221">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-223">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-223">1.1</span></span>|
|[<span data-ttu-id="1dd6a-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-226">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="1dd6a-227">host： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="1dd6a-228">获取Office加载项的加载项应用程序。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1dd6a-229">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取主机。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-230">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-230">Type</span></span>

*   [<span data-ttu-id="1dd6a-231">HostType</span><span class="sxs-lookup"><span data-stu-id="1dd6a-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-232">Requirements</span></span>

|<span data-ttu-id="1dd6a-233">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-233">Requirement</span></span>| <span data-ttu-id="1dd6a-234">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-236">1.5</span><span class="sxs-lookup"><span data-stu-id="1dd6a-236">1.5</span></span>|
|[<span data-ttu-id="1dd6a-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-239">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="1dd6a-240">[officeTheme：OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="1dd6a-241">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="1dd6a-242">此成员仅在 Outlook 支持Windows。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="1dd6a-243">使用 Office 主题颜色，可以将外接程序的配色方案与用户通过文件 > Office **帐户 > Office** 主题 UI 选择的当前 Office 主题协调，该 UI 适用于所有 Office 客户端应用程序。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="1dd6a-244">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-245">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-245">Type</span></span>

*   [<span data-ttu-id="1dd6a-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="1dd6a-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="1dd6a-247">属性</span><span class="sxs-lookup"><span data-stu-id="1dd6a-247">Properties</span></span>

|<span data-ttu-id="1dd6a-248">名称</span><span class="sxs-lookup"><span data-stu-id="1dd6a-248">Name</span></span>| <span data-ttu-id="1dd6a-249">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-249">Type</span></span>| <span data-ttu-id="1dd6a-250">描述</span><span class="sxs-lookup"><span data-stu-id="1dd6a-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="1dd6a-251">字符串</span><span class="sxs-lookup"><span data-stu-id="1dd6a-251">String</span></span>|<span data-ttu-id="1dd6a-252">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="1dd6a-253">String</span><span class="sxs-lookup"><span data-stu-id="1dd6a-253">String</span></span>|<span data-ttu-id="1dd6a-254">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="1dd6a-255">字符串</span><span class="sxs-lookup"><span data-stu-id="1dd6a-255">String</span></span>|<span data-ttu-id="1dd6a-256">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="1dd6a-257">字符串</span><span class="sxs-lookup"><span data-stu-id="1dd6a-257">String</span></span>|<span data-ttu-id="1dd6a-258">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dd6a-259">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-259">Requirements</span></span>

|<span data-ttu-id="1dd6a-260">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-260">Requirement</span></span>| <span data-ttu-id="1dd6a-261">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-262">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-263">预览</span><span class="sxs-lookup"><span data-stu-id="1dd6a-263">Preview</span></span>|
|[<span data-ttu-id="1dd6a-264">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-265">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-266">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-266">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="1dd6a-267">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="1dd6a-268">提供运行加载项的平台。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="1dd6a-269">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-270">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-270">Type</span></span>

*   [<span data-ttu-id="1dd6a-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1dd6a-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-272">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-272">Requirements</span></span>

|<span data-ttu-id="1dd6a-273">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-273">Requirement</span></span>| <span data-ttu-id="1dd6a-274">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-275">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-276">1.5</span><span class="sxs-lookup"><span data-stu-id="1dd6a-276">1.5</span></span>|
|[<span data-ttu-id="1dd6a-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-278">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-279">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1dd6a-280">requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1dd6a-281">提供用于确定当前应用程序和平台上支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-282">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-282">Type</span></span>

*   [<span data-ttu-id="1dd6a-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1dd6a-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-284">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-284">Requirements</span></span>

|<span data-ttu-id="1dd6a-285">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-285">Requirement</span></span>| <span data-ttu-id="1dd6a-286">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-287">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-288">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-288">1.1</span></span>|
|[<span data-ttu-id="1dd6a-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dd6a-291">示例</span><span class="sxs-lookup"><span data-stu-id="1dd6a-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1dd6a-292">[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1dd6a-293">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1dd6a-294">该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可供该外接程序使用 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-295">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-295">Type</span></span>

*   [<span data-ttu-id="1dd6a-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1dd6a-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-297">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-297">Requirements</span></span>

|<span data-ttu-id="1dd6a-298">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-298">Requirement</span></span>| <span data-ttu-id="1dd6a-299">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-300">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-301">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-301">1.1</span></span>|
|[<span data-ttu-id="1dd6a-302">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1dd6a-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1dd6a-303">受限</span><span class="sxs-lookup"><span data-stu-id="1dd6a-303">Restricted</span></span>|
|[<span data-ttu-id="1dd6a-304">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-305">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1dd6a-306">[ui：UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1dd6a-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1dd6a-307">提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。</span><span class="sxs-lookup"><span data-stu-id="1dd6a-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1dd6a-308">类型</span><span class="sxs-lookup"><span data-stu-id="1dd6a-308">Type</span></span>

*   [<span data-ttu-id="1dd6a-309">UI</span><span class="sxs-lookup"><span data-stu-id="1dd6a-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1dd6a-310">Requirements</span><span class="sxs-lookup"><span data-stu-id="1dd6a-310">Requirements</span></span>

|<span data-ttu-id="1dd6a-311">要求</span><span class="sxs-lookup"><span data-stu-id="1dd6a-311">Requirement</span></span>| <span data-ttu-id="1dd6a-312">值</span><span class="sxs-lookup"><span data-stu-id="1dd6a-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dd6a-313">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1dd6a-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1dd6a-314">1.1</span><span class="sxs-lookup"><span data-stu-id="1dd6a-314">1.1</span></span>|
|[<span data-ttu-id="1dd6a-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1dd6a-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1dd6a-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1dd6a-316">Compose or Read</span></span>|
