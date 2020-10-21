---
title: Office. context-预览要求集
description: 使用邮箱 API preview 要求集的 Outlook 外接程序可用的 Context 对象成员。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 8286434d2cbfc11cf0d16f8bd014b4760f0337ff
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626405"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="bcf5d-103"> (邮箱预览要求集的上下文) </span><span class="sxs-lookup"><span data-stu-id="bcf5d-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="bcf5d-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="bcf5d-104">[Office](office.md).context</span></span>

<span data-ttu-id="bcf5d-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="bcf5d-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bcf5d-107">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-107">Requirements</span></span>

|<span data-ttu-id="bcf5d-108">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-108">Requirement</span></span>| <span data-ttu-id="bcf5d-109">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-111">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-111">1.1</span></span>|
|[<span data-ttu-id="bcf5d-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="bcf5d-114">属性</span><span class="sxs-lookup"><span data-stu-id="bcf5d-114">Properties</span></span>

| <span data-ttu-id="bcf5d-115">属性</span><span class="sxs-lookup"><span data-stu-id="bcf5d-115">Property</span></span> | <span data-ttu-id="bcf5d-116">型号</span><span class="sxs-lookup"><span data-stu-id="bcf5d-116">Modes</span></span> | <span data-ttu-id="bcf5d-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-117">Return type</span></span> | <span data-ttu-id="bcf5d-118">最小值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-118">Minimum</span></span><br><span data-ttu-id="bcf5d-119">要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bcf5d-120">认证</span><span class="sxs-lookup"><span data-stu-id="bcf5d-120">auth</span></span>](#auth-auth) | <span data-ttu-id="bcf5d-121">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-121">Compose</span></span><br><span data-ttu-id="bcf5d-122">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-122">Read</span></span> | [<span data-ttu-id="bcf5d-123">Auth</span><span class="sxs-lookup"><span data-stu-id="bcf5d-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-124">IdentityAPI 1。3</span><span class="sxs-lookup"><span data-stu-id="bcf5d-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="bcf5d-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="bcf5d-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="bcf5d-126">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-126">Compose</span></span><br><span data-ttu-id="bcf5d-127">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-127">Read</span></span> | <span data-ttu-id="bcf5d-128">String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-128">String</span></span> | [<span data-ttu-id="bcf5d-129">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-130">过程</span><span class="sxs-lookup"><span data-stu-id="bcf5d-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="bcf5d-131">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-131">Compose</span></span><br><span data-ttu-id="bcf5d-132">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-132">Read</span></span> | [<span data-ttu-id="bcf5d-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="bcf5d-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="bcf5d-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="bcf5d-136">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-136">Compose</span></span><br><span data-ttu-id="bcf5d-137">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-137">Read</span></span> | <span data-ttu-id="bcf5d-138">String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-138">String</span></span> | [<span data-ttu-id="bcf5d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-140">host</span><span class="sxs-lookup"><span data-stu-id="bcf5d-140">host</span></span>](#host-hosttype) | <span data-ttu-id="bcf5d-141">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-141">Compose</span></span><br><span data-ttu-id="bcf5d-142">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-142">Read</span></span> | [<span data-ttu-id="bcf5d-143">HostType</span><span class="sxs-lookup"><span data-stu-id="bcf5d-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="bcf5d-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="bcf5d-146">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-146">Compose</span></span><br><span data-ttu-id="bcf5d-147">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-147">Read</span></span> | [<span data-ttu-id="bcf5d-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="bcf5d-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="bcf5d-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="bcf5d-151">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-151">Compose</span></span><br><span data-ttu-id="bcf5d-152">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-152">Read</span></span> | [<span data-ttu-id="bcf5d-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="bcf5d-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-154">预览</span><span class="sxs-lookup"><span data-stu-id="bcf5d-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="bcf5d-155">平台</span><span class="sxs-lookup"><span data-stu-id="bcf5d-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="bcf5d-156">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-156">Compose</span></span><br><span data-ttu-id="bcf5d-157">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-157">Read</span></span> | [<span data-ttu-id="bcf5d-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="bcf5d-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-160">满足</span><span class="sxs-lookup"><span data-stu-id="bcf5d-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="bcf5d-161">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-161">Compose</span></span><br><span data-ttu-id="bcf5d-162">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-162">Read</span></span> | [<span data-ttu-id="bcf5d-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="bcf5d-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-164">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="bcf5d-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="bcf5d-166">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-166">Compose</span></span><br><span data-ttu-id="bcf5d-167">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-167">Read</span></span> | [<span data-ttu-id="bcf5d-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="bcf5d-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-169">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bcf5d-170">ui</span><span class="sxs-lookup"><span data-stu-id="bcf5d-170">ui</span></span>](#ui-ui) | <span data-ttu-id="bcf5d-171">撰写</span><span class="sxs-lookup"><span data-stu-id="bcf5d-171">Compose</span></span><br><span data-ttu-id="bcf5d-172">读取</span><span class="sxs-lookup"><span data-stu-id="bcf5d-172">Read</span></span> | [<span data-ttu-id="bcf5d-173">UI</span><span class="sxs-lookup"><span data-stu-id="bcf5d-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bcf5d-174">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="bcf5d-175">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="bcf5d-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="bcf5d-176">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="bcf5d-177">通过提供允许 Office 应用程序获取对加载项 web 应用程序的访问令牌的方法，支持 [单一登录 (SSO) ](../../../outlook/authenticate-a-user-with-an-sso-token.md) 。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="bcf5d-178">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-179">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-179">Type</span></span>

*   [<span data-ttu-id="bcf5d-180">Auth</span><span class="sxs-lookup"><span data-stu-id="bcf5d-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-181">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-181">Requirements</span></span>

|<span data-ttu-id="bcf5d-182">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-182">Requirement</span></span>| <span data-ttu-id="bcf5d-183">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-185">预览</span><span class="sxs-lookup"><span data-stu-id="bcf5d-185">Preview</span></span>|
|[<span data-ttu-id="bcf5d-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-188">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="bcf5d-189">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-189">contentLanguage: String</span></span>

<span data-ttu-id="bcf5d-190">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="bcf5d-191">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-192">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-192">Type</span></span>

*   <span data-ttu-id="bcf5d-193">String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bcf5d-194">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-194">Requirements</span></span>

|<span data-ttu-id="bcf5d-195">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-195">Requirement</span></span>| <span data-ttu-id="bcf5d-196">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-198">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-198">1.1</span></span>|
|[<span data-ttu-id="bcf5d-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-201">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="bcf5d-202">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="bcf5d-203">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-204">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-204">Type</span></span>

*   [<span data-ttu-id="bcf5d-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="bcf5d-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-206">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-206">Requirements</span></span>

|<span data-ttu-id="bcf5d-207">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-207">Requirement</span></span>| <span data-ttu-id="bcf5d-208">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-210">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-210">1.1</span></span>|
|[<span data-ttu-id="bcf5d-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-212">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-213">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="bcf5d-214">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-214">displayLanguage: String</span></span>

<span data-ttu-id="bcf5d-215">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="bcf5d-216">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的**File > Options > 语言**指定的当前**显示语言**设置。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-217">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-217">Type</span></span>

*   <span data-ttu-id="bcf5d-218">String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bcf5d-219">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-219">Requirements</span></span>

|<span data-ttu-id="bcf5d-220">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-220">Requirement</span></span>| <span data-ttu-id="bcf5d-221">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-223">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-223">1.1</span></span>|
|[<span data-ttu-id="bcf5d-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-226">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="bcf5d-227">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="bcf5d-228">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-229">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-229">Type</span></span>

*   [<span data-ttu-id="bcf5d-230">HostType</span><span class="sxs-lookup"><span data-stu-id="bcf5d-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-231">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-231">Requirements</span></span>

|<span data-ttu-id="bcf5d-232">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-232">Requirement</span></span>| <span data-ttu-id="bcf5d-233">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-235">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-235">1.1</span></span>|
|[<span data-ttu-id="bcf5d-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-238">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="bcf5d-239">officeTheme： [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="bcf5d-240">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="bcf5d-241">此成员仅在 Windows 中的 Outlook 中受支持。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="bcf5d-242">使用 Office 主题颜色，您可以将加载项的配色方案与用户选择的当前 Office 主题进行协调，以供用户使用 **office > Office 帐户 > Office 主题 UI**，该用户在所有 Office 客户端应用程序中应用。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="bcf5d-243">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-244">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-244">Type</span></span>

*   [<span data-ttu-id="bcf5d-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="bcf5d-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="bcf5d-246">属性：</span><span class="sxs-lookup"><span data-stu-id="bcf5d-246">Properties:</span></span>

|<span data-ttu-id="bcf5d-247">名称</span><span class="sxs-lookup"><span data-stu-id="bcf5d-247">Name</span></span>| <span data-ttu-id="bcf5d-248">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-248">Type</span></span>| <span data-ttu-id="bcf5d-249">说明</span><span class="sxs-lookup"><span data-stu-id="bcf5d-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="bcf5d-250">String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-250">String</span></span>|<span data-ttu-id="bcf5d-251">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="bcf5d-252">String</span><span class="sxs-lookup"><span data-stu-id="bcf5d-252">String</span></span>|<span data-ttu-id="bcf5d-253">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="bcf5d-254">字符串</span><span class="sxs-lookup"><span data-stu-id="bcf5d-254">String</span></span>|<span data-ttu-id="bcf5d-255">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="bcf5d-256">字符串</span><span class="sxs-lookup"><span data-stu-id="bcf5d-256">String</span></span>|<span data-ttu-id="bcf5d-257">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bcf5d-258">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-258">Requirements</span></span>

|<span data-ttu-id="bcf5d-259">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-259">Requirement</span></span>| <span data-ttu-id="bcf5d-260">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-262">预览</span><span class="sxs-lookup"><span data-stu-id="bcf5d-262">Preview</span></span>|
|[<span data-ttu-id="bcf5d-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-265">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="bcf5d-266">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="bcf5d-267">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-268">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-268">Type</span></span>

*   [<span data-ttu-id="bcf5d-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="bcf5d-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-270">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-270">Requirements</span></span>

|<span data-ttu-id="bcf5d-271">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-271">Requirement</span></span>| <span data-ttu-id="bcf5d-272">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-274">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-274">1.1</span></span>|
|[<span data-ttu-id="bcf5d-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-277">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="bcf5d-278">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="bcf5d-279">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-280">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-280">Type</span></span>

*   [<span data-ttu-id="bcf5d-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="bcf5d-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-282">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-282">Requirements</span></span>

|<span data-ttu-id="bcf5d-283">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-283">Requirement</span></span>| <span data-ttu-id="bcf5d-284">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-286">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-286">1.1</span></span>|
|[<span data-ttu-id="bcf5d-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-288">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bcf5d-289">示例</span><span class="sxs-lookup"><span data-stu-id="bcf5d-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="bcf5d-290">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="bcf5d-291">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="bcf5d-292">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-293">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-293">Type</span></span>

*   [<span data-ttu-id="bcf5d-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="bcf5d-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-295">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-295">Requirements</span></span>

|<span data-ttu-id="bcf5d-296">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-296">Requirement</span></span>| <span data-ttu-id="bcf5d-297">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-299">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-299">1.1</span></span>|
|[<span data-ttu-id="bcf5d-300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bcf5d-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="bcf5d-301">受限</span><span class="sxs-lookup"><span data-stu-id="bcf5d-301">Restricted</span></span>|
|[<span data-ttu-id="bcf5d-302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-303">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="bcf5d-304">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="bcf5d-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="bcf5d-305">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="bcf5d-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="bcf5d-306">类型</span><span class="sxs-lookup"><span data-stu-id="bcf5d-306">Type</span></span>

*   [<span data-ttu-id="bcf5d-307">UI</span><span class="sxs-lookup"><span data-stu-id="bcf5d-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="bcf5d-308">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-308">Requirements</span></span>

|<span data-ttu-id="bcf5d-309">要求</span><span class="sxs-lookup"><span data-stu-id="bcf5d-309">Requirement</span></span>| <span data-ttu-id="bcf5d-310">值</span><span class="sxs-lookup"><span data-stu-id="bcf5d-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcf5d-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bcf5d-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bcf5d-312">1.1</span><span class="sxs-lookup"><span data-stu-id="bcf5d-312">1.1</span></span>|
|[<span data-ttu-id="bcf5d-313">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bcf5d-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bcf5d-314">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bcf5d-314">Compose or Read</span></span>|
