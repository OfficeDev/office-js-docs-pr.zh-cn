---
title: Office 外接程序的身份验证设计准则
description: 了解如何在 Office 外接程序中直观地设计登录页或注册页。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cbd90bc6eba277b0fb313df6ce442aa73e8a997d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718480"
---
# <a name="authentication-patterns"></a><span data-ttu-id="43ea8-103">身份验证模式</span><span class="sxs-lookup"><span data-stu-id="43ea8-103">Authentication patterns</span></span>

<span data-ttu-id="43ea8-104">加载项可能需要用户登录或注册才能访问特性和功能。</span><span class="sxs-lookup"><span data-stu-id="43ea8-104">Add-ins may require users to sign-in or sign-up in order to access features and functionality.</span></span> <span data-ttu-id="43ea8-105">用户名和密码的输入框或启动第三方凭据流的按钮是身份验证体验中常见的界面控件。</span><span class="sxs-lookup"><span data-stu-id="43ea8-105">Input boxes for username and password or buttons that start third party credential flows are common interface controls in authentication experiences.</span></span> <span data-ttu-id="43ea8-106">获得简单高效的身份验证体验是让用户开始使用加载项的重要第一步。</span><span class="sxs-lookup"><span data-stu-id="43ea8-106">A simple and efficient authentication experience is an important first step to getting users started with your add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="43ea8-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="43ea8-107">Best practices</span></span>

|<span data-ttu-id="43ea8-108">允许事项</span><span class="sxs-lookup"><span data-stu-id="43ea8-108">Do</span></span>|<span data-ttu-id="43ea8-109">禁止事项</span><span class="sxs-lookup"><span data-stu-id="43ea8-109">Don't</span></span>|
|:----|:----|
|<span data-ttu-id="43ea8-110">在登录之前，无需帐户即可介绍加载项或演示功能的价值。</span><span class="sxs-lookup"><span data-stu-id="43ea8-110">Prior to sign-in, describe the value of your add-in or demonstrate functionality without requiring an account.</span></span> |<span data-ttu-id="43ea8-111">希望用户无需了解加载项的价值和好处即可登录。</span><span class="sxs-lookup"><span data-stu-id="43ea8-111">Expect users to sign-in without understanding the value and benefits of your add-in.</span></span>|
|<span data-ttu-id="43ea8-112">指导用户通过身份验证流程，并在每个屏幕上使用主要的高度可见的按钮。</span><span class="sxs-lookup"><span data-stu-id="43ea8-112">Guide users through authentication flows with a primary, highly visible button on each screen.</span></span> |<span data-ttu-id="43ea8-113">通过竞争按钮和号召性用语，引起对二级和三级任务的关注。</span><span class="sxs-lookup"><span data-stu-id="43ea8-113">Draw attention to secondary and tertiary tasks with competing buttons and calls to action.</span></span>|
|<span data-ttu-id="43ea8-114">使用清晰的按钮标签来描述“登录”或“创建帐户”等特定任务。</span><span class="sxs-lookup"><span data-stu-id="43ea8-114">Use clear button labels that describe specific tasks like "Sign in" or "Create account".</span></span>   |<span data-ttu-id="43ea8-115">使用模糊的按钮标签，如“提交”或“入门”来指导用户完成身份验证流程。</span><span class="sxs-lookup"><span data-stu-id="43ea8-115">Use vague button labels like "Submit" or "Get started" to guide users through authentication flows.</span></span>|
|<span data-ttu-id="43ea8-116">使用对话框将用户的注意力集中在身份验证表单上。</span><span class="sxs-lookup"><span data-stu-id="43ea8-116">Use a dialog to focus users' attention on authentication forms.</span></span>    |<span data-ttu-id="43ea8-117">使用初次运行体验和身份验证表单塞满任务窗格。</span><span class="sxs-lookup"><span data-stu-id="43ea8-117">Overcrowd your task pane with a first run experience and authentication forms.</span></span>|
|<span data-ttu-id="43ea8-118">在流程中寻找细处的效率，如自动对焦输入框。</span><span class="sxs-lookup"><span data-stu-id="43ea8-118">Find small efficiencies in the flow like auto-focusing on input boxes.</span></span> |<span data-ttu-id="43ea8-119">为交互添加如要求用户单击表单域等不必要的步骤。</span><span class="sxs-lookup"><span data-stu-id="43ea8-119">Add unnecessary steps to the interaction like requiring users to click into form fields.</span></span>|
|<span data-ttu-id="43ea8-120">为用户提供注销和重新进行身份验证的方法。</span><span class="sxs-lookup"><span data-stu-id="43ea8-120">Provide a way for users to sign-out and reauthenticate.</span></span>    |<span data-ttu-id="43ea8-121">强制用户卸载以切换标识。</span><span class="sxs-lookup"><span data-stu-id="43ea8-121">Force users to uninstall to switch identities.</span></span>|

## <a name="authentication-flow"></a><span data-ttu-id="43ea8-122">身份验证流</span><span class="sxs-lookup"><span data-stu-id="43ea8-122">Authentication flow</span></span>

<span data-ttu-id="43ea8-123">在单一登录处于预览期间，生产加载项应允许用户可以选择直接使用服务或 Microsoft 等标识提供者进行登录。</span><span class="sxs-lookup"><span data-stu-id="43ea8-123">Until single sign-on is out of preview, production add-ins should give users a choice to sign-in directly with your service or an identity provider like Microsoft.</span></span>

1. <span data-ttu-id="43ea8-124">首先运行 Placemat - 将登录按钮设置为加载项初次运行体验中的明确号召性用语。</span><span class="sxs-lookup"><span data-stu-id="43ea8-124">First Run Placemat - Place your sign-in button as a clear call-to action inside your add-in's first run experience.</span></span>
<span data-ttu-id="43ea8-125">![Office 应用程序中的加载项任务窗格屏幕截图](../images/add-in-fre-value-placemat.png)</span><span class="sxs-lookup"><span data-stu-id="43ea8-125">![A screenshot of an add-in task pane in an Office application](../images/add-in-fre-value-placemat.png)</span></span>

2. <span data-ttu-id="43ea8-126">标识提供者选项对话框 - 显示明确的标识提供者列表，包括用户名和密码表单（如适用）。</span><span class="sxs-lookup"><span data-stu-id="43ea8-126">Identity Provider Choices Dialog - Display a clear list of identity providers including a username and password form if applicable.</span></span> <span data-ttu-id="43ea8-127">身份验证对话框处于打开状态时，加载项 UI 可能会被阻止。</span><span class="sxs-lookup"><span data-stu-id="43ea8-127">Your add-in UI may be blocked while the authentication dialog is open.</span></span>
<span data-ttu-id="43ea8-128">![Office 应用程序中的身份提供程序选项对话框的屏幕截图](../images/add-in-auth-choices-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="43ea8-128">![A screenshot of the Identity Provider Choices dialog in an Office application](../images/add-in-auth-choices-dialog.png)</span></span>



3. <span data-ttu-id="43ea8-129">身份提供程序登录 - 身份提供程序将拥有其自己的 UI。</span><span class="sxs-lookup"><span data-stu-id="43ea8-129">Identity Provider Sign-in - The identity provider will have their own UI.</span></span> <span data-ttu-id="43ea8-130">Microsoft Azure Active Directory 允许自定义登录和访问面板页面，以便与服务保持一致的外观和体验。 [了解详细信息](/azure/active-directory/fundamentals/customize-branding)。</span><span class="sxs-lookup"><span data-stu-id="43ea8-130">Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service. [Learn More](/azure/active-directory/fundamentals/customize-branding).</span></span>
<span data-ttu-id="43ea8-131">![Office 应用程序中的身份提供程序登录对话框的屏幕截图](../images/add-in-auth-identity-sign-in.png)</span><span class="sxs-lookup"><span data-stu-id="43ea8-131">![A screenshot of the Identity Provider Sign-in dialog in an Office application](../images/add-in-auth-identity-sign-in.png)</span></span>

4. <span data-ttu-id="43ea8-132">进度 - 表示设置和 UI 加载时的进度。</span><span class="sxs-lookup"><span data-stu-id="43ea8-132">Progress - Indicate progress while settings and UI load.</span></span>
<span data-ttu-id="43ea8-133">![显示 Office 应用程序中进度指示器的对话框的屏幕截图](../images/add-in-auth-modal-interstitial.png)</span><span class="sxs-lookup"><span data-stu-id="43ea8-133">![A screenshot of a dialog that shows a progress indicator in an Office application](../images/add-in-auth-modal-interstitial.png)</span></span>

> [!NOTE] 
> <span data-ttu-id="43ea8-134">使用 Microsoft 的标识服务时，你将有机会使用可定制的浅色和深色主题的品牌登录按钮。</span><span class="sxs-lookup"><span data-stu-id="43ea8-134">When using Microsoft's Identity service you'll have the opportunity to use a branded sign-in button that is customizable to light and dark themes.</span></span><span data-ttu-id="43ea8-135">了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="43ea8-135"> Learn more.</span></span>

## <a name="single-sign-on-authentication-flow-preview"></a><span data-ttu-id="43ea8-136">单一登录身份验证流程（预览）</span><span class="sxs-lookup"><span data-stu-id="43ea8-136">Single Sign-On authentication flow (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="43ea8-137">目前，Word、Excel、Outlook 和 PowerPoint 在预览版中支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="43ea8-137">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="43ea8-138">有关单一登录支持的详细信息，请参阅  [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="43ea8-138">For more information about single sign-on support, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="43ea8-139">如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="43ea8-139">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="43ea8-140">若要了解如何执行此操作，请参阅  [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（Exchange Online：如何为租户启用新式验证）。</span><span class="sxs-lookup"><span data-stu-id="43ea8-140">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="43ea8-141">一旦用于生产加载项的单一登录正式发布后，即可使用该正式发布版获取流畅的最终用户体验。</span><span class="sxs-lookup"><span data-stu-id="43ea8-141">Once single sign-on is generally available for production add-ins, use it for the smoother end-user experience.</span></span> <span data-ttu-id="43ea8-142">Office 中的用户标识（Microsoft 帐户或 Office 365 标识）用于登录到加载项。</span><span class="sxs-lookup"><span data-stu-id="43ea8-142">The user's identity within Office (either a Microsoft Account or an Office 365 identity) is used to sign-in to your add-in.</span></span> <span data-ttu-id="43ea8-143">因此，用户只登录一次。</span><span class="sxs-lookup"><span data-stu-id="43ea8-143">As a result users only sign-in once.</span></span> <span data-ttu-id="43ea8-144">这样便使你的客户更容易上手，体验更为顺畅。</span><span class="sxs-lookup"><span data-stu-id="43ea8-144">This removes friction in the experience making it easier for your customers to get started.</span></span>

1. <span data-ttu-id="43ea8-145">安装加载项时，用户将会看到一个与以下窗口类似的同意窗口：![安装加载项时，Office 应用程序中的同意窗口的屏幕截图](../images/add-in-auth-SSO-consent-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="43ea8-145">As an add-in is being installed, a user will see a consent window similar to the one below: ![A screenshot of the consent window in an Office application when an add-in is being installed](../images/add-in-auth-SSO-consent-dialog.png)</span></span>
> [!NOTE]
> <span data-ttu-id="43ea8-146">加载项发布服务器将控制同意窗口中包含的徽标、字符串和权限范围。</span><span class="sxs-lookup"><span data-stu-id="43ea8-146">The add-in publisher will have control over the logo, strings and permission scopes included in the consent window.</span></span> <span data-ttu-id="43ea8-147">这一 UI 由 Microsoft 预配置。</span><span class="sxs-lookup"><span data-stu-id="43ea8-147">The UI is pre-configured by Microsoft.</span></span>

2. <span data-ttu-id="43ea8-148">加载项将在用户同意后加载。</span><span class="sxs-lookup"><span data-stu-id="43ea8-148">The add-in will load after the user consents.</span></span> <span data-ttu-id="43ea8-149">它可以提取并显示任何必要的用户自定义信息。</span><span class="sxs-lookup"><span data-stu-id="43ea8-149">It can extract and display any necessary user customized information.</span></span>
<span data-ttu-id="43ea8-150">![Office 应用程序功能区中显示的加载项按钮的屏幕截图](../images/add-in-ribbon.png)</span><span class="sxs-lookup"><span data-stu-id="43ea8-150">![A screenshot of an Office application with add-in buttons displayed in the ribbon](../images/add-in-ribbon.png)</span></span>

## <a name="see-also"></a><span data-ttu-id="43ea8-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="43ea8-151">See also</span></span>

- <span data-ttu-id="43ea8-152">详细了解[开发 SSO 加载项（预览版）](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="43ea8-152">Learn more about [developing SSO Add-ins (preview)](../develop/sso-in-office-add-ins.md)</span></span>
