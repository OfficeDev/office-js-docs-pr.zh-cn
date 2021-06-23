---
title: Office 外接程序的身份验证设计准则
ms.date: 02/09/2021
description: 了解如何在加载项中直观地设计登录Office注册页。
localization_priority: Normal
ms.openlocfilehash: bc047e6b001b7a491aa8117ba1b5901716ca1555
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076383"
---
# <a name="authentication-patterns"></a><span data-ttu-id="e6ed9-103">身份验证模式</span><span class="sxs-lookup"><span data-stu-id="e6ed9-103">Authentication patterns</span></span>

<span data-ttu-id="e6ed9-104">加载项可能需要用户登录或注册才能访问特性和功能。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-104">Add-ins may require users to sign-in or sign-up in order to access features and functionality.</span></span> <span data-ttu-id="e6ed9-105">用户名和密码的输入框或启动第三方凭据流的按钮是身份验证体验中常见的界面控件。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-105">Input boxes for username and password or buttons that start third party credential flows are common interface controls in authentication experiences.</span></span> <span data-ttu-id="e6ed9-106">获得简单高效的身份验证体验是让用户开始使用加载项的重要第一步。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-106">A simple and efficient authentication experience is an important first step to getting users started with your add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="e6ed9-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="e6ed9-107">Best practices</span></span>

|<span data-ttu-id="e6ed9-108">允许事项</span><span class="sxs-lookup"><span data-stu-id="e6ed9-108">Do</span></span>|<span data-ttu-id="e6ed9-109">禁止事项</span><span class="sxs-lookup"><span data-stu-id="e6ed9-109">Don't</span></span>|
|:----|:----|
|<span data-ttu-id="e6ed9-110">在登录之前，无需帐户即可介绍加载项或演示功能的价值。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-110">Prior to sign-in, describe the value of your add-in or demonstrate functionality without requiring an account.</span></span> |<span data-ttu-id="e6ed9-111">希望用户无需了解加载项的价值和好处即可登录。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-111">Expect users to sign-in without understanding the value and benefits of your add-in.</span></span>|
|<span data-ttu-id="e6ed9-112">指导用户通过身份验证流程，并在每个屏幕上使用主要的高度可见的按钮。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-112">Guide users through authentication flows with a primary, highly visible button on each screen.</span></span> |<span data-ttu-id="e6ed9-113">通过竞争按钮和号召性用语，引起对二级和三级任务的关注。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-113">Draw attention to secondary and tertiary tasks with competing buttons and calls to action.</span></span>|
|<span data-ttu-id="e6ed9-114">使用清晰的按钮标签来描述“登录”或“创建帐户”等特定任务。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-114">Use clear button labels that describe specific tasks like "Sign in" or "Create account".</span></span> |<span data-ttu-id="e6ed9-115">使用模糊的按钮标签，如“提交”或“入门”来指导用户完成身份验证流程。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-115">Use vague button labels like "Submit" or "Get started" to guide users through authentication flows.</span></span>|
|<span data-ttu-id="e6ed9-116">使用对话框将用户的注意力集中在身份验证表单上。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-116">Use a dialog to focus users' attention on authentication forms.</span></span> |<span data-ttu-id="e6ed9-117">使用初次运行体验和身份验证表单塞满任务窗格。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-117">Overcrowd your task pane with a first run experience and authentication forms.</span></span>|
|<span data-ttu-id="e6ed9-118">在流程中寻找细处的效率，如自动对焦输入框。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-118">Find small efficiencies in the flow like auto-focusing on input boxes.</span></span> |<span data-ttu-id="e6ed9-119">为交互添加如要求用户单击表单域等不必要的步骤。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-119">Add unnecessary steps to the interaction like requiring users to click into form fields.</span></span>|
|<span data-ttu-id="e6ed9-120">为用户提供一种注销和重新进行身份验证的方法。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-120">Provide a way for users to sign out and reauthenticate.</span></span> |<span data-ttu-id="e6ed9-121">强制用户卸载以切换标识。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-121">Force users to uninstall to switch identities.</span></span>|

## <a name="authentication-flow"></a><span data-ttu-id="e6ed9-122">身份验证流</span><span class="sxs-lookup"><span data-stu-id="e6ed9-122">Authentication flow</span></span>

1. <span data-ttu-id="e6ed9-123">首先运行 Placemat - 将登录按钮设置为加载项初次运行体验中的明确号召性用语。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-123">First Run Placemat - Place your sign-in button as a clear call-to action inside your add-in's first run experience.</span></span>

    ![Screenshot showing an add-in task pane in an Office application.](../images/add-in-fre-value-placemat.png)

1. <span data-ttu-id="e6ed9-125">标识提供者选项对话框 - 显示明确的标识提供者列表，包括用户名和密码表单（如适用）。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-125">Identity Provider Choices Dialog - Display a clear list of identity providers including a username and password form if applicable.</span></span> <span data-ttu-id="e6ed9-126">身份验证对话框处于打开状态时，加载项 UI 可能会被阻止。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-126">Your add-in UI may be blocked while the authentication dialog is open.</span></span>

    ![Screenshot showing the Identity Provider Choices dialog in an Office application.](../images/add-in-auth-choices-dialog.png)

1. <span data-ttu-id="e6ed9-128">身份提供程序登录 - 身份提供程序将拥有其自己的 UI。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-128">Identity Provider Sign-in - The identity provider will have their own UI.</span></span> <span data-ttu-id="e6ed9-129">Microsoft Azure Active Directory允许自定义登录和访问面板页面，以与服务保持一致的外观。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-129">Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service.</span></span> <span data-ttu-id="e6ed9-130">[了解更多信息](/azure/active-directory/fundamentals/customize-branding)。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-130">[Learn More](/azure/active-directory/fundamentals/customize-branding).</span></span>

    ![Screenshot showing the Identity Provider Sign-in dialog in an Office application.](../images/add-in-auth-identity-sign-in.png)

1. <span data-ttu-id="e6ed9-132">进度 - 表示设置和 UI 加载时的进度。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-132">Progress - Indicate progress while settings and UI load.</span></span>

    ![Screenshot showing a dialog with a progress indicator in an Office application.](../images/add-in-auth-modal-interstitial.png)

> [!NOTE]
> <span data-ttu-id="e6ed9-134">使用 Microsoft 的标识服务时，你将有机会使用可定制的浅色和深色主题的品牌登录按钮。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-134">When using Microsoft's Identity service you'll have the opportunity to use a branded sign-in button that is customizable to light and dark themes.</span></span> <span data-ttu-id="e6ed9-135">了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-135">Learn more.</span></span>

## <a name="single-sign-on-authentication-flow"></a><span data-ttu-id="e6ed9-136">单Sign-On身份验证流</span><span class="sxs-lookup"><span data-stu-id="e6ed9-136">Single Sign-On authentication flow</span></span>

> [!NOTE]
> <span data-ttu-id="e6ed9-137">Word、Excel、Outlook 和 PowerPoint 目前支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-137">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="e6ed9-138">有关单一登录支持的信息，请参阅 [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-138">For more information about single sign-on support, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="e6ed9-139">如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-139">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="e6ed9-140">若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-140">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="e6ed9-141">使用单一登录实现更流畅的最终用户体验。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-141">Use single sign-on for a smoother end-user experience.</span></span> <span data-ttu-id="e6ed9-142">Microsoft 帐户或 Office (中的用户标识Microsoft 365) 登录外接程序。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-142">The user's identity within Office (either a Microsoft Account or a Microsoft 365 identity) is used to sign in to your add-in.</span></span> <span data-ttu-id="e6ed9-143">因此，用户只能登录一次。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-143">As a result users only sign in once.</span></span> <span data-ttu-id="e6ed9-144">这样便使你的客户更容易上手，体验更为顺畅。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-144">This removes friction in the experience making it easier for your customers to get started.</span></span>

1. <span data-ttu-id="e6ed9-145">在安装外接程序时，用户将看到类似于以下内容的同意窗口：</span><span class="sxs-lookup"><span data-stu-id="e6ed9-145">As an add-in is being installed, a user will see a consent window similar to the one following:</span></span>

    ![Screenshot showing the consent window in an Office application when an add-in is installed.](../images/add-in-auth-SSO-consent-dialog.png)

    > [!NOTE]
    > <span data-ttu-id="e6ed9-147">加载项发布服务器将控制同意窗口中包含的徽标、字符串和权限范围。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-147">The add-in publisher will have control over the logo, strings and permission scopes included in the consent window.</span></span> <span data-ttu-id="e6ed9-148">这一 UI 由 Microsoft 预配置。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-148">The UI is pre-configured by Microsoft.</span></span>

1. <span data-ttu-id="e6ed9-149">加载项将在用户同意后加载。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-149">The add-in will load after the user consents.</span></span> <span data-ttu-id="e6ed9-150">它可以提取并显示任何必要的用户自定义信息。</span><span class="sxs-lookup"><span data-stu-id="e6ed9-150">It can extract and display any necessary user customized information.</span></span>

    ![Screenshot showing an Office application with add-in buttons displayed in the ribbon.](../images/add-in-ribbon.png)

## <a name="see-also"></a><span data-ttu-id="e6ed9-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e6ed9-152">See also</span></span>

- <span data-ttu-id="e6ed9-153">详细了解如何 [开发 SSO 加载项](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e6ed9-153">Learn more about [developing SSO Add-ins](../develop/sso-in-office-add-ins.md)</span></span>
