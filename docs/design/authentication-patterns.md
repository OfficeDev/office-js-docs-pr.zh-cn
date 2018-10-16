# <a name="authentication-patterns"></a>身份验证模式

加载项可能需要用户登录或注册才能访问特性和功能。用户名和密码输入框或启动第三方验证流程的按钮是身份验证体验中的常用接口控件。简单高效的身份验证体验是用户开始使用加载项的重要第一步。

## <a name="best-practices"></a>最佳做法

|建议|不建议|
|:----|:----|
|使用单点登录（SSO）将用户认证到您的加载项中。|要求用户从他们的个人Microsoft帐户或他们的Office 365帐户（工作或学校）中分别登录您的加载项。|
|在登录之前，请描述您的加载项的价值或演示其无需帐户的功能。 |其实用户无需了解加载项的价值和好处即可登录。|
|指导用户通过身份验证流程，并在每个屏幕上使用主要的高度可见的按钮. |借助竞争按钮和号召性行动，引起对二级和三级任务的关注。|
|使用清晰的按钮标签来描述“登录”或“创建帐户”等特定任务。   |使用模糊的按钮标签，如“提交”或“入门”来指导用户完成认证流程。|
|使用对话框将用户的注意力集中在认证表单上。    |使用第一次运行体验和身份验证表单超容您的任务窗格。|
|在流程中寻找细处的效率，如自动对焦输入框。 |为交互添加不必要的步骤，例如要求用户点击表单域。|
|为用户提供注销和重新认证的方法。    |强制用户卸载以切换身份。|

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 预览版支持单一登录 API。若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)。如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。若要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。


## <a name="authentication-flow"></a>身份验证流程
如果单一登录对用户不可用，请考虑替代身份验证流程。让用户选择直接使用您的服务或 Microsoft 等身份提供商进行登录。

1. 首次运行占位图片  - 在加载项首次运行体验中，将登录按钮作为明确的号召性操作。
![](../images/add-in-fre-value-placemat.png)

2. 身份验证提供商选择对话框 - 显示身份验证提供者的明确列表，包括用户名和密码表单（如果适用）。验证对话框打开时，加载项 UI 可能会被拦截。 ![](../images/add-in-auth-choices-dialog.png)



3. 身份验证提供商登录 - 身份验证提供商将拥有自己的用户界面。Microsoft Azure Active Directory 允许自定义登录和访问面板页面，以便与您的服务保持一致的外观。[了解更多信息](https://docs.microsoft.com/azure/active-directory/fundamentals/customize-branding) 。 ![](../images/add-in-auth-identity-sign-in.png)

4. 进度 - 在设置和 UI 加载时指示进度。
![](../images/add-in-auth-modal-interstitial.png)

> [!NOTE] 
> 使用 Microsoft 的身份识别服务时，您将有机会使用可定制的明亮和黑暗主题的品牌登录按钮。了解更多信息。

## <a name="single-sign-on-authentication-flow"></a>单一登录认证流程
单一登录仍处于预览。一旦公开发布，就可以使用它来获得最顺畅的最终用户体验。在 Office 中的用户身份用于登录到您的加载项。因此用户只需登录一次。这样便让您客户更易于上手，让操作更顺畅。

1. 当正在安装插件时，用户将看到一个类似于下面的同意窗口： ![](../images/add-in-auth-SSO-consent-dialog.png)
> [!NOTE]
> 加载项发布者可以控制同意窗口中包含的徽标、字符串和权限范围。这一 UI 由 Microsoft 预配置。

2. 加载项将在用户同意后加载。它可以提取并显示任何自定义的必要用户信息。 ![](../images/add-in-ribbon.png)

## <a name="see-also"></a>另请参阅
- 了解更多关于[开发 SSO 插件的信息](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)