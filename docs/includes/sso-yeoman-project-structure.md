### <a name="configuration"></a><span data-ttu-id="e944f-101">配置</span><span class="sxs-lookup"><span data-stu-id="e944f-101">Configuration</span></span>

<span data-ttu-id="e944f-102">以下文件指定外接程序的配置设置。</span><span class="sxs-lookup"><span data-stu-id="e944f-102">The following files specify configuration settings for the add-in.</span></span>

- <span data-ttu-id="e944f-103">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="e944f-103">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="e944f-104">**./。** 项目根目录中的环境文件定义外接程序项目使用的常量。</span><span class="sxs-lookup"><span data-stu-id="e944f-104">The **./.ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>

### <a name="task-pane"></a><span data-ttu-id="e944f-105">任务窗格</span><span class="sxs-lookup"><span data-stu-id="e944f-105">Task pane</span></span> 

<span data-ttu-id="e944f-106">以下文件定义加载项的任务窗格 UI 和功能。</span><span class="sxs-lookup"><span data-stu-id="e944f-106">The following files define the add-in's task pane UI and functionality.</span></span>

- <span data-ttu-id="e944f-107">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="e944f-107">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>

- <span data-ttu-id="e944f-108">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="e944f-108">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>

- <span data-ttu-id="e944f-109">在 JavaScript 项目中， **/src/taskpane/taskpane.js**文件包含用于初始化加载项的代码。</span><span class="sxs-lookup"><span data-stu-id="e944f-109">In a JavaScript project, the **./src/taskpane/taskpane.js** file contains code to initialize the add-in.</span></span> <span data-ttu-id="e944f-110">在 TypeScript 项目中， **/src/taskpane/taskpane.ts**文件包含用于初始化外接程序的代码，以及使用 Office JavaScript 库将数据从 Microsoft Graph 添加到 Office 文档的代码。</span><span class="sxs-lookup"><span data-stu-id="e944f-110">In a TypeScript project, the **./src/taskpane/taskpane.ts** file contains code to initialize the add-in and also code that uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>

### <a name="authentication"></a><span data-ttu-id="e944f-111">身份验证</span><span class="sxs-lookup"><span data-stu-id="e944f-111">Authentication</span></span>

<span data-ttu-id="e944f-112">以下文件可帮助 SSO 进程并将数据写入 Office 文档。</span><span class="sxs-lookup"><span data-stu-id="e944f-112">The following files facilitate the SSO process and write data to the Office document.</span></span>

- <span data-ttu-id="e944f-113">在 JavaScript 项目中， **/src/helpers/documentHelper.js**文件包含使用 Office JavaScript 库将数据从 Microsoft Graph 添加到 Office 文档的代码。</span><span class="sxs-lookup"><span data-stu-id="e944f-113">In a JavaScript project, the **./src/helpers/documentHelper.js** file contains code that uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span> <span data-ttu-id="e944f-114">在 TypeScript 项目中没有此类文件;使用 Office JavaScript 库将数据从 Microsoft Graph 添加到 Office 文档的代码在 **/src/taskpane/taskpane.ts**中存在。</span><span class="sxs-lookup"><span data-stu-id="e944f-114">There is no such file in a TypeScript project; the code that uses the Office JavaScript library to add the data from Microsoft Graph to the Office document exists in **./src/taskpane/taskpane.ts** instead.</span></span>

- <span data-ttu-id="e944f-115">**./Src/helpers/fallbackauthdialog.html**文件是为回退身份验证策略加载 JavaScript 的无 UI 页面。</span><span class="sxs-lookup"><span data-stu-id="e944f-115">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the JavaScript for the fallback authentication strategy.</span></span>

- <span data-ttu-id="e944f-116">**/Src/helpers/fallbackauthdialog.js**文件包含使用 msal 登录用户的回退身份验证策略的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="e944f-116">The **./src/helpers/fallbackauthdialog.js** file contains the JavaScript for the fallback authentication strategy that signs in the user with msal.js.</span></span>

- <span data-ttu-id="e944f-117">在不支持 SSO 身份验证的情况下， **/src/helpers/fallbackauthhelper.js**文件包含用于调用回退身份验证策略的任务窗格 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="e944f-117">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication strategy in scenarios when SSO authentication is not supported.</span></span>

- <span data-ttu-id="e944f-118">**./src/helpers/ssoauthhelper.js** 文件包含调用 SSO API、`getAccessToken` 的 JavaScript ，接收引导令牌，针对 Microsoft Graph 访问令牌启动引导令牌交换，同时调用 Microsoft Graph 以获得数据。</span><span class="sxs-lookup"><span data-stu-id="e944f-118">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>