---
title: Office 加载项的可用性测试
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 38f0416d56f3fc43c6d5f68df9b5c84586b03c8c
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270927"
---
# <a name="usability-testing-for-office-add-ins"></a><span data-ttu-id="0222e-102">Office 加载项的可用性测试</span><span class="sxs-lookup"><span data-stu-id="0222e-102">Usability testing for Office Add-ins</span></span>

<span data-ttu-id="0222e-p101">出色的外接程序设计会考虑到用户行为。因为自己的预想会影响设计决策，所以务必要通过实际用户测试设计来确保客户可正常使用外接程序。</span><span class="sxs-lookup"><span data-stu-id="0222e-p101">A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.</span></span> 

<span data-ttu-id="0222e-p102">可以使用不同方法运行可用性测试。对于许多外接程序开发人员而言，远程、未经审阅的可用性研究最为节省时间且最具有成本效益。一些热门测试服务使其变得简单；以下是部分示例：</span><span class="sxs-lookup"><span data-stu-id="0222e-p102">You can run usability tests in different ways. For many add-in developers, remote, unmoderated usability studies are the most time and cost effective. Several popular testing services make this easy; the following are some examples:</span></span> 

 - [<span data-ttu-id="0222e-108">UserTesting.com</span><span class="sxs-lookup"><span data-stu-id="0222e-108">UserTesting.com</span></span>](https://www.UserTesting.com)
 - [<span data-ttu-id="0222e-109">Optimalworkshop.com</span><span class="sxs-lookup"><span data-stu-id="0222e-109">Optimalworkshop.com</span></span>](https://www.Optimalworkshop.com)
 - [<span data-ttu-id="0222e-110">Userzoom.com</span><span class="sxs-lookup"><span data-stu-id="0222e-110">Userzoom.com</span></span>](https://www.Userzoom.com)

<span data-ttu-id="0222e-111">这些测试服务可帮助你简化测试计划的创建并且不需要寻找参与者或审阅测试。</span><span class="sxs-lookup"><span data-stu-id="0222e-111">These testing services help you to streamline test plan creation and remove the need to seek out participants or moderate the tests.</span></span> 

<span data-ttu-id="0222e-p103">你只需五名参与者即可发现设计中的大多数可用性问题。在整个开发周期内定期进行小型测试，以确保产品以用户为中心。</span><span class="sxs-lookup"><span data-stu-id="0222e-p103">You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.</span></span>

> [!NOTE]
> <span data-ttu-id="0222e-p104">建议跨多个平台测试加载项的可用性。若要[将加载项发布到 AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)，加载项必须适用于[支持已定义方法的所有平台](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="0222e-p104">We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store), it must work on all [platforms that support the methods that you define](../overview/office-add-in-availability.md).</span></span>

## <a name="1---sign-up-for-a-testing-service"></a><span data-ttu-id="0222e-116">1. 注册测试服务</span><span class="sxs-lookup"><span data-stu-id="0222e-116">1.   Sign up for a testing service</span></span>

<span data-ttu-id="0222e-117">有关详细信息，请参阅[选择联机工具进行未加管制的远程用户测试](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)。</span><span class="sxs-lookup"><span data-stu-id="0222e-117">For more information, see [Selecting an Online Tool for Unmoderated Remote User Testing.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)</span></span>

## <a name="2-develop-your-research-questions"></a><span data-ttu-id="0222e-118">2.制定研究问题</span><span class="sxs-lookup"><span data-stu-id="0222e-118">2. Develop your research questions</span></span>
 
<span data-ttu-id="0222e-p105">研究问题定义研究的目标并指导测试计划。这些问题将帮助你确定要招募的参与者及其要执行的任务。将研究问题尽可能地具体化。还可以尽量回答较为宽泛的问题。</span><span class="sxs-lookup"><span data-stu-id="0222e-p105">Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.</span></span>
 
<span data-ttu-id="0222e-123">以下是研究问题的一些示例：</span><span class="sxs-lookup"><span data-stu-id="0222e-123">The following are some examples of research questions:</span></span>
  
<span data-ttu-id="0222e-124">**具体**</span><span class="sxs-lookup"><span data-stu-id="0222e-124">**Specific**</span></span>  

 - <span data-ttu-id="0222e-125">用户是否注意到了登陆页面上的“免费试用版”链接？</span><span class="sxs-lookup"><span data-stu-id="0222e-125">Do users notice the "free trial" link on the landing page?</span></span>
 - <span data-ttu-id="0222e-126">用户将内容从外接程序插入他们的文档时，用户是否知道内容在文档何处插入？</span><span class="sxs-lookup"><span data-stu-id="0222e-126">When users insert content from the add-in to their document, do they understand where in the document it is inserted?</span></span>

<span data-ttu-id="0222e-127">**宽泛**</span><span class="sxs-lookup"><span data-stu-id="0222e-127">**Broad**</span></span>  

 - <span data-ttu-id="0222e-128">用户在我们的外接程序中遇到的最大问题是什么？</span><span class="sxs-lookup"><span data-stu-id="0222e-128">What are the biggest pain points for the user in our add-in?</span></span>
 - <span data-ttu-id="0222e-129">用户在单击命令栏中的图标前是否了解他们的含义？</span><span class="sxs-lookup"><span data-stu-id="0222e-129">Do users understand the meaning of the icons in our command bar, before they click on them?</span></span>
 - <span data-ttu-id="0222e-130">用户能否轻松地找到设置菜单？</span><span class="sxs-lookup"><span data-stu-id="0222e-130">Can users easily find the settings menu?</span></span>

<span data-ttu-id="0222e-p106">获取从发现外接程序到安装并使用外接程序的整个用户操作体验的相关数据至关重要。考虑可解决外接程序用户体验以下方面的研究问题：</span><span class="sxs-lookup"><span data-stu-id="0222e-p106">It’s important to get data on the entire user journey – from discovering your add-in, to installing and using it. Consider research questions that address the following aspects of the add-in user experience:</span></span>
 
 - <span data-ttu-id="0222e-133">在 AppSource 中查找加载项</span><span class="sxs-lookup"><span data-stu-id="0222e-133">Finding your add-in in AppSource</span></span>
 - <span data-ttu-id="0222e-134">选择安装加载项</span><span class="sxs-lookup"><span data-stu-id="0222e-134">Choosing to install your add-in</span></span>
 - <span data-ttu-id="0222e-135">初次运行体验</span><span class="sxs-lookup"><span data-stu-id="0222e-135">First run experience</span></span>
 - <span data-ttu-id="0222e-136">功能区命令</span><span class="sxs-lookup"><span data-stu-id="0222e-136">Ribbon commands</span></span>
 - <span data-ttu-id="0222e-137">外接程序 UI</span><span class="sxs-lookup"><span data-stu-id="0222e-137">Add-in UI</span></span>
 - <span data-ttu-id="0222e-138">外接程序如何与 Office 应用程序的文档空间交互</span><span class="sxs-lookup"><span data-stu-id="0222e-138">How the add-in interacts with the document space of the Office application</span></span>
 - <span data-ttu-id="0222e-139">用户对任意内容插入流的掌控程度如何</span><span class="sxs-lookup"><span data-stu-id="0222e-139">How much control the user has over any content insertion flows</span></span>

<span data-ttu-id="0222e-140">有关详细信息，请参阅[收集实际响应与主观数据](https://help.usertesting.com/hc/zh-CN/articles/115003378572-Writing-effective-questions)。</span><span class="sxs-lookup"><span data-stu-id="0222e-140">For more information, see [Gathering factual responses vs. subjective data](https://help.usertesting.com/hc/zh-CN/articles/115003378572-Writing-effective-questions).</span></span>
 
## <a name="3-identify-participants-to-target"></a><span data-ttu-id="0222e-141">3.确定所要面向的参与者</span><span class="sxs-lookup"><span data-stu-id="0222e-141">3. Identify participants to target</span></span>
 
<span data-ttu-id="0222e-p107">通过远程测试服务，你可以控制测试参与者的许多特性。认真考虑想要将哪类用户确定为目标。在数据收集的早期阶段，最好招募各种类型的参与者以识别出较为显著的可用性问题。后面可以选择将类似高级 Office 用户、特定职业或特定年龄段的组确定为目标。</span><span class="sxs-lookup"><span data-stu-id="0222e-p107">Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.</span></span>
 
## <a name="4-create-the-participant-screener"></a><span data-ttu-id="0222e-146">4.创建参与者筛选器</span><span class="sxs-lookup"><span data-stu-id="0222e-146">4. Create the participant screener</span></span>
 
<span data-ttu-id="0222e-p108">筛选程序是将向潜在测试参与者提供的问题和要求集，以对其进行测试筛选。请注意 UserTesting.com 等服务的参与者参加测试是想要获得经济收益。如果想要将特定用户从测试排除，那么在筛选程序中加入技巧性问题是个不错的主意。</span><span class="sxs-lookup"><span data-stu-id="0222e-p108">The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test.</span></span> 
 
<span data-ttu-id="0222e-150">例如，想要找出熟悉 GitHub 的参与者，要筛选出对自己进行了不当描述的用户，包括可能的答案列表中的不实之处。</span><span class="sxs-lookup"><span data-stu-id="0222e-150">For example, if you want to find participants who are familiar with GitHub, to filter out users who might misrepresent themselves, include fakes in the list of possible answers.</span></span>

<span data-ttu-id="0222e-151">**熟悉下面哪些源代码存储库？**</span><span class="sxs-lookup"><span data-stu-id="0222e-151">**Which of the following source code repositories are you familiar with?**</span></span>  
 <span data-ttu-id="0222e-p109">a. SourceShelf  [*拒绝*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p109">a. SourceShelf  [*Reject*]</span></span>  
 <span data-ttu-id="0222e-p110">b. CodeContainer  [*拒绝*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p110">b. CodeContainer  [*Reject*]</span></span>  
 <span data-ttu-id="0222e-p111">c. GitHub  [*必选*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p111">c. GitHub  [*Must select*]</span></span>  
 <span data-ttu-id="0222e-p112">d. BitBucket  [*可选*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p112">d. BitBucket  [*May select*]</span></span>  
 <span data-ttu-id="0222e-p113">e. CloudForge  [*可选*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p113">e. CloudForge  [*May select*]</span></span>  

<span data-ttu-id="0222e-162">如果计划测试外接程序的实时生成，以下问题可以筛选出可以执行此任务的用户。</span><span class="sxs-lookup"><span data-stu-id="0222e-162">If you are planning to test a live build of your add-in, the following questions can screen for users who will be able to do this.</span></span> 

<span data-ttu-id="0222e-163">**该测试需要安装最新版本的 Microsoft PowerPoint。是否拥有最新版本的 PowerPoint？**</span><span class="sxs-lookup"><span data-stu-id="0222e-163">**This test requires you to have the latest version of Microsoft PowerPoint. Do you have the latest version of PowerPoint?**</span></span>  
 <span data-ttu-id="0222e-p114">a. 是 [*必选*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p114">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="0222e-p115">b. 否 [*拒绝*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p115">b. No [*Reject*]</span></span>  
 <span data-ttu-id="0222e-p116">c. 不知道 [*拒绝*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p116">c. I don’t know [*Reject*]</span></span>  

<span data-ttu-id="0222e-170">**此测试要求安装适用于 PowerPoint 的免费加载项，并创建免费帐户以进行使用。是否愿意安装加载项并创建免费帐户？**</span><span class="sxs-lookup"><span data-stu-id="0222e-170">**This test requires you to install a free add-in for PowerPoint 2016, and create a free account to use it. Are you willing to install an add-in and create a free account?**</span></span>  
 <span data-ttu-id="0222e-p117">a. 是 [*必选*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p117">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="0222e-p118">b. 否 [*拒绝*]</span><span class="sxs-lookup"><span data-stu-id="0222e-p118">b. No [*Reject*]</span></span>  

<span data-ttu-id="0222e-175">有关详细信息，请参阅[筛选程序问题最佳做法](https://help.usertesting.com/hc/zh-CN/articles/115003370731-Screener-question-best-practices)。</span><span class="sxs-lookup"><span data-stu-id="0222e-175">For more information, see [Screener Questions Best Practices.](https://help.usertesting.com/hc/zh-CN/articles/115003370731-Screener-question-best-practices)</span></span>
 
## <a name="5-create-tasks-and-questions-for-participants"></a><span data-ttu-id="0222e-176">5.创建针对参与者的任务和问题</span><span class="sxs-lookup"><span data-stu-id="0222e-176">5. Create tasks and questions for participants</span></span>
 
<span data-ttu-id="0222e-p119">尝试对要测试的内容设置优先级，以便限制针对参与者的任务和问题数量。一些服务仅在特定时间内向参与者付费，你需要确保不会超过该时间。</span><span class="sxs-lookup"><span data-stu-id="0222e-p119">Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.</span></span>

<span data-ttu-id="0222e-p120">尽可能地尝试观察参与者行为，而不是向其提问。如果需要询问其行为，询问参与者过去做过什么，而不是询问其在某个场景下会做什么。这样提供的结果往往更为可靠。</span><span class="sxs-lookup"><span data-stu-id="0222e-p120">Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.</span></span>
 
<span data-ttu-id="0222e-p121">未经审阅的测试的主要挑战在于确保参与者了解你的任务和方案。你的指示应*简洁明了*。不可避免的是，如果可能存在混淆，则某些人会感到困惑。</span><span class="sxs-lookup"><span data-stu-id="0222e-p121">The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.</span></span> 

<span data-ttu-id="0222e-p122">在测试期间的任何给定时刻，都不要假设用户会位于其应位于的屏幕上。考虑告诉用户要开始下一个任务他们需要位于哪个屏幕。</span><span class="sxs-lookup"><span data-stu-id="0222e-p122">Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.</span></span> 

<span data-ttu-id="0222e-187">有关详细信息，请参阅[编写出色任务](https://help.usertesting.com/hc/zh-CN/articles/115003371651-Writing-great-tasks)。</span><span class="sxs-lookup"><span data-stu-id="0222e-187">For more information, see [Writing Great Tasks.](https://help.usertesting.com/hc/zh-CN/articles/115003371651-Writing-great-tasks)</span></span>

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a><span data-ttu-id="0222e-188">6.创建用于匹配任务和问题的原型</span><span class="sxs-lookup"><span data-stu-id="0222e-188">6. Create a prototype to match the tasks and questions</span></span>
 
<span data-ttu-id="0222e-189">可以测试实时加载项，或者可以测试原型。</span><span class="sxs-lookup"><span data-stu-id="0222e-189">You can either test your live add-in, or you can test a prototype.</span></span> <span data-ttu-id="0222e-190">注意，如果要测试实时加载项，则需要筛选出已安装 Office、愿意安装加载项且愿意注册帐户的参与者（除非你具有可以向参与者提供的登录凭据）。然后需要确保他们成功安装加载项。</span><span class="sxs-lookup"><span data-stu-id="0222e-190">You can either test your live add-in, or you can test a prototype. Keep in mind that if you want to test the live add-in, you need to screen for participants that have Office 2016, are willing to install the add-in, and are willing to sign up for an account (unless you have logon credentials to provide them.) You'll then need to make sure that they successfully install your add-in.</span></span> 

<span data-ttu-id="0222e-p124">通常，逐步指导用户如何安装外接程序需要大约 5 分钟。以下是简洁明了的安装步骤示例。请根据测试的具体情况调整步骤。</span><span class="sxs-lookup"><span data-stu-id="0222e-p124">On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.</span></span>

<span data-ttu-id="0222e-194">**请使用以下说明安装适用于 PowerPoint 的加载项（在此处插入加载项名称）：**</span><span class="sxs-lookup"><span data-stu-id="0222e-194">**Please install the (insert your add-in name here) add-in for PowerPoint 2016, using the following instructions:**</span></span> 

1. <span data-ttu-id="0222e-195">打开 Microsoft PowerPoint。</span><span class="sxs-lookup"><span data-stu-id="0222e-195">Open Microsoft PowerPoint 2016.</span></span>
2. <span data-ttu-id="0222e-196">选择“**空白演示文稿**”。</span><span class="sxs-lookup"><span data-stu-id="0222e-196">Select **Blank Presentation.**</span></span>
3. <span data-ttu-id="0222e-197">转到“**插入 > 我的外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="0222e-197">Go to **Insert > My Add-ins.**</span></span>
5. <span data-ttu-id="0222e-198">在弹出窗口中，选择“**应用商店**”。</span><span class="sxs-lookup"><span data-stu-id="0222e-198">In the popup window, choose **Store.**</span></span>
6. <span data-ttu-id="0222e-199">在搜索框中键入（外接程序名称）。</span><span class="sxs-lookup"><span data-stu-id="0222e-199">Type (Add-in name) in the search box.</span></span>
7. <span data-ttu-id="0222e-200">选择（外接程序名称）。</span><span class="sxs-lookup"><span data-stu-id="0222e-200">Choose (Add-in name).</span></span>
8. <span data-ttu-id="0222e-201">花费一些时间查看“应用商店”页面以熟悉外接程序。</span><span class="sxs-lookup"><span data-stu-id="0222e-201">Take a moment to look at the Store page to familiarize yourself with the add-in.</span></span>
9. <span data-ttu-id="0222e-202">选择“**添加**”安装外接程序。</span><span class="sxs-lookup"><span data-stu-id="0222e-202">Choose **Add** to install the add-in.</span></span>

<span data-ttu-id="0222e-p125">可以以任意基本的交互和外观一致性来测试原型。对于更为复杂的链接和交互性，请考虑使用 [InVision](https://www.invisionapp.com) 等原型制作工具。如果只想测试静态屏幕，可以在线托管图像并向参与者发送相应的 URL，或向其提供指向在线 PowerPoint 演示文稿的链接。</span><span class="sxs-lookup"><span data-stu-id="0222e-p125">You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation.</span></span> 

## <a name="7-run-a-pilot-test"></a><span data-ttu-id="0222e-206">7.运行试点测试</span><span class="sxs-lookup"><span data-stu-id="0222e-206">7. Run a pilot test</span></span>

<span data-ttu-id="0222e-p126">正确设置原型和任务/问题列表可能会比较困难。用户可能会对任务感到疑惑，或者对原型不知所措。应通过 1-3 名用户运行试点测试来解决测试格式存在的难以避免的问题。这将有助于确保问题清楚明了、原型得到正确设置并捕获所寻找的数据类型。</span><span class="sxs-lookup"><span data-stu-id="0222e-p126">It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.</span></span>

## <a name="8-run-the-test"></a><span data-ttu-id="0222e-211">8.运行测试</span><span class="sxs-lookup"><span data-stu-id="0222e-211">8. Run the test</span></span>

<span data-ttu-id="0222e-p127">指令进行测试后，参与者完成测试后你将获得电子邮件通知。除非你将特定参与者组确定为目标，否则测试通常会在数小时内完成。</span><span class="sxs-lookup"><span data-stu-id="0222e-p127">After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.</span></span>

## <a name="9-analyze-results"></a><span data-ttu-id="0222e-214">9.分析结果</span><span class="sxs-lookup"><span data-stu-id="0222e-214">9. Analyze results</span></span>

<span data-ttu-id="0222e-p128">在这一部分中，你将尝试分析所收集到的数据。在观看测试视频时，记录用户遇到的问题和成功之处。避免尝试在查看所有结果后才解释数据的含义。</span><span class="sxs-lookup"><span data-stu-id="0222e-p128">This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.</span></span> 

<span data-ttu-id="0222e-p129">单个参与者具有可用性问题不足以作为更改设计的依据。两个或更多参与者遇到同一问题则表明普通人群中的其他用户也会遇到此问题。</span><span class="sxs-lookup"><span data-stu-id="0222e-p129">A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.</span></span>

<span data-ttu-id="0222e-p130">通常，要谨慎对待使用数据作出结论的方式。不要陷入尝试将数据匹配特定叙述的困境；对数据实际证明、驳斥或者无法提供任何相关见解的内容实事求是。保持开放的心态；用户行为经常会违背设计人员的预期。</span><span class="sxs-lookup"><span data-stu-id="0222e-p130">In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.</span></span>
 

## <a name="see-also"></a><span data-ttu-id="0222e-223">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0222e-223">See also</span></span>
 
 - [<span data-ttu-id="0222e-224">如何执行可用性测试</span><span class="sxs-lookup"><span data-stu-id="0222e-224">How to Conduct Usability Testing</span></span>](https://whatpixel.com/howto-conduct-usability-testing/)  
 - [<span data-ttu-id="0222e-225">用户测试的最佳做法</span><span class="sxs-lookup"><span data-stu-id="0222e-225">Best Practices for UserTesting</span></span>](https://help.usertesting.com/hc/zh-CN/articles/115003370231-Best-practices-for-UserTesting)  
 - [<span data-ttu-id="0222e-226">最小化偏差</span><span class="sxs-lookup"><span data-stu-id="0222e-226">Minimizing Bias</span></span>](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
