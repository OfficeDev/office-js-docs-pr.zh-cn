---
title: Office 加载项的可用性测试
description: 了解如何使用真实用户测试外接程序设计。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 49a2af983615779160886961e8269e4588d0fc9e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810279"
---
# <a name="usability-testing-for-office-add-ins"></a>Office 加载项的可用性测试

A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.

可以使用不同方法运行可用性测试。 对于许多外接程序开发人员而言，远程、未经审阅的可用性研究最为节省时间且最具有成本效益。 几个流行的测试服务使这一点变得容易：下面是一些示例。

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

这些测试服务可帮助你简化测试计划的创建并且不需要寻找参与者或审阅测试。

You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.

> [!NOTE]
> We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](/javascript/api/requirement-sets).

## <a name="1-sign-up-for-a-testing-service"></a>1. 注册测试服务

有关详细信息，请参阅[选择联机工具进行未加管制的远程用户测试](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)。

## <a name="2-develop-your-research-questions"></a>2.制定研究问题

Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.

下面是研究问题的一些示例。

**具体**

- 用户是否注意到了登陆页面上的“免费试用版”链接？
- 用户将内容从外接程序插入他们的文档时，用户是否知道内容在文档何处插入？

**宽泛**

- 用户在我们的外接程序中遇到的最大问题是什么？
- 用户在单击命令栏中的图标前是否了解他们的含义？
- 用户能否轻松地找到设置菜单？

获取从发现外接程序到安装并使用外接程序的整个用户操作体验的相关数据至关重要。 考虑解决加载项用户体验的以下方面的研究问题。

- 在 AppSource 中查找加载项
- 选择安装加载项
- 初次运行体验
- 功能区命令
- 外接程序 UI
- 外接程序如何与 Office 应用程序的文档空间交互
- 用户对任意内容插入流的掌控程度如何

有关详细信息，请参阅[收集实际响应与主观数据](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions)。

## <a name="3-identify-participants-to-target"></a>3.确定所要面向的参与者

Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.

## <a name="4-create-the-participant-screener"></a>4.创建参与者筛选器

The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test. 

例如，想要找出熟悉 GitHub 的参与者，要筛选出对自己进行了不当描述的用户，包括可能的答案列表中的不实之处。

**熟悉下面哪些源代码存储库？**  
 a. SourceShelf  [*Reject*]  
 b. CodeContainer  [*Reject*]  
 c. GitHub  [*Must select*]  
 d. BitBucket  [*May select*]  
 e. CloudForge  [*May select*]  

如果计划测试外接程序的实时生成，以下问题可以筛选出可以执行此任务的用户。

**该测试需要安装最新版本的 Microsoft PowerPoint。是否拥有最新版本的 PowerPoint？**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  
 c. I don’t know [*Reject*]  

**此测试要求安装适用于 PowerPoint 的免费加载项，并创建免费帐户以进行使用。是否愿意安装加载项并创建免费帐户？**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  

有关详细信息，请参阅[筛选程序问题最佳做法](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices)。

## <a name="5-create-tasks-and-questions-for-participants"></a>5.创建针对参与者的任务和问题

Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.

Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.

The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.

Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.

有关详细信息，请参阅[编写出色任务](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks)。

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6.创建用于匹配任务和问题的原型

可以测试实时加载项，或者可以测试原型。 注意，如果要测试实时加载项，则需要筛选出已安装 Office、愿意安装加载项且愿意注册帐户的参与者（除非你具有可以向参与者提供的登录凭据）。然后需要确保他们成功安装加载项。

On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.

**请按照以下说明安装 (在此处插入外接程序名称) 适用于 PowerPoint 的外接程序。**

1. 打开 Microsoft PowerPoint。
1. 选择“**空白演示文稿**”。
1. 转到 **“插入** > **我的外接程序**”。
1. 在弹出窗口中，选择“ **应用商店**”。
1. 在搜索框中键入（外接程序名称）。
1. 选择（外接程序名称）。
1. 花费一些时间查看“应用商店”页面以熟悉外接程序。
1. 选择“**添加**”安装外接程序。

You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation. 

## <a name="7-run-a-pilot-test"></a>7.运行试点测试

It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.

## <a name="8-run-the-test"></a>8.运行测试

After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.

## <a name="9-analyze-results"></a>9.分析结果

This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.

A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.

In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.

## <a name="see-also"></a>另请参阅

- [如何执行可用性测试](https://whatpixel.com/howto-conduct-usability-testing/)  
- [用户测试的最佳做法](https://help.usertesting.com/hc/articles/115003370231-Best-practices-for-UserTesting)  
- [最小化偏差](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
