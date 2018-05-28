---
title: ?? Visual Studio ??????????
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e03959294536eeb416a1531d2d281ba83f2d3732
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="5d62e-102">?? Visual Studio ??????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-102">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="5d62e-103">Office ?????? XML [????](../develop/add-in-manifests.md)???????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-103">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="5d62e-104">????????? Web ???????</span><span class="sxs-lookup"><span data-stu-id="5d62e-104">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="5d62e-105">???????? Visual Studio 2015 ?? Web ?????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-105">This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a><span data-ttu-id="5d62e-106">?? Visual Studio 2015 ?? Web ??</span><span class="sxs-lookup"><span data-stu-id="5d62e-106">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="5d62e-107">????????? Visual Studio 2015 ?? Web ???</span><span class="sxs-lookup"><span data-stu-id="5d62e-107">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="5d62e-108">????????????****???????????????????????****?</span><span class="sxs-lookup"><span data-stu-id="5d62e-108">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="5d62e-109">????**??????**???</span><span class="sxs-lookup"><span data-stu-id="5d62e-109">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="5d62e-110">??????????****???????????????????****???????</span><span class="sxs-lookup"><span data-stu-id="5d62e-110">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="5d62e-111">???????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-111">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="5d62e-p102">??????**??...**???????**????????**???????????????????? Microsoft Azure?????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="5d62e-114">????????????????????????????? [????????](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile)?</span><span class="sxs-lookup"><span data-stu-id="5d62e-114">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="5d62e-115">??**??????**???????**?? Web ??**????</span><span class="sxs-lookup"><span data-stu-id="5d62e-115">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="5d62e-p103">??**??? Web?**????????????????????[???? Visual Studio ???????????? Web ??](http://msdn.microsoft.com/en-us/library/dd465337.aspx)?</span><span class="sxs-lookup"><span data-stu-id="5d62e-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a><span data-ttu-id="5d62e-118">?? Visual Studio 2015 ??????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-118">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="5d62e-119">????????? Visual Studio 2015 ??????</span><span class="sxs-lookup"><span data-stu-id="5d62e-119">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="5d62e-120">????????****????????????****???</span><span class="sxs-lookup"><span data-stu-id="5d62e-120">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="5d62e-121">?????? Office ? SharePoint ????****?????</span><span class="sxs-lookup"><span data-stu-id="5d62e-121">The **Publish Office and SharePoint Add-ins** wizard appears.</span></span>
    
2. <span data-ttu-id="5d62e-122">???????????****???????????????????????? HTTPS URL????????****?</span><span class="sxs-lookup"><span data-stu-id="5d62e-122">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> 
    
    <span data-ttu-id="5d62e-p104">????? HTTPS ????? URL???????????????? HTTP ???????????????????? XML ?????????? HTTPS ????? HTTP ???</span><span class="sxs-lookup"><span data-stu-id="5d62e-p104">You must specify a URL that begins with the HTTPS prefix to complete this wizard. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="5d62e-125"> Azure ?????? HTTPS ???</span><span class="sxs-lookup"><span data-stu-id="5d62e-125">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="5d62e-126">Visual Studio ??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-126">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span> 
    
<span data-ttu-id="5d62e-p105">??????????? AppSource?????????????****????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="5d62e-p105">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="5d62e-p106">?????? XML ???????????[?????](../publish/publish.md)?XML ???? `app.publish` ???? `OfficeAppManifests` ?????</span><span class="sxs-lookup"><span data-stu-id="5d62e-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="5d62e-132">????</span><span class="sxs-lookup"><span data-stu-id="5d62e-132">See also</span></span>

- [<span data-ttu-id="5d62e-133">?? Office ???</span><span class="sxs-lookup"><span data-stu-id="5d62e-133">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="5d62e-134">???????? AppSource ? Office ????</span><span class="sxs-lookup"><span data-stu-id="5d62e-134">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)
    
