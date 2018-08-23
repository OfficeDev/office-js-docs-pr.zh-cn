<span data-ttu-id="95cac-101">在本教程中，请先设置开发项目。</span><span class="sxs-lookup"><span data-stu-id="95cac-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="95cac-102">此为 Excel 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="95cac-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="95cac-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="95cac-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="95cac-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="95cac-104">Prerequisites</span></span>

<span data-ttu-id="95cac-105">若要使用本教程，需要安装以下项。</span><span class="sxs-lookup"><span data-stu-id="95cac-105">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="95cac-106">Excel 2016 版本 1711（生成号 8730.1000 即点即用）或更高版本。</span><span class="sxs-lookup"><span data-stu-id="95cac-106">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="95cac-107">可能必须成为 Office 预览体验成员，才能获取此版本。</span><span class="sxs-lookup"><span data-stu-id="95cac-107">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="95cac-108">有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。</span><span class="sxs-lookup"><span data-stu-id="95cac-108">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>
- [<span data-ttu-id="95cac-109">Node 和 npm</span><span class="sxs-lookup"><span data-stu-id="95cac-109">Node and npm</span></span>](https://nodejs.org/en/) 
- <span data-ttu-id="95cac-110">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="95cac-110">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="setup"></a><span data-ttu-id="95cac-111">设置</span><span class="sxs-lookup"><span data-stu-id="95cac-111">Setup</span></span>

1. <span data-ttu-id="95cac-112">克隆 GitHub 存储库 [Excel 加载项教程](https://github.com/OfficeDev/Excel-Add-in-Tutorial)。</span><span class="sxs-lookup"><span data-stu-id="95cac-112">Clone the GitHub repository [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>
2. <span data-ttu-id="95cac-113">打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。</span><span class="sxs-lookup"><span data-stu-id="95cac-113">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
3. <span data-ttu-id="95cac-114">运行命令 `npm install`，以安装 package.json 文件中列出的工具和库。</span><span class="sxs-lookup"><span data-stu-id="95cac-114">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 
4. <span data-ttu-id="95cac-115">按照[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中的步骤操作，信任开发计算机操作系统的证书。</span><span class="sxs-lookup"><span data-stu-id="95cac-115">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

