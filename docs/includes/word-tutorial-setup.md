<span data-ttu-id="c9d85-101">在本教程中，请先设置开发项目。</span><span class="sxs-lookup"><span data-stu-id="c9d85-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="c9d85-p101">此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="c9d85-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

> [!TIP]
> <span data-ttu-id="c9d85-104">如果尚未这样做，请参阅[生成首个 Word 加载项](../quickstarts/word-quickstart.md?tabs=visual-studio-code)。</span><span class="sxs-lookup"><span data-stu-id="c9d85-104">If you haven't already done so, please read [Build your first Word add-in](../quickstarts/word-quickstart.md?tabs=visual-studio-code).</span></span> <span data-ttu-id="c9d85-105">特别是，请务必了解如何旁加载 Word 加载项以供测试。</span><span class="sxs-lookup"><span data-stu-id="c9d85-105">In particular, be sure that you know how to sideload a Word add-in for testing.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c9d85-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="c9d85-106">Prerequisites</span></span>

<span data-ttu-id="c9d85-107">若要学习本教程，需要安装以下各项。</span><span class="sxs-lookup"><span data-stu-id="c9d85-107">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="c9d85-108">Word 2016 版本 1711（生成号 8730.1000 即点即用）或更高版本。</span><span class="sxs-lookup"><span data-stu-id="c9d85-108">Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="c9d85-109">可能必须成为 Office 预览体验成员，才能获取此版本。</span><span class="sxs-lookup"><span data-stu-id="c9d85-109">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="c9d85-110">有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。</span><span class="sxs-lookup"><span data-stu-id="c9d85-110">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>
- [<span data-ttu-id="c9d85-111">Node 和 npm</span><span class="sxs-lookup"><span data-stu-id="c9d85-111">Node and npm</span></span>](https://nodejs.org/en/) 
- <span data-ttu-id="c9d85-112">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="c9d85-112">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="setup"></a><span data-ttu-id="c9d85-113">设置</span><span class="sxs-lookup"><span data-stu-id="c9d85-113">Setup</span></span>

1. <span data-ttu-id="c9d85-114">克隆 [Word 加载项教程](https://github.com/OfficeDev/Word-Add-in-Tutorial) GitHub 存储库。</span><span class="sxs-lookup"><span data-stu-id="c9d85-114">Clone the GitHub repository [Word Add-in Tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span></span>
2. <span data-ttu-id="c9d85-115">|||UNTRANSLATED_CONTENT_START|||Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="c9d85-115">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
3. <span data-ttu-id="c9d85-116">运行命令 `npm install`，以安装 package.json 文件中列出的工具和库。</span><span class="sxs-lookup"><span data-stu-id="c9d85-116">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 
4. <span data-ttu-id="c9d85-117">按照[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中的步骤操作，信任开发计算机操作系统的证书。</span><span class="sxs-lookup"><span data-stu-id="c9d85-117">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

