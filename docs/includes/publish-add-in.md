<span data-ttu-id="32fa9-101">Office 加载项由一个 Web 应用程序和一个清单文件构成。</span><span class="sxs-lookup"><span data-stu-id="32fa9-101">An Office Add-in consists of a web application and a manifest file.</span></span> <span data-ttu-id="32fa9-102">Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="32fa9-102">The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span> 

<span data-ttu-id="32fa9-103">开发加载项时，可在本地 Web 服务器 (`localhost`) 上运行加载项，但如果已准备好发布该加载项供其他用户访问，则需要将 Web 应用程序部署到 Web 服务器或 Web 托管服务（例如 Microsoft Azure），并更新清单以指定所部署的应用程序的 URL。</span><span class="sxs-lookup"><span data-stu-id="32fa9-103">While you're developing your add-in, you can run the add-in on your local web server (`localhost`), but when you're ready to publish it for other users to access, you'll need to deploy the web application to a web server or web hosting service (for example, Microsoft Azure) and update the manifest to specify the URL of the deployed application.</span></span> 

<span data-ttu-id="32fa9-104">如果加载项如期工作且你已准备好发布它供其他用户访问，请完成以下步骤：</span><span class="sxs-lookup"><span data-stu-id="32fa9-104">When your add-in is working as desired and you're ready to publish it for other users to access, complete the following steps:</span></span>

1. <span data-ttu-id="32fa9-105">通过命令行，在加载项项目的根目录中运行以下命令，准备所有文件供生产部署使用：</span><span class="sxs-lookup"><span data-stu-id="32fa9-105">From the command line, in the root directory of your add-in project, run the following command to prepare all files for production deployment:</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    <span data-ttu-id="32fa9-106">生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。</span><span class="sxs-lookup"><span data-stu-id="32fa9-106">When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.</span></span>

2. <span data-ttu-id="32fa9-107">将 **dist** 文件夹的内容上传到要托管你的加载项的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="32fa9-107">Upload the contents of the **dist** folder to the web server that'll host your add-in.</span></span> <span data-ttu-id="32fa9-108">可使用任意类型的 Web 服务器或 Web 托管服务来托管加载项。</span><span class="sxs-lookup"><span data-stu-id="32fa9-108">You can use any type of web server or web hosting service to host your add-in.</span></span>

3. <span data-ttu-id="32fa9-109">在 VS Code 中，打开项目根目录中的加载项清单文件 (`manifest.xml`)。</span><span class="sxs-lookup"><span data-stu-id="32fa9-109">In VS Code, open the add-in's manifest file, located in the root directory of the project (`manifest.xml`).</span></span> <span data-ttu-id="32fa9-110">将所有出现的 `https://localhost:3000` 都替换为上一步中已部署到 Web 服务器的 Web 应用程序的 URL。</span><span class="sxs-lookup"><span data-stu-id="32fa9-110">Replace all occurrences of `https://localhost:3000` with the URL of the web application that you deployed to a web server in the previous step.</span></span>

4. <span data-ttu-id="32fa9-111">选择要用来[部署 Office 加载项](../publish/publish.md)的方法，再按照说明发布清单文件。</span><span class="sxs-lookup"><span data-stu-id="32fa9-111">Choose the method you'd like to use to [deploy and publish your Office Add-in](../publish/publish.md) your add-in, and follow the instructions to publish the manifest file.</span></span>
