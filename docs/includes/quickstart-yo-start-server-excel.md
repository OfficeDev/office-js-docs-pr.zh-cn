
<span data-ttu-id="49d44-101">完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="49d44-101">Complete the following steps to start the local web server and sideload your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="49d44-102">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="49d44-102">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="49d44-103">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="49d44-103">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

> [!TIP]
> <span data-ttu-id="49d44-104">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="49d44-104">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="49d44-105">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="49d44-105">When you run this command, the local web server will start.</span></span>
>
> ```command&nbsp;line
> npm run dev-server
> ```

- <span data-ttu-id="49d44-106">若要在 Excel 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="49d44-106">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="49d44-107">运行此命令时，本地 Web 服务器将启动（如尚未运行），Excel 将打开且加载项已载入。</span><span class="sxs-lookup"><span data-stu-id="49d44-107">When you run this command, the local web server will start and Word will open with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="49d44-108">若要在 Excel Online 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="49d44-108">To test your add-in in Excel Online, run the following command in the root directory of your project.</span></span> <span data-ttu-id="49d44-109">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="49d44-109">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="49d44-110">若要使用加载项，请在 Excel Online 中打开新的工作簿，然后按照[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)中的说明旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="49d44-110">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).</span></span>

