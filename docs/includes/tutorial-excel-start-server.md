<span data-ttu-id="655ff-101">如果本地 web 服务器已在运行，并且您的外接程序已加载到 Excel 中，请继续执行步骤2。</span><span class="sxs-lookup"><span data-stu-id="655ff-101">If the local web server is already running and your add-in is already loaded in Excel, proceed to step 2.</span></span> <span data-ttu-id="655ff-102">否则，启动本地 web 服务器并旁加载您的外接程序：</span><span class="sxs-lookup"><span data-stu-id="655ff-102">Otherwise, start the local web server and sideload your add-in:</span></span> 

- <span data-ttu-id="655ff-103">若要在 Excel 中测试外接程序，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="655ff-103">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="655ff-104">这将启动本地 web 服务器（如果尚未运行），并在加载外接程序的情况中打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="655ff-104">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="655ff-105">若要在 web 上的 Excel 中测试外接程序，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="655ff-105">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="655ff-106">运行此命令时，本地 web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="655ff-106">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="655ff-107">若要使用外接程序，请在 web 上的 Excel 中打开一个新文档，然后按照旁加载中的 office[加载项旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以重新添加外接程序。</span><span class="sxs-lookup"><span data-stu-id="655ff-107">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>
