1. <span data-ttu-id="b592c-101">在项目 (**[...] 的根目录中打开 bash 终端 Office 加载项的 /My**) 并运行以下命令以启动开发服务器。</span><span class="sxs-lookup"><span data-stu-id="b592c-101">Open a bash terminal in the root of the project and run the following command to start the dev server.</span></span>

    ```bash
    npm start
    ```

    <span data-ttu-id="b592c-102">这会启动 Web 服务器（地址为 `https://localhost:3000`），并在默认浏览器中打开此地址。</span><span class="sxs-lookup"><span data-stu-id="b592c-102">This will start a web server at `https://localhost:3000` and open your default browser to that address.</span></span>

2. <span data-ttu-id="b592c-103">Office Web 加载项应使用 HTTPS（而不是 HTTP），即使在开发时，也不例外。</span><span class="sxs-lookup"><span data-stu-id="b592c-103">Office Web Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="b592c-104">如果浏览器指明网站证书不受信任，需要将证书添加为受信任的证书。</span><span class="sxs-lookup"><span data-stu-id="b592c-104">If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate.</span></span> <span data-ttu-id="b592c-105">有关详细信息，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="b592c-105">See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b592c-106">Chrome（Web 浏览器）可能会继续指明网站证书不受信任，即使已完成[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中所述的过程，也是如此。</span><span class="sxs-lookup"><span data-stu-id="b592c-106">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span> <span data-ttu-id="b592c-107">可以忽略 Chrome 中的此警告，并转到 Internet Explorer 或 Microsoft Edge 中的 `https://localhost:3000`，以验证证书是否受信任。</span><span class="sxs-lookup"><span data-stu-id="b592c-107">You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="b592c-108">如果浏览器在加载加载项页面后没有显示任何证书错误，就可以准备测试加载项了。</span><span class="sxs-lookup"><span data-stu-id="b592c-108">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 
