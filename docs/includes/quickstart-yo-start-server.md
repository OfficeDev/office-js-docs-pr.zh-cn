1. 在项目 (**[...] 的根目录中打开 bash 终端 Office 加载项的 /My**) 并运行以下命令以启动开发服务器。

    ```bash
    npm start
    ```

    这会启动 Web 服务器（地址为 `https://localhost:3000`），并在默认浏览器中打开此地址。

2. Office Web 加载项应使用 HTTPS（而不是 HTTP），即使在开发时，也不例外。 如果浏览器指明网站证书不受信任，需要将证书添加为受信任的证书。 有关详细信息，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

    > [!NOTE]
    > Chrome（Web 浏览器）可能会继续指明网站证书不受信任，即使已完成[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中所述的过程，也是如此。 可以忽略 Chrome 中的此警告，并转到 Internet Explorer 或 Microsoft Edge 中的 `https://localhost:3000`，以验证证书是否受信任。 

3. 如果浏览器在加载加载项页面后没有显示任何证书错误，就可以准备测试加载项了。 
