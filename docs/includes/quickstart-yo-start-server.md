1. 打开项目根目录 (**[...]/My Office Add-in**) 中的 bash 终端，并运行下面的命令以启动开发服务器。

    ```bash
    npm start
    ```

2. 打开 Internet Explorer 或 Microsoft Edge 并导航至 `https://localhost:3000`。 如果页面正常加载没有任何证书错误，请继续本文的下一节（**试用**）。 如果浏览器指示网站的证书不受信任，请继续执行以下步骤。

3. Office Web 加载项应使用 HTTPS（而不是 HTTP），即使在开发时，也不例外。 如果浏览器指明网站证书不受信任，需要将证书添加为受信任的证书。 有关详细信息，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

    > [!NOTE]
    > Chrome（Web 浏览器）可能会继续指明网站证书不受信任，即使已完成[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中所述的过程，也是如此。 因此，应使用 Internet Explorer 或 Microsoft Edge 验证此证书是否受信任。 

4. 如果浏览器在加载加载项页面后没有显示任何证书错误，就可以准备测试加载项了。
