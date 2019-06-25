
完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。

> [!NOTE]
> Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

> [!TIP]
> 如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。 运行此命令时，本地 Web 服务器将启动。
>
> ```command&nbsp;line
> npm run dev-server
> ```

- 若要在 Excel 中测试加载项，请在项目的根目录中运行以下命令。 如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话），而且 Excel 也将打开并载入加载项。

    ```command&nbsp;line
    npm start
    ```

- 若要在浏览器版 Excel 中测试加载项，请在项目的根目录中运行以下命令。 如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。

    ```command&nbsp;line
    npm run start:web
    ```

    若要使用加载项，请在 Excel 网页版中打开新的工作簿，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。

