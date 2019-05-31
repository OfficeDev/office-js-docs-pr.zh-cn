
完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。

> [!NOTE]
> Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

> [!TIP]
> 如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。 运行此命令时，本地 Web 服务器将启动。
>
> ```command&nbsp;line
> npm run dev-server
> ```

- 若要在 Excel 中测试加载项，请在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动（如尚未运行），Excel 将打开且加载项已载入。

    ```command&nbsp;line
    npm start
    ```

- 若要在 Excel Online 中测试加载项，请在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动（如果尚未运行）。

    ```command&nbsp;line
    npm run start:web
    ```

    若要使用加载项，请在 Excel Online 中打开新的工作簿，然后按照[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)中的说明旁加载你的加载项。

