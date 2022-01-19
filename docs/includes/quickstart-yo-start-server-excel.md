
完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。

[!INCLUDE [alert use https](alert-use-https.md)]

> [!TIP]
> 如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。 运行此命令时，本地 Web 服务器将启动。
>
> ```command&nbsp;line
> npm run dev-server
> ```

- 若要在 Excel 中测试加载项，请在项目的根目录中运行以下命令。 这将启动本地的 Web 服务器, 并使用加载的加载项打开 Excel。

    ```command&nbsp;line
    npm start
    ```

- 若要在浏览器版 Excel 中测试加载项，请在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动。 将 "{url}" 替换为你拥有权限的 OneDrive 或 SharePoint 库上 Excel 文档的 URL。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]
