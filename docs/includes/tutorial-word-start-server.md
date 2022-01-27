如果本地 Web 服务器已在运行，并且加载项已加载到 Word 中，则继续执行步骤 2。 否则，启动本地 Web 服务器并旁加载你的加载项： 

- 若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。 这将启动本地的 Web 服务器（如果尚未运行的话），并使用加载的加载项打开 Word。

    ```command&nbsp;line
    npm start
    ```

- 若要在 Word 网页版中测试加载项，请在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动。 将“{url}”替换为你有权访问的 OneDrive 或 SharePoint 库中 Word 文档的 URL。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

