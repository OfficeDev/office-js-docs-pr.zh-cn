如果本地 web 服务器已在运行，并且你的外接程序已在 Word 中加载，请继续执行步骤2。 否则，启动本地 web 服务器并旁加载您的外接程序： 

- 若要在 Word 中测试您的外接程序，请在项目的根目录中运行以下命令。 这将启动本地 web 服务器（如果尚未运行），并在加载的外接程序中打开 Word。

    ```command&nbsp;line
    npm start
    ```

- 若要在 web 上的 Word 中测试您的外接程序，请在项目的根目录中运行以下命令。 运行此命令时，本地 web 服务器将启动（如果尚未运行）。

    ```command&nbsp;line
    npm run start:web
    ```

    若要使用外接程序，请在 web 上的 Word 中打开一个新文档，然后按照旁加载 Office 加载项旁加载中的说明操作，以在[web 上的 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中进行添加。
