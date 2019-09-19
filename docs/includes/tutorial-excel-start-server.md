如果本地 web 服务器已在运行，并且您的外接程序已加载到 Excel 中，请继续执行步骤2。 否则，启动本地 web 服务器并旁加载您的外接程序： 

- 若要在 Excel 中测试外接程序，请在项目的根目录中运行以下命令。 这将启动本地 web 服务器（如果尚未运行），并在加载外接程序的情况中打开 Excel。

    ```command&nbsp;line
    npm start
    ```

- 若要在 web 上的 Excel 中测试外接程序，请在项目的根目录中运行以下命令。 运行此命令时，本地 web 服务器将启动（如果尚未运行）。

    ```command&nbsp;line
    npm run start:web
    ```

    若要使用外接程序，请在 web 上的 Excel 中打开一个新文档，然后按照旁加载中的 office[加载项旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以重新添加外接程序。
