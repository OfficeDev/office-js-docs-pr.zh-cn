如果项目是基于node.js的 (（不是使用 Visual Studio 和 Internet Information server (IIS) ) 开发的），则可以强制 Windows 上的 Office 使用 Edge 旧版或 Internet Explorer 运行加载项，即使你拥有通常使用最新浏览器的 Windows 和 Office 版本的组合。 有关 Windows 和 Office 版本的各种组合使用哪些浏览器的详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!NOTE]
> 仅 Microsoft 365 的 Beta 订阅通道支持用于强制浏览器更改的工具。 加入 [Office 预览体验计划](https://insider.office.com/join/windows) 并选择 **Beta 频道** 选项以访问 Office Beta 生成。 另请参阅[关于 Office：我使用的是哪个版本的 Office？](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
>
> 严格地说，此 `webview` 工具的开关 (请参阅需要 Beta 通道的 **步骤 2**) 。 该工具具有其他不具有此要求的开关。

1. 如果项目 *不是* 使用 [适用于 Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 工具创建的，则需要安装 office-addin-dev-settings 工具。 在命令提示符下运行以下命令。

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. 在项目根目录的命令提示符中指定希望 Office 与以下命令一起使用的浏览器。 替换 `<path-to-manifest>` 为相对路径，如果该路径位于项目的根目录中，则该路径只是清单文件名。 替换 `<webview>` 为或 `ie` `edge-legacy`.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    示例如下。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    命令行中应会显示一条消息，指出 Webview 类型现在设置为 IE (或 Edge 旧版) 。

1. 完成后，使用默认浏览器将 Office 设置为继续使用以下命令组合 Windows 和 Office 版本。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
