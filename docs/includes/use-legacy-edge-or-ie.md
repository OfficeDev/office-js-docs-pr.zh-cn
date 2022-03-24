如果您的项目基于 node.js 的 (，即不是使用 Visual Studio 和 Internet 信息服务器 (IIS) ) 开发的，您可以强制 Office Windows 在 Windows 上使用 Edge Legacy 或 Internet Explorer 运行外接程序，即使您具有通常使用较新浏览器的 Windows 和 Office 版本的组合。 有关各种浏览器组合使用哪些浏览器Windows Office，请参阅 Office [外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

1. 如果你的项目 *不是使用* Office 加载项工具的 [Yeoman](../develop/yeoman-generator-overview.md) 生成器创建的，则需要安装 office-addin-dev-settings 工具。 在命令提示符中运行以下命令。

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. 在项目根目录Office命令提示符中指定要与以下命令一同使用的浏览器。 将 `<path-to-manifest>` 替换为相对路径，如果清单文件名位于项目的根目录，则只是清单文件名。 将 `<webview>` 替换为 或 `ie` `edge-legacy`。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    示例如下。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    您应该在命令行中看到一条消息，表明 Webview 类型现在设置为 IE (或 Edge 旧版) 。

1. 完成后，通过以下Office，将 Windows 和 Office 版本组合使用默认浏览器恢复。

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
