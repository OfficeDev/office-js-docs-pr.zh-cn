使用以下过程可安装使用 Microsoft Edge 旧版 webview (EdgeHTML) 的订阅 Office 版本或使用 Internet Explorer (Trident) 的版本。

1. 在任何Office应用程序中，打开功能 **区** 上的"文件"选项卡，然后选择"Office帐户"**或**"**帐户"。** 选择" **关于 _主机名_** "按钮 (例如，"关于 **Word**) "。
1. 在打开的对话框中，找到完整的 xx.x.xxxxx.xxxxx 内部版本号，并制作它的一个副本。
1. 下载并安装[Office 部署工具](https://www.microsoft.com/download/details.aspx?id=49117)。
1. 在安装该工具的文件夹中 (文件所在的) ，创建一个包含名称的文本文件并 `setup.exe` `config.xml` 添加以下内容。

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. 更改 `Version` 值。

    - 若要安装使用旧版 Edge 的版本，请将其更改为 `16.0.11929.20946` 。
    - 若要安装使用 Internet Explorer，请将其更改为 `16.0.10730.20348` 。

1. （可选）将 的值更改为 以安装 32 Office，并根据需要更改值以 `OfficeClientEdition` `"32"` `Language ID` Office其他语言安装客户端。
1. 以管理员角色 *打开命令提示符*。
1. 导航到包含 和 `setup.exe` 文件 `config.xml` 的文件夹。
1. 运行以下命令。

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    此命令将安装Office。 此过程可能需要几分钟时间。

1. [清除Office缓存](../testing/clear-cache.md)。

> [!IMPORTANT]
> 安装后，请确保关闭 Office 的自动更新，以便 Office 不会更新为在使用完 Web 视图之前不使用的 Web 视图的版本。 **这可在安装后数分钟内发生。** 请按照以下步骤操作。
>
> 1. 启动任何Office应用程序并打开一个新文档。
> 1. 打开功能 **区上的**"文件"选项卡，然后选择"Office **帐户**"或"**帐户"。**
> 1. 在"**产品信息"** 列中，选择"**更新选项**"，然后选择"**禁用更新"。** 如果该选项不可用，则Office配置为不自动更新。

使用完旧版本的 Office后，请通过编辑文件，将 更改为之前复制的版本号来重新安装 `config.xml` `Version` 较新版本。 然后在管理员 `setup.exe /configure config.xml` 命令提示符中重复该命令。 （可选）重新启用自动更新。
