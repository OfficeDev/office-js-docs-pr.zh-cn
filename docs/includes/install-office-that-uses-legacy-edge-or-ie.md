使用以下过程安装使用 Microsoft Edge 旧版 webview (EdgeHTML) 运行加载项的订阅Office版本，或者安装使用 Internet Explorer (Trident) 的版本。

1. 在任何Office应用程序中，打开功能区上的 **“文件”** 选项卡，然后选择 **Office帐户** 或 **帐户**。 选择 **“关于 _主机名_** ”按钮 (例如“ **关于 Word**) ”。
1. 在打开的对话框中，找到完整的 xx.x.xxxxx.xxxxx 内部版本号，并将其复制到某处。
1. 下载 [Office 部署工具](https://www.microsoft.com/download/details.aspx?id=49117)。
1. 运行下载的文件以提取该工具。 系统会提示你选择在何处安装该工具。
1. 在安装该工具的文件夹中， (`setup.exe` 文件所在的) ，创建包含该名称 `config.xml` 的文本文件并添加以下内容。

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. 更改值 `Version` 。

    - 若要安装使用 Edge Legacy 的版本，请将其更改为 `16.0.11929.20946`。
    - 若要安装使用 Internet Explorer 的版本，请将其更改为 `16.0.10730.20348`。

1. （可选）更改为安装 32 位Office的值`OfficeClientEdition``"32"`，并根据需要更改`Language ID`以其他语言安装Office的值。
1. *以管理员身份* 打开命令提示符。
1. 使用文件导航到文件夹`setup.exe``config.xml`。
1. 运行以下命令：

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    此命令安装Office。 这个过程可能需要几分钟。

1. [清除Office缓存](../testing/clear-cache.md)。

> [!IMPORTANT]
> 安装后，请确保关闭自动更新Office，以便Office不会更新为在使用完之前不使用要使用的 Web 视图的版本。 **这可能在安装后几分钟内发生。** 请按照以下步骤操作。
>
> 1. "开始"菜单任何Office应用程序并打开新文档。
> 1. 打开功能区上的 **“文件”** 选项卡，然后选择 **Office帐户** 或 **帐户**。
> 1. 在 **“产品信息** ”列中，选择 **“更新选项**”，然后选择 **“禁用更新**”。 如果该选项不可用，则Office已配置为不自动更新。

使用完旧版本的Office后，通过编辑`config.xml`文件并更改`Version`为之前复制的生成号来重新安装较新版本。 然后在 `setup.exe /configure config.xml` 管理员命令提示符中重复该命令。 （可选）重新启用自动更新。
