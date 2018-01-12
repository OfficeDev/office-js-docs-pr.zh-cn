
# <a name="sideload-office-add-ins-for-testing"></a>旁加载 Office 外接程序以进行测试

你可以安装 Office 外接程序以在 Windows 上运行的 Office 客户端中进行测试（通过使用共享文件夹，以将清单发布到网络文件共享）。 

如果不在 Windows 上测试 Word、Excel 或 PowerPoint 外接程序，则请参阅以下主题之一来旁加载外接程序：

- [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [在 iPad 和 Mac 上旁加载 Office 外接程序进行测试](sideload-an-office-add-in-on-ipad-and-mac.md)
- [旁加载 Outlook 外接程序进行测试](sideload-outlook-add-ins-for-testing.md)

下面的视频演示将外接程序旁加载到 Office 桌面或 Office Online 上的流程。

<iframe width="560" height="315" src="https://www.youtube.com/embed/XXsAw2UUiQo" frameborder="0" allowfullscreen></iframe>


## <a name="share-a-folder"></a>共享文件夹

1. 在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。

2. 打开（右键单击）文件夹的上下文菜单并选择“**属性**”。

3. 打开“**共享**”选项卡。

4. 在“**选择人员...**”页上，添加你自己以及想要与其共享外接程序的其他任何人。 如果他们都是安全组的成员，那么你可以添加该组。 你将至少需要该文件夹的**读/写**权限。 

5. 依次选择“**共享**”、“ > **完成**”和“ > **关闭**”。

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>将共享文件夹指定为受信任的目录

      
3. 在 Excel、Word 或 PowerPoint 中打开一个新文档。
    
4. 选择“**文件**”选项卡，然后选择“**选项**”。
    
5. 选择“**信任中心**”，然后选择“**信任中心设置**”按钮。
    
6. 选择“**受信任的外接程序目录**”。
    
7. 在“**目录 URL**”框中，输入共享文件夹目录的完整网络路径，然后选择“**添加目录**”。
    
8. 选中“**显示在菜单中**”复选框，然后选择“**确定**”。

9. 关闭 Office 应用程序，你的更改将生效。
    
## <a name="sideload-your-add-in"></a>旁加载外接程序


1. 放入在共享文件夹目录中进行测试的所有外接程序的清单文件。 请务必将 Web 应用程序本身部署到 Web 服务器。 务必在清单文件的 **SourceLocation** 元素中指定 URL。

    >**重要说明：**若要帮助确保访问外部数据和服务的外接程序更加安全，外接程序应使用一个安全协议（例如 HTTPS）连接到外部数据和服务。 如果外接程序使用外接程序命令，则必须使用 HTTPS。

2. 在 Excel、Word 或 PowerPoint 中，在功能区的“**插入**”选项卡上选择“**我的外接程序**”。

3. 在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。

4. 选择外接程序的名称并选择“**确定**”以插入外接程序。


## <a name="additional-resources"></a>其他资源

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)
    
