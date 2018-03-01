---
title: 旁加载 Office 加载项以供测试
description: ''
ms.date: 01/25/2018
---

# <a name="sideload-office-add-ins-for-testing"></a>旁加载 Office 加载项以供测试

你可以安装 Office 外接程序以在 Windows 上运行的 Office 客户端中进行测试（通过使用共享文件夹，以将清单发布到网络文件共享）。 

如果不在 Windows 上测试 Word、Excel 或 PowerPoint 外接程序，则请参阅以下主题之一来旁加载外接程序：

- [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [在 iPad 和 Mac 上旁加载 Office 加载项以供测试](sideload-an-office-add-in-on-ipad-and-mac.md)

下面的视频逐步展示了如何在 Office 桌面或 Office Online 上旁加载加载项。  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>共享文件夹

1. 在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。

2. 打开（右键单击）文件夹的上下文菜单并选择“**属性**”。

3. 打开“**共享**”选项卡。

4. 在“**选择人员...**”页上，添加你自己以及想要与其共享外接程序的其他任何人。如果他们都是安全组的成员，那么可以添加该组。将至少需要该文件夹的**读/写**权限。 

5. 依次选择“**共享**”、“ > **完成**”和“ > **关闭**”。


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>将共享文件夹指定为受信任的目录
      
1. 在 Excel、Word 或 PowerPoint 中打开一个新文档。
    
2. 选择“**文件**”选项卡，然后选择“**选项**”。
    
3. 选择“**信任中心**”，然后选择“**信任中心设置**”按钮。
    
4. 选择“**受信任的外接程序目录**”。
    
5. 在“**目录 URL**”框中，输入共享文件夹目录的完整网络路径，然后选择“**添加目录**”。
    
6. 选中“**显示在菜单中**”复选框，然后选择“**确定**”。

7. 关闭 Office 应用程序，你的更改将生效。
    

## <a name="sideload-your-add-in"></a>旁加载外接程序

1. 放入在共享文件夹目录中进行测试的所有外接程序的清单文件。请务必将 Web 应用程序本身部署到 Web 服务器。务必在清单文件的 **SourceLocation** 元素中指定 URL。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. 在 Excel、Word 或 PowerPoint 中，选择功能区上“插入”****选项卡中的“我的加载项”****。

3. 在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。

4. 依次选择加载项名称和“确定”****，以插入加载项。


## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)
    
