在本教程中，请先设置开发项目。 

> [!NOTE]
> 此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。

> [!TIP]
> 如果尚未这样做，请参阅[生成首个 Word 加载项](../quickstarts/word-quickstart.md?tabs=visual-studio-code)。 特别是，请务必了解如何旁加载 Word 加载项以供测试。

## <a name="prerequisites"></a>先决条件

若要学习本教程，需要安装以下各项。 

- Word 2016 版本 1711（生成号 8730.1000 即点即用）或更高版本。 可能必须成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。
- [Node 和 npm](https://nodejs.org/en/) 
- [Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）

## <a name="setup"></a>设置

1. 克隆 GitHub 存储库 [Word 加载项教程](https://github.com/OfficeDev/Word-Add-in-Tutorial)。
2. 打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。
3. 运行命令 `npm install`，以安装 package.json 文件中列出的工具和库。 
4. 按照[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中的步骤操作，信任开发计算机操作系统的证书。

