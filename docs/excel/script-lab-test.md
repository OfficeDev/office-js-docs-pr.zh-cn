---
title: 测试 Script Lab 集成
description: 此示例测试文件展示了即将推出的 ScriptLab 功能，可方便开发人员在 Excel、Word 和 PowerPoint 中试调用自己的代码片段。
ms.date: 03/14/2018
---


# <a name="testing-script-lab-integration"></a>测试 Script Lab 集成

此示例测试文件展示了即将推出的 ScriptLab 功能，可方便开发人员在 Excel、Word 和 PowerPoint 中试调用自己的代码片段。 

## <a name="prerequisites"></a>先决条件

- 需要 ScriptLab 代码片段中的视图 URL。

> [!NOTE] 
> *应*指明，ScriptLab 必须使用 Office 365 才能探索最新代码片段。 开发人员可以通过 [Office 365 开发人员计划](https://developer.microsoft.com/en-us/office/dev-program)获取 Office 365 开发人员订阅，将它仅用于开发。 请参阅 [Office 365 开发人员计划文档](https://docs.microsoft.com/zh-cn/office/developer-program/office-365-developer-program)，逐步了解如何加入 Office 365 开发人员计划并注册和配置订阅。 


## <a name="try-it-out-button"></a>“试调用”按钮

这样，我们就会添加“试调用”****按钮，建议将它与代码片段相关联。为此，使用 Office UI Fabric 类将链接设置为采用按钮样式。请务必对链接本身设置 `aria label` 属性。

### <a name="demo"></a>演示

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">试调用</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">试调用</button>


### <a name="code"></a>代码

```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>将 Script Lab 作为 iframe 嵌入

在这种模式下，将直接把代码片段作为 iframe 嵌入文档中。宽度已设置为 95%（以其他所有代码片段的宽度为依据），建议删除 iframe 的 fameborder。高度通常应调整为与代码片段一致。

### <a name="demo"></a>演示

<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

### <a name="code"></a>代码

```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>测试注意事项

需要验证非 Office 365 移动订阅（我们收到的 office-js-docs 反馈表明，其中很多开发人员使用的是 2013 或更低版本）。  

对于嵌入路径，需要进行最终签核，并确保在视图要点页面中公开的内容符合辅助功能指南。


