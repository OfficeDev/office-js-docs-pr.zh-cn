---
title: 测试 Script Lab 集成
description: ''
ms.date: 12/04/2017
---


# <a name="testing-script-lab-integration"></a>测试 Script Lab 集成

这是一个示例测试文件，旨在演示即将推出的 ScriptLab 功能，这将使开发人员能够在 Excel、Word、PowerPoint 中试用他们的片段。  

## <a name="prerequisites"></a>先决条件
- 需要 ScriptLab 代码片段提供的查看 URL
- 注意：*应*指出 ScriptLab 需要使用 Office 365，才能探索最新代码片段。开发人员可以通过 [Office 365 开发人员计划](https://dev.office.com/devprogram)，获取仅用于开发的 Office 365 订阅。  


## <a name="try-it-out-button"></a>“试用”按钮
这样，我们将添加一个“试用”按钮，建议将其与代码片段相关联。为实现此操作，我们使用 Office UI Fabric 类将链接设置为按钮。在链接本身上，请务必设置 *aria label* 属性。

**演示：**

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">试用</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">试用</button>


**代码：**
```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>将脚本实验室作为 iframe 嵌入
在这种模式下，我们直接将片段作为 iframe 嵌入到文档中。宽度已设置为 95％（基于所有其他片段的宽度），建议删除 iframe 的 fameborder。通常应调整高度以匹片段。

**演示：**
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

**代码：**
```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>测试注意事项
我们需要验证非 Office 365 移动订阅（我们具有针对 office js docs 的反馈，其中很多开发人员使用的是 2013 的一个版本或更早版本。  

对于嵌入路径，我们需要进行最终签署，并且需要确保在查看梗概页面中显示的内容符合我们的辅助功能准则。

## <a name="see-also"></a>另请参阅
