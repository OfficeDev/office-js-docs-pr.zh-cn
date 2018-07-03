# <a name="navigation-patterns"></a>导航模式

加载项的主要功能通过特定的命令类型和有限的屏幕区域进行访问。 导航很直观，提供上下文并允许用户在整个加载项中轻松移动，这一点很重要。

## <a name="best-practices"></a>最佳做法

| 允许事项    | 禁止事项 |
| :---- | :---- |
| 确保用户有一个清晰可见的导航选项。 | 不要使用非标准的 UI 来使导航流程复杂化。
| 根据适用情况，使用以下组件来允许用户浏览加载项。 | 不要让用户难以理解他们在加载项中的当前位置或上下文



## <a name="command-bar"></a>命令栏

CommandBar 是一个表面，其中包含对上面所在的窗口、面板或父区域的内容进行操作的命令。 可选功能包括汉堡菜单访问点、搜索和侧面命令。

![命令 - 桌面任务窗格的规范](../images/add-in-command-bar.png)



## <a name="tab-bar"></a>选项卡栏

data-id="undefined" class="unusedGlossaryTerm">选项卡栏 使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。

![选项卡栏 - 桌面任务窗格的规范](../images/add-in-tab-bar.png)


## <a name="back-button"></a>后退按钮

后退按钮允许用户从深化导航操作中恢复。 这个模式有助于确保用户遵循一系列有序的步骤。  

![后退按钮 - 桌面任务窗格的规范](../images/add-in-back-button.png)
