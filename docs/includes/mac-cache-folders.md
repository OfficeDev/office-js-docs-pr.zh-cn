由于性能方面的原因, 外接程序通常缓存在 Office for Mac 中。 通常情况下，将通过重载外接程序清除缓存。 如果同一个文档中存在多个外接程序, 则在重新加载时自动清除缓存的过程可能不可靠。

您可以使用任何任务窗格外接程序的 "个性" 菜单清除缓存。
- 选择 "个性" 菜单。 然后选择 "**清除 Web 缓存**"。
    > [!NOTE]
    > 您必须运行 macOS 版本10.13.6 或更高版本, 才能看到 "个性" 菜单。
    
    !["个性" 菜单上的 "清除 web 缓存" 选项的屏幕截图。](../images/mac-clear-cache-menu.png)

您还可以通过删除该`~/Library/Containers/com.Microsoft.OsfWebHost/Data/`文件夹的内容来手动清除缓存。

> [!NOTE]
> 如果该文件夹不存在, 请检查以下文件夹, 如果找到, 则删除该文件夹的内容:
>    - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`其中`{host}` , 是 Office 主机 (例如, `Excel`)
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
