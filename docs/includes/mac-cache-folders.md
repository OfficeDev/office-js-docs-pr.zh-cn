出于性能原因，加载项通常缓存在 mac Office中。 通常情况下，将通过重新加载加载项清除缓存。 如果同一文档中存在多个加载项，则重载后自动清除缓存的过程可能不可靠。

可以通过使用任何任务窗格加载项的个性菜单来清除缓存。

- 选择个性菜单。 然后选择“**清除 Web 缓存**”。
    > [!NOTE]
    > 必须运行 macOS 版本 10.13.6 或更高版本才能看到个性菜单。

    ![个性菜单上“清除 Web 缓存”选项的屏幕截图](../images/mac-clear-cache-menu.png)

也可以通过删除 `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的内容来手动清除缓存。

> [!NOTE]
> 如果文件夹不存在，请检查是否存在以下文件夹，如果找到，请删除文件夹的内容。
>
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`，其中，`{host}` 是 Office 应用程序（例如 `Excel`）
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`，其中，`{host}` 是 Office 应用程序（例如 `Excel`）
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
