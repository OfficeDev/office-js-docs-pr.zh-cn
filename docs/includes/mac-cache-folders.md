<span data-ttu-id="33c8b-p101">出于性能方面的考虑，加载项通常在 Office for Mac 中缓存。通常情况下，将通过重载加载项清除缓存。如果同一文档中存在多个加载项，则重载后自动清除缓存的过程可能不可靠。</span><span class="sxs-lookup"><span data-stu-id="33c8b-p101">Add-ins are often cached in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="33c8b-104">可以通过使用任何任务窗格加载项的个性菜单来清除缓存。</span><span class="sxs-lookup"><span data-stu-id="33c8b-104">You can clear the cache by using the personality menu of any task pane add-in.</span></span>
- <span data-ttu-id="33c8b-105">选择个性菜单。</span><span class="sxs-lookup"><span data-stu-id="33c8b-105">Choose the personality menu.</span></span> <span data-ttu-id="33c8b-106">然后选择“**清除 Web 缓存**”。</span><span class="sxs-lookup"><span data-stu-id="33c8b-106">Then choose **Clear Web Cache**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="33c8b-107">必须运行 macOS 版本 10.13.6 或更高版本才能看到个性菜单。</span><span class="sxs-lookup"><span data-stu-id="33c8b-107">You must run macOS version 10.13.6 or later to see the personality menu.</span></span>

    ![个性菜单上“清除 Web 缓存”选项的屏幕截图](../images/mac-clear-cache-menu.png)

<span data-ttu-id="33c8b-109">也可以通过删除 `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的内容来手动清除缓存。</span><span class="sxs-lookup"><span data-stu-id="33c8b-109">You can also clear the cache manually by deleting the contents of the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8b-110">如果文件夹不存在，请检查是否存在以下文件夹，如果找到，请删除文件夹的内容：</span><span class="sxs-lookup"><span data-stu-id="33c8b-110">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
>    - <span data-ttu-id="33c8b-111">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`，其中，`{host}` 是 Office 应用程序（例如 `Excel`）</span><span class="sxs-lookup"><span data-stu-id="33c8b-111">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
>    - <span data-ttu-id="33c8b-112">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`，其中，`{host}` 是 Office 应用程序（例如 `Excel`）</span><span class="sxs-lookup"><span data-stu-id="33c8b-112">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
>    - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
>    - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
