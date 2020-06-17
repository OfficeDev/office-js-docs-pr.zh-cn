<span data-ttu-id="84861-101">由于性能方面的原因，外接程序通常缓存在 Office for Mac 中。</span><span class="sxs-lookup"><span data-stu-id="84861-101">Add-ins are often cached in Office for Mac, for performance reasons.</span></span> <span data-ttu-id="84861-102">通常情况下，将通过重载外接程序清除缓存。</span><span class="sxs-lookup"><span data-stu-id="84861-102">Normally, the cache is cleared by reloading the add-in.</span></span> <span data-ttu-id="84861-103">如果同一个文档中存在多个外接程序，则在重新加载时自动清除缓存的过程可能不可靠。</span><span class="sxs-lookup"><span data-stu-id="84861-103">If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="84861-104">您可以使用任何任务窗格外接程序的 "个性" 菜单清除缓存。</span><span class="sxs-lookup"><span data-stu-id="84861-104">You can clear the cache by using the personality menu of any task pane add-in.</span></span>
- <span data-ttu-id="84861-105">选择 "个性" 菜单。</span><span class="sxs-lookup"><span data-stu-id="84861-105">Choose the personality menu.</span></span> <span data-ttu-id="84861-106">然后选择 "**清除 Web 缓存**"。</span><span class="sxs-lookup"><span data-stu-id="84861-106">Then choose **Clear Web Cache**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="84861-107">您必须运行 macOS 版本10.13.6 或更高版本，才能看到 "个性" 菜单。</span><span class="sxs-lookup"><span data-stu-id="84861-107">You must run macOS version 10.13.6 or later to see the personality menu.</span></span>
    
    !["个性" 菜单上的 "清除 web 缓存" 选项的屏幕截图。](../images/mac-clear-cache-menu.png)

<span data-ttu-id="84861-109">您还可以通过删除该文件夹的内容来手动清除缓存 `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` 。</span><span class="sxs-lookup"><span data-stu-id="84861-109">You can also clear the cache manually by deleting the contents of the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

> [!NOTE]
> <span data-ttu-id="84861-110">如果文件夹不存在，查看下列文件夹，如果找到，删除文件夹的内容：</span><span class="sxs-lookup"><span data-stu-id="84861-110">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
>    - <span data-ttu-id="84861-111">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`此位置`{host}`是 Office 主机（如 `Excel` ）</span><span class="sxs-lookup"><span data-stu-id="84861-111">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
>    - <span data-ttu-id="84861-112">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`此位置`{host}`是 Office 主机（如 `Excel` ）</span><span class="sxs-lookup"><span data-stu-id="84861-112">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
>    - `com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
