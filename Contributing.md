# <a name="contribute-to-this-documentation"></a>参与此文档

感谢你对文档的关注！

* [参与方法](#ways-to-contribute)
* [通过 GitHub 参与](#contribute-using-github)
* [通过 Git 参与](#contribute-using-git)
* [如何使用 Markdown 设置主题格式](#how-to-use-markdown-to-format-your-topic)
* [常见问题解答](#faq)
* [更多资源](#more-resources)

## <a name="ways-to-contribute"></a>参与方法

以下是参与此文档的一些方法：

* 进行小幅度的更改，[通过 GitHub 参与](#contribute-using-github)。
* 进行大幅度的更改，或更改涉及代码，[通过 Git 参与](#contribute-using-git)。
* 通过“GitHub 问题”举报文档缺陷。
* 通过 [Office 开发人员平台 UserVoice](http://officespdev.uservoice.com) 网站请求查看新文档。

## <a name="contribute-using-github"></a>通过 GitHub 参与

通过 GitHub 参与此文档无需将报告复制到桌面。这是在存储库中创建拉取请求的最简单的方法。使用此方法进行不涉及代码更改的少量改动。 

**注意**：此方法允许一次参与一篇文章。

### <a name="to-contribute-using-github"></a>通过 GitHub 参与

1. 在 GitHub 上找到想要参与的文章。
2. 进入 GitHub 上的相应文章后，登录 GitHub（获取免费帐户[加入 GitHub](https://github.com/join)）。
3. 选择“**铅笔图标**”（编辑项目分叉中的对应文件），然后在“**<>编辑文件**”窗口进行更改。 
4. 滚动到底部，输入说明。
5. 依次选择“**建议文件更改**”>“**创建拉取请求**”。

现已成功提交拉取请求。拉取请求通常在 10 个工作日内完成审核。 


## <a name="contribute-using-git"></a>通过 Git 参与

通过 Git 参与提供重大更改，例如：

* 提供代码。
* 提供影响主旨的更改。
* 提供对文本的大量更改。
* 添加新主题。

### <a name="to-contribute-using-git"></a>通过 Git 参与

1. 如果没有 GitHub 帐户，可通过 [GitHub](https://github.com/join) 建立帐户。 
2. 拥有帐户后，在计算机上安装 Git。 按[安装 Git] 教程中的步骤操作。
3. 若要通过 Git 提交拉取请求，遵循[使用 GitHub、Git 和此存储库](#use-github-git-and-this-repository)中的步骤操作。
4. 以下人员需签署参与者许可协议：

    * Microsoft Open Technologies 组的成员。
    * 不在 Microsoft 供职的参与者。

社区成员必须签署参与许可协议 (CLA) 才能向项目提供大量提交内容。只需完成并提交此文档一次。请仔细查看该文档。可能要求你让你的员工签署此文档。

签署此 CLA 并未授予你操纵主存储库的权利，但确实表示 Office 开发人员和 Office 开发人员内容发布团队可以查看并批准你提供的内容。 针对提供的内容，将获得相应奖励。

拉取请求通常在 10 个工作日内完成审核。

## <a name="use-github-git-and-this-repository"></a>使用 GitHub、Git 和此存储库

**注意**：此部分中的大多数信息都可以在 [GitHub 帮助]文章中找到。  如果熟悉 Git 和 GitHub，请跳至**参与和编辑内容**部分，了解此存储库的代码/内容流的具体详情。

### <a name="to-set-up-your-fork-of-the-repository"></a>设置存储库分支的具体步骤

1.  建立 GitHub 帐户以参与此项目。如果还未进行此操作，请转至 [GitHub](https://github.com/join) 立即进行。
2.  在计算机上安装 Git。 按[安装 Git] 教程中的步骤操作。
3.  对此存储库创建你自己的分叉。为此，在页面顶部，选择“分叉”**** 按钮。
4.  将自己的分叉复制到计算机上。打开 Git Bash 以完成此步骤。在命令提示符中输入：

        git clone https://github.com/<your user name>/<repo name>.git

    然后，通过输入以下命令来创建对根库的引用：

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

恭喜你！现已设置好存储库。无需再次重复上述步骤。

### <a name="contribute-and-edit-content"></a>参与和编辑内容

若要使参与过程无缝进行，请遵循以下步骤。

#### <a name="to-contribute-and-edit-content"></a>参与和编辑内容

1. 新建分支。
2. 添加新内容或编辑现有内容。
3. 向主存储库提交拉取请求。
4. 删除分支。

**重要说明**每个分支限制为单个概念/文章以简化工作流并降低合并冲突的机率。适用于新分支的内容包括：

* 新文章。
* 拼写和语法编辑。
* 对大量文章应用单个格式更改（例如，应用新的版权页脚）。

#### <a name="to-create-a-new-branch"></a>新建分支

1.  打开 Git Bash。
2.  在 Git Bash 命令提示符中键入 `git pull upstream master:<new branch name>`。此操作将最新的 OfficeDev 母版分支中复制为本地新建分支。
3.  在 Git Bash 命令提示符中键入 `git push origin <new branch name>`。此操作将对 GitHub 提示该新分支。现在将可以在 GitHub 的存储库分叉上看到新分支。
4.  在 Git Bash 命令提示符中键入 `git checkout <new branch name>` 以转至新分支。

#### <a name="add-new-content-or-edit-existing-content"></a>添加新内容或编辑现有内容

通过使用文件资源管理器导航至计算机上的存储库。存储库文件位于 `C:\Users\<yourusername>\<repo name>`。

若要编辑文件，请在自己选择的编辑器中将其打开并进行修改。若要新建文件，请使用自己选择的编辑器并将新文件存储在本地存储库副本中的适当位置。执行操作期间，经常保存进行的操作。

`C:\Users\<yourusername>\<repo name>` 中的文件是你在本地存储库中创建的新分支的工作副本。提交更改前，在此文件夹的任何更改都不会影响本地存储库。若要向本地存储库提交更改，在 GitBash 中键入以下命令：

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

`add` 命令将更改添加到临时区域以准备将其提交到存储库。`add` 命令表明你希望将所有已添加或修改的文件暂存后，对子文件夹进行递归查看。（若不希望提交所有更改，可以添加特定的文件）。也可以撤消提交。要请求帮助，请键入 `git add -help` 或 `git status`。）

`commit` 命令将临时更改应用到存储库。`-m` 开关表示你正在向命令行提供要提交的注释。-v 和 -a 开关可以忽略。-v 开关用于命令的详细输出，-a 执行添加命令已执行的操作。

参与期间可以多次提交，也可结束参与后一次性提交。

#### <a name="submit-a-pull-request-to-the-main-repository"></a>向主存储库提交拉取请求

结束参与准备将其合并到主存储库时，请遵循以下步骤。

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>向主存储库提交拉取请求

1.  在 Git Bash 命令提示符中键入 `git push origin <new branch name>`。在本地存储库中，`origin` 代表从其中复制本地存储库的 GitHub 存储库。此命令将新分支的当前状态（包括上述步骤中的所有提交）推送到 GitHub 分叉。
2.  在 GitHub 网站上，从自己的分叉中导航到新分支。
3.  选择页面顶部的“**拉取请求**”按钮。
4.  验证基本分支是否是 `OfficeDev/<repo name>@master` 且头分支是否是 `<your username>/<repo name>@<branch name>`。
5.  选择“**更新提交范围**”按钮。
6.  向拉取请求添加标题，然后说明进行的所有更改。
7.  提交拉取请求。

其中一个网站管理员将处理你的拉取请求。你的拉取请求将出现在 OfficeDev/<repo name> 网站中的“问题”部分。接受拉取请求后，将解决此问题。

#### <a name="create-a-new-branch-after-merge"></a>合并后新建分支

分支成功合并后（即已接受拉取请求），不能继续在该本地分支中操作。这会导致提交其他拉取请求时出现合并冲突。若要进行其他更新，从已成功合并的上游分支中新建本地分支，然后删除最初的本地分支。

例如，如果你的本地分支 X 已成功合并到 OfficeDev/microsoft-graph-docs 母版分支，而你希望对已合并的内容进行其他更新。从 OfficeDev/microsoft-graph-docs 母版分支中新建本地分支 X2。若要实现此操作，请打开 GitBash，然后执行以下命令：

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

现在新本地分支中已具备在分支 X 中提交的内容的本地副本。X2 分支还包含其他作者已合并的所有参与内容，因此如果你的参与内容以其他人的参与内容为基础（例如，共享的图像），也可在新分支中获取该内容。你可以通过查看新分支来验证之前的参与内容（及其他人的参与内容）在此分支中...

    git checkout X2

…是否包含在此分支中。（`checkout` 命令将 `C:\Users\<yourusername>\microsoft-graph-docs` 中的文件更新为 X2 分支的当前状态。）查看新分支后，可以对内容进行更新并以常用方式将其提交。但是为了避免误在已合并的分支 (X) 进行操作，最好将其删除（请参阅下列“**删除分支**”部分）。

#### <a name="delete-a-branch"></a>删除分支

成功将更改并入主存储库后，将已使用的分支删除，因为将无需再使用该分支了。其他任何操作都应在新分支中完成。  

#### <a name="to-delete-a-branch"></a>删除分支

1.  在 Git Bash 命令提示符中键入 `git checkout master`。这将确保你不在即将删除的分支中（不允许在即将删除的分支中）。
2.  然后，在命令提示符中键入 `git branch -d <branch name>`。只要分支已成功合并到上流存储库，此操作即会在计算机上将其删除。（可以使用 `–D` 标记替代此行为，但首先请确认你要这么做。）
3.  最后，在命令提示符中键入 `git push origin :<branch name>`（冒号前有一个空格，之后没有空格）。这将删除 github 分叉中的分支。  

恭喜！你已成功参与此项目！

## <a name="how-to-use-markdown-to-format-your-topic"></a>如何使用 Markdown 设置主题格式

### <a name="markdown"></a>Markdown

此存储库中的所有文章都使用 Markdown。 [Daring Fireball - Markdown] 中有完整介绍（并列出了所有语法）。
 
## <a name="faq"></a>FAQ

### <a name="how-do-i-get-a-github-account"></a>如何获取 GitHub 帐户？

填写[加入 GitHub](https://github.com/join) 中的表格，以开立免费 GitHub 帐户。 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>从何处获取参与者许可协议？ 

拉取请求请求获取此协议时，将会自动向你发送需要签署参与者许可协议 (CLA) 的通知。 

社区成员**必须签署参与许可协议 (CLA) 才能向项目提供大量提交内容**。只需完成并提交此文档一次。请仔细查看该文档。可能要求你让你的员工签署此文档。

### <a name="what-happens-with-my-contributions"></a>会如何处理我的参与内容？

提交更改后，会通过拉取请求通知我们的团队，然后会对你的拉取请求进行审核。你将收到 GitHub 对你的拉取请求的相关通知；我们需要进一步信息时，也可能收到我们团队人员的通知。如果你的拉取请求得到批准，我们将更新文档。我们保留出于合法性、风格、简洁性或其他的问题对你的提交内容进行编辑的权利。

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>我能成为此存储库的 GitHub 拉取请求的审批者吗？

目前，我们不允许外部参与者对此存储库中的拉取请求进行审批。

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>我多久能获得对我的更改请求的回复？

拉取请求通常在 10 个工作日内完成审核。


## <a name="more-resources"></a>更多资源

* 若要了解有关 Markdown 的详细信息，请转到 Markdown 创建者的网站 [Daring Fireball]。
* 若要详细了解如何使用 Git 和 GitHub，请先查看 [GitHub 帮助]。

[GitHub Home]: http://github.com
[GitHub 帮助]: http://help.github.com/
[安装 Git]: https://help.github.com/articles/set-up-git/
[Daring Fireball - Markdown]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
