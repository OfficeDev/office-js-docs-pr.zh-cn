> [!NOTE]
> Outlook Windows：如果从 localhost 运行外接程序，并且看到错误"很抱歉，我们无法访问 *{your-add-in-name-here}*。 请确保具有网络连接。 如果问题继续，请稍后重试。"，您可能需要启用环回豁免。
>
> 1. 关闭 Outlook。
> 1. 打开 **任务管理器** ， **并确保msoadfsb.exe进程** 未运行。
> 1. 在 [提升的提示中](/previous-versions/windows/apps/hh780593(v=win.10)?redirectedfrom=MSDN) 设置环回豁免。
>     - 如果要使用 `https://localhost` 3000 并移植 3000 (默认配置) 运行以下命令。
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>     - 如果使用的是 和 `http://localhost` 端口 3000，请运行以下命令。
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>      **注意**：如果未使用默认端口 3000，请将命令中的端口号替换为实际端口号。
> 1. 重新启动 Outlook。
