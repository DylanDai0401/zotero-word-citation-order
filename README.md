# Word 引文顺序

[下载 Zotero 插件安装包（.xpi）](https://github.com/DylanDai0401/zotero-word-citation-order/releases/latest)

适用于 Zotero 10 和 Windows 桌面版 Microsoft Word。

1. 在 Zotero 10 桌面端打开「工具 → 插件」，从齿轮菜单选择「从文件安装插件」，选中 `Zotero引文标记.xpi`。
2. 在 Zotero 条目列表右键点击任意列标题，勾选「引文顺序」。点击这列的标题可按编号排序。选中「我的文库」可以在同一个列表中查看来自不同分类的条目。
3. 在 Windows 桌面版 Word 中打开稿件。插件会自动读取当前文档。在 Zotero 引文窗口点击右上角的「✓」后，Word 完成引文写入时，「引文顺序」列立即更新。直接在 Word 中删除或移动引文后，插件会在下一次自动检查时更新编号。

正文和附件分别从 1 开始编号，切换文档或改名后重新打开即可更新。插件保存最近一次识别的编号，关闭 Word 或重启 Zotero 后仍可查看。Word 的「Zotero → Refresh」也会更新并保存编号。

需要从磁盘读取时，在 Zotero 选择「工具 → 读取 Word…」并选取 `.docx`。

「引文顺序」是当前文稿的显示列。只有匹配到 Zotero「我的文库」的条目会显示编号。无法匹配的引文保留其原序号，后续编号可以出现跳号。同一处合并引文按字段内的条目顺序编号。插件读取 Zotero 的活动引文字段，支持正文和引用于正文的脚注、尾注。Word 自动读取适用于 Windows 桌面版 Microsoft Word。
