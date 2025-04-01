# word_mcp: Word 文档 MCP 服务器

## Overview 概述

一个模型上下文协议服务器，用于操作和处理 Word 文档  
此服务器提供了一系列工具，用于创建、编辑、转换Word文档，支持格式转换、复杂文本查询替换、文档合并,提取数据,文档对比,文档质量评估功能

## Tools 工具

1. `create_txt_file`
    - 创建一个新的 TXT 文件

2. `create_word_document`
    - 创建一个新的 Word 文档

3. `read_word_document`
    - 从 Word 文档中读取内容

4. `get_word_document_info`
    - 查询 Word 文档的基本信息

5. `add_text_to_word_document`
    - 向 Word 文档添加文本内容

6. `set_paragraph_format`
    - 设置 Word 文档中指定段落的文本格式

7. `set_paragraph_spacing`
    - 设置 Word 文档中指定段落的间距

8. `insert_image_into_word_document`
    - 在 Word 文档中插入图片

9. `insert_table_into_word_document`
    - 在 Word 文档中插入表格

10. `edit_table_cell`
    - 编辑 Word 文档中表格的单元格内容

11. `save_word_as_pdf`
    - 将 Word 文档保存为 PDF 格式

12. `convert_word_to_format`
    - 将 Word 文档转换为指定格式文件，支持格式包括 ["docx", "pdf", "txt", "html"]

13. `convert_to_word`
    - 将指定格式文件转换为 Word 文档，支持格式包括 ["docx", "pdf", "txt", "html"]

14. `close_word_document`
    - 关闭 Word 文档，可选是否保存更改

15. `edit_paragraph_text`
    - 编辑 Word 文档中指定段落的文本内容

16. `find_and_replace_text`
    - 在 Word 文档中查找并替换文本

17. `delete_paragraph`
    - 删除 Word 文档中指定的段落

18. `insert_table_of_contents`
    - 在 Word 文档中插入目录

19. `add_header_and_footer`
    - 为 Word 文档添加页眉和页脚

20. `set_page_layout`
    - 设置 Word 文档的页面布局

21. `merge_word_documents`
    - 合并多个 Word 文档

22. `complex_query_document`
    - 复杂查询 Word 文档（不限于文本）

23. `complex_replace_document`
    - 复杂替换 Word 文档

24. `extract_document_info` 文档提取函数,提取文档各种信息让AI进行分析
    - 辅助函数:`save_to_json` `save_to_csv`
    - 段落提取：提取文档中的各个段落，转换为结构化数据格式
    - 表格提取：提取文档中的表格，转换为结构化数据格式
    - 图片提取：提取文档中的图片，转换为结构化数据格式
    - 标题提取：提取文档中的各级标题，了解文档的层次结构
    - 文本提取：提取文档中的文本内容，了解文档的主题、结构和意图
    - 关键词提取：提取文档中的关键词，快速了解文档主题

25. `compare_documents` 文档对比函数
    - 文档对比：比较两个或多个文档的内容差异，识别新增、删除或修改的部分 
    - 一致性检查：检查多个文档之间的一致性，确保信息的准确性

26. `assess_document_quality` 文档质量评估函数
    - 辅助函数: `calculate_readability`,`count_syllables`
    - 可读性评估：评估文档的可读性，识别复杂的句子和段落
    - 一致性评估：评估文档中格式和内容的一致性，确保文档的专业性

## Installation 安装

