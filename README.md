# Excel2Json

- -header 表格中有几行是表头.
- -encoding 指定编码格式 默认UTF8
- -excel_path excel文件夹位置
- -json_path 输出json位置
- -script_path 输出脚本位置
- -script_template_path 脚本模板位置
- -script_manager_template_path ScriptManager模板位置



- 一般的bool ，int，float，string都支持

- arr_int为List<int>  其他基本类型类似

- dic_int_string 为Dictionary<int,string> 其他基本类型类似

- 数组和字典都是用  |  来分割 ， 如1|2|3 ，1:a|2:b|3:c

- //END为结束行，如只想生成前三行，则可在第四行添加//END

- '#' 开头的行会忽略

