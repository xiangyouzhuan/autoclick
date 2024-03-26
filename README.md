# autoclick
它用于搜索并复制可以匹配特定字符串的 IP，并将该 IP 粘贴到一个 .XLSX 文件中。


这里的search_box_position参数可通过执行 `mouse_coordinates.py`脚本文件获取实时的鼠标的坐标
```bash
# 指定 Excel 文件名、要读取的列（例如'A'）、搜索框位置（以屏幕坐标表示）
filename = 'test.xlsx'
filenameOut = 'test.xlsx'
column_letter = 'A'
search_box_position = (949, 334)  # 替换为实际的搜索框坐标位置
search_button_position = (1763, 335)
```

注意：在执行程序时候，需要关闭已经在excel或者wps中打开的.xlsx文件，否则出现错误。


