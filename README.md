# RankSearch
在chafenba.com查分网站中进行爬虫获取同专业学生的成绩

1. 将`src`文件夹中的浏览器驱动更新为你当前浏览器版本（仓库自带版本为105.0.1343.33），点此下载合适版本的驱动 https://developer.microsoft.com/zh-cn/microsoft-edge/tools/webdriver/

2. 修改查询网站
```
search_link = '【你的网站链接】'
```

3. 修改查询下拉列表序号（下标从1开始）
```
index = 【下拉菜单序号】)")
```

4. 读取学号姓名excel表
* 要求格式

| 学号     | 姓名      |
| -------  | -------- |
| 12345678 | 张三     |

```
info_file = "【读取学号姓名表】.xlsx"
```

5. 修改输出csv名称
```
score_file = "【成绩单名称】.csv"
```
