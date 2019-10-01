# 收作业智能姬

## 介绍
一个用于将当前 Outlook 文件夹下的所有邮件附件保存到本地的插件。 

如想拿走这个奇怪的小插件，请将 [SaveAllTheHomework.Source.HomeworkBot](SaveAllTheHomework/Source/HomeworkBot.cs#L191) 以下的学号识别代码修改一下就好辣！

## 特性
- 更加灵活的学号解析功能。
	- 会尝试从发件人地址解析 ```^(\d{13})@whu.edu.cn``` 。
	- 会尝试从邮件标题解析  ```(\d{13})``` 。
	- 会尝试从附件文件名解析  ```(\d{13})``` 。
- 更加智能的附件保存功能。
	- 如有同一学号多次发送带附件的邮件，只会保存最新的一次邮件发送的附件。
	- 如遇某邮件携带多个附件，会将多个附件保存到以学号命名的子文件夹中。

## 使用

### 安装

1. 克隆或下载本仓库。
2. 使用 Visual Studio 打开本项目。
3. 编译并运行本项目。

### 使用

1. 在 Outlook 插件中将会看到本插件注册的一个按钮。 
点击按钮即可将当前 Outlook 文件夹下的所有邮件附件保存到本地。