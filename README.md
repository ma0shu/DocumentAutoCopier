
<h2 align="center">第二十三届济南市学生信息素养
提升实践活动参赛作品 课件自动复制上传器 DocumentAutoCopier</h2>

![](https://img.shields.io/badge/Latest-0.99.20220227-yellow.svg?style=for-the-badge&logo=superuser)
[![](https://img.shields.io/badge/Author-Mayiyi_A_Beginner-green.svg?style=for-the-badge)](https://space.bilibili.com/162182447)
![](https://img.shields.io/badge/Language-Python-blue.svg?style=for-the-badge&logo=python)


A python program to auto copy recently opened Office documents to a certain folder each minute.

Especially Useful in school classroom.

What's more, it even could auto upload them to your own file sharing server (Available in V0.99.20220214 or later)(Webdav Needed),

If you need this, please clone/download lnk.py, edit **line 12&37** and Compile it via pip:

```
pip install pypiwin32 webdav4 pyinstaller

pyinstaller -F -w --clean --win-private-assemblies lnkdev.py
```

If not, please use compiled versions in the Release, which only support local storage.（Need to make a dir named "PPTS" on the Desktop）

Btw I'm studying in high school so I may not read issues in time——and maybe I'm unable to solve all the problems. 

but I'm pleased to provide possible help whatever I could provide.
