
### 第二十三届济南市学生信息素养提升实践活动 参赛作品

# DocumentAutoCopier 课件自动复制上传器

![](https://img.shields.io/badge/Latest-0.99.20230119-yellow.svg?style=for-the-badge&logo=superuser)
[![](https://img.shields.io/badge/Author-Mayiyi_A_Beginner-green.svg?style=for-the-badge&logo=darkreader)](https://space.bilibili.com/162182447)
![](https://img.shields.io/badge/Language-Python-blue.svg?style=for-the-badge&logo=python)

## Introduction

A python program to auto copy recently opened Office documents to a certain folder each minute.

Especially Useful in school classroom.

What's more, it even could auto upload them to your own file sharing server (Available in V0.99.20220214 or later)(Webdav Needed),

## To Use:

- 1.Clone(Donwload) this repo

- 2.Compile Manually (Suggested)

(The compiled versions in the Release is old and only support local storage, what's worse, it may contain some bugs)

(Python3 environment needed)

```
pip install pypiwin32 webdav4 pyinstaller

pyinstaller -F -w --clean --win-private-assemblies lnk.py
```

-3.edit the config.ini and copy config.ini to the same directory of lnk.exe.

-4.make a shortcut of lnk.exe and copy it to Startup folder (Win+R and run shell:startup)

- - -

Btw I'm studying in high school so I may not read issues in time — and maybe I'm unable to solve all the problems. But I'm pleased to provide possible help. 
