# dbcCov工具说明

## 简介

本工具用python编写，未使用cantools、canmatrix等第三方库，仅将dbc当作文本文件做了简单的处理，目前实现了报文、结点、信号及报文描述、信号描述的转换，尚未实现信号值描述、信号值表的功能(VAL_, VAL_TABLE_)，满足一般情况下的使用。

## 命令行参数说明

```
dbcCov.exe option file
    option:     excel --> excel转dbc
                dbc   --> dbc转excel
    file:       待转换文件
```

## Excel格式

    参考Template.xls：
    Matrix表记录通信矩阵，Nodes表记录结点描述

## 开发环境

    pip install -r requirements.txt
    exe生成使用 pyinstaller -F dbcCov.p

## 开源

本工程是在[energystoryhhl](https://github.com/energystoryhhl/dbc2excel.git)和[GYemperor](https://github.com/GYemperor/excel2dbc-py.git)两位大佬的工程上修改得到，遵循LGPLv2.1开源协议。

## 参考博文：

1. [DBC文件格式解析-知乎@bingbing](https://zhuanlan.zhihu.com/p/141408513)
2. [DBC文件详细说明-CSDN@江南侠客(上海)](https://blog.csdn.net/weixin_47712251/article/details/130144332)
3. [DBC文件解析-CSDN@bitbug123](https://blog.csdn.net/u010808702/article/details/104152745)
4. [dbc文件的value description 如何编辑-CSDN@沉默的大羚羊](https://blog.csdn.net/weixin_42376614/article/details/112112910)
