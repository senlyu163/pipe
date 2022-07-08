## 目录结构

需要把代表每个人的文件夹与merge_run.py放在统一目录下，目录结构如下：
2022-07-06-xx-xx-xx-测评报告

|-- 张三

    |-- aaa.pdf

    |-- bbb.pdf

    |-- xxx.pdf

|-- 李四

    |-- aaa.pdf

    |-- bbb.pdf

    |-- xxx.pdf

|-- 通知信息.xlsx

|-- merge_run.py

## 一、环境配置

首先进入指定目录

1. 首先安装python，3.7版本及以上。
2. 安装所依赖的包，命令如下：

```shell
pip install -r requirements.txt
```

## 二、运行

```python
python merge_run.py
```
