# Pipe Data Between Excel and Powerpoint

This is a project to transfer data between excel and powerpoint.
Acccording to the demands, only support data from excel to powerpoint. If necessary, maybe support reverse later.

For rapid development, we decided to use python for this project development. Because there are some really nice python libraries such as [python-pptx](https://github.com/scanny/python-pptx), [openpyxl](https://github.com/ericgazoni/openpyxl) and so on.

## Python interpreter problem
My python version is 3.10.4, there is an error: 
```bash
AttributeError: module 'collections' has no attribute 'Container'
```
replace
```python
import collections
```
with:
```python
import collections.abc
```

## Reference
1. https://python-pptx.readthedocs.io/en/latest/
2. https://openpyxl.readthedocs.io/en/stable/