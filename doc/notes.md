# Notes
Take some notes and problems encountered.

## Notes


## Problems
1. python interpreter problem

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