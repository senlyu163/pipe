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
2. Text size not resizing shape

    ```python
    text_frame = shape.text_frame
    text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    text_frame.text = 'A loooooooong text'
    ```
    Shape resizing depends on the PowerPoint rendering engine, so only happens at run-time unfortunately.

    If you click into the shape I think you'll see it adjust to fit.

    LibreOffice seems to update shape size automatically during start-up but PowerPoint doesn't.
    Ref:
    https://github.com/scanny/python-pptx/issues/147#issuecomment-76804459
    