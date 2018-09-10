# ppt-util
### An utility to copy, move, delete slides by python-pptx. 
### 3 methods are defined here:
+ #### duplicate_slide(pres, index, new_index) 
    - pres: the pptx.Presentation object
    - index: the source position of slide to be copied
    - new_index: the target position of the new slide
+ #### move_slide(pres, slide, index)
    - pres: the pptx.Presentation object
    - slide: the slide object to be moved
    - index: the target position moving to
+ #### delete_slide(pres, index)
    - pres: the pptx.Presentation object
    - index: the source position of slide to be deleted
