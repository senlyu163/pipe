import sys
sys.path.append("..")

from pipe.pptx import PPTX 

pptx = PPTX("sample.pptx")

pptx.replace_scatter_data_at_page(14, [5.5, 2.2])
pptx.del_slide_page(3)
pptx.save("result.pptx")
