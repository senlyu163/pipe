import sys
sys.path.append("..")

from pipe.pptx import PPTX 

pptx = PPTX("sample.pptx")

pptx.replace_scatter_data_at_page(14, [5.5, 2.2])
pptx.del_slide_page(3)
dul_pg = pptx.dul_slide(0)
pptx.insert_slide(5, dul_pg)
pptx.save("result.pptx")
