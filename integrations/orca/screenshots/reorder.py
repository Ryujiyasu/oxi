import sys
from pptx import Presentation
src, first, out = sys.argv[1], int(sys.argv[2]), sys.argv[3]
p = Presentation(src)
lst = p.slides._sldIdLst
ids = list(lst)
el = ids[first]
lst.remove(el); lst.insert(0, el)
p.save(out)
