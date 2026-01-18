# HParaShape 속성 확인
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글 없음")
    exit()

pset = hwp.HParameterSet.HParaShape
attrs = [x for x in dir(pset) if not x.startswith('_')]
print("HParaShape 속성:")
for attr in sorted(attrs):
    print(f"  {attr}")
