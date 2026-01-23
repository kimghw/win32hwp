# python-docx 라이브러리 사용 가이드

python-docx는 Python에서 Microsoft Word (.docx) 파일을 읽고 쓰기 위한 라이브러리입니다.
Word 설치 없이 서버 환경에서도 사용 가능합니다.

## 설치

```bash
pip install python-docx
```

---

## 1. 문서 열기/저장

### 1.1 새 문서 생성

```python
from docx import Document

# 새 문서 생성
doc = Document()

# 내용 추가
doc.add_paragraph('Hello, World!')

# 저장
doc.save('new_document.docx')
```

### 1.2 기존 문서 열기

```python
from docx import Document

# 파일 경로로 열기
doc = Document('existing_document.docx')

# 파일 객체로 열기
with open('document.docx', 'rb') as f:
    doc = Document(f)
```

### 1.3 문서 저장

```python
# 같은 파일에 저장
doc.save('document.docx')

# 다른 이름으로 저장
doc.save('new_name.docx')

# 스트림에 저장
from io import BytesIO
stream = BytesIO()
doc.save(stream)
```

---

## 2. 텍스트 읽기 (문단, 런)

### 2.1 문단(Paragraph) 개념

- **Paragraph**: 문서의 기본 블록 요소 (줄바꿈으로 구분)
- **Run**: 동일한 서식을 가진 텍스트 조각 (문단 내부)
- 하나의 문단은 여러 개의 Run으로 구성

```python
# 문서 내 모든 문단의 텍스트 읽기
for para in doc.paragraphs:
    print(para.text)
```

### 2.2 문단 순회 및 접근

```python
# 모든 문단
paragraphs = doc.paragraphs

# 첫 번째 문단
first_para = doc.paragraphs[0]
print(first_para.text)

# 문단 개수
para_count = len(doc.paragraphs)

# 문단과 테이블을 문서 순서대로 순회
for element in doc.iter_inner_content():
    if hasattr(element, 'text'):
        print('Paragraph:', element.text)
    else:
        print('Table')
```

### 2.3 Run 읽기

```python
for para in doc.paragraphs:
    print(f'문단: {para.text}')
    for run in para.runs:
        print(f'  Run: "{run.text}"')
        print(f'    Bold: {run.bold}')
        print(f'    Italic: {run.italic}')
```

### 2.4 문단 추가

```python
# 기본 문단 추가
doc.add_paragraph('새 문단 텍스트')

# 스타일과 함께 추가
doc.add_paragraph('목록 항목', style='List Bullet')
doc.add_paragraph('번호 목록', style='List Number')

# Run으로 텍스트 추가
para = doc.add_paragraph()
para.add_run('일반 텍스트')
run = para.add_run('굵은 텍스트')
run.bold = True
```

### 2.5 제목(Heading) 추가

```python
# 제목 추가 (level: 0=Title, 1-9=Heading 1-9)
doc.add_heading('문서 제목', level=0)  # Title
doc.add_heading('1장 소개', level=1)   # Heading 1
doc.add_heading('1.1 배경', level=2)   # Heading 2
```

---

## 3. 서식 조회 (폰트, 크기, 굵기, 색상, 정렬)

### 3.1 필요한 import

```python
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.enum.style import WD_STYLE_TYPE
```

### 3.2 Font 속성 조회

```python
for para in doc.paragraphs:
    for run in para.runs:
        font = run.font

        # 폰트 이름
        print(f'Font name: {font.name}')

        # 폰트 크기 (Pt 단위)
        if font.size:
            print(f'Font size: {font.size.pt} pt')

        # 굵게 (True/False/None)
        # None = 스타일에서 상속
        print(f'Bold: {run.bold}')

        # 기울임
        print(f'Italic: {run.italic}')

        # 밑줄
        print(f'Underline: {font.underline}')

        # 취소선
        print(f'Strike: {font.strike}')

        # 위첨자/아래첨자
        print(f'Superscript: {font.superscript}')
        print(f'Subscript: {font.subscript}')
```

### 3.3 폰트 색상 조회

```python
for run in para.runs:
    font = run.font
    color = font.color

    if color.rgb:
        # RGB 값 (예: RGBColor(255, 0, 0))
        print(f'RGB: {color.rgb}')

    if color.theme_color:
        # 테마 색상 (MSO_THEME_COLOR 값)
        print(f'Theme color: {color.theme_color}')
```

### 3.4 Font 속성 설정

```python
run = para.add_run('서식 적용 텍스트')
font = run.font

# 폰트 이름
font.name = '맑은 고딕'

# 폰트 크기
font.size = Pt(12)

# 굵게/기울임
run.bold = True
run.italic = True

# 밑줄 (True=단일 밑줄, WD_UNDERLINE.DOUBLE=이중 밑줄)
font.underline = True
# font.underline = WD_UNDERLINE.DOUBLE
# font.underline = WD_UNDERLINE.DOTTED

# 취소선
font.strike = True

# 색상 (RGB)
font.color.rgb = RGBColor(255, 0, 0)  # 빨강
font.color.rgb = RGBColor(0x42, 0x24, 0xE9)  # Hex 값
```

### 3.5 문단 정렬 조회 및 설정

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 정렬 조회
alignment = para.alignment
print(f'Alignment: {alignment}')
# None = 상속, 0=LEFT, 1=CENTER, 2=RIGHT, 3=JUSTIFY

# 정렬 설정
para.alignment = WD_ALIGN_PARAGRAPH.LEFT      # 왼쪽
para.alignment = WD_ALIGN_PARAGRAPH.CENTER    # 가운데
para.alignment = WD_ALIGN_PARAGRAPH.RIGHT     # 오른쪽
para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY   # 양쪽
```

### 3.6 문단 서식 (ParagraphFormat)

```python
from docx.shared import Pt, Inches

para_format = para.paragraph_format

# 들여쓰기
para_format.left_indent = Inches(0.5)     # 왼쪽 들여쓰기
para_format.right_indent = Inches(0.5)    # 오른쪽 들여쓰기
para_format.first_line_indent = Inches(0.5)  # 첫 줄 들여쓰기

# 줄 간격
para_format.line_spacing = 1.5            # 1.5줄 간격
para_format.line_spacing = Pt(18)         # 18pt 고정
# para_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY  # 정확히

# 문단 간격
para_format.space_before = Pt(12)         # 문단 앞 간격
para_format.space_after = Pt(12)          # 문단 뒤 간격

# 페이지 나누기
para_format.keep_together = True          # 문단 같은 페이지 유지
para_format.keep_with_next = True         # 다음 문단과 같은 페이지
para_format.page_break_before = True      # 문단 앞 페이지 나누기
```

### 3.7 정렬 열거형 값 (WD_ALIGN_PARAGRAPH)

| 열거형 | 값 | 설명 |
|--------|---|------|
| LEFT | 0 | 왼쪽 정렬 |
| CENTER | 1 | 가운데 정렬 |
| RIGHT | 2 | 오른쪽 정렬 |
| JUSTIFY | 3 | 양쪽 정렬 |
| DISTRIBUTE | 4 | 균등 배분 |
| JUSTIFY_MED | 5 | 중간 문자 압축 양쪽 정렬 |
| JUSTIFY_HI | 6 | 높은 문자 압축 양쪽 정렬 |
| JUSTIFY_LOW | 7 | 낮은 문자 압축 양쪽 정렬 |
| THAI_JUSTIFY | 8 | 태국어 양쪽 정렬 |

---

## 4. 테이블 읽기 (행, 열, 셀, 병합)

### 4.1 테이블 접근

```python
# 문서 내 모든 테이블
tables = doc.tables

# 첫 번째 테이블
table = doc.tables[0]

# 테이블 개수
print(f'Table count: {len(doc.tables)}')
```

### 4.2 행/열/셀 접근

```python
table = doc.tables[0]

# 행 개수
row_count = len(table.rows)

# 열 개수
col_count = len(table.columns)

# 모든 셀 순회 (행 기준)
for row in table.rows:
    for cell in row.cells:
        print(cell.text)

# 특정 셀 접근 (0-based 인덱스)
cell = table.cell(0, 0)  # 첫 번째 행, 첫 번째 열
print(cell.text)

# 열 기준 순회
for col in table.columns:
    for cell in col.cells:
        print(cell.text)
```

### 4.3 셀 내용 읽기

```python
cell = table.cell(0, 0)

# 셀 전체 텍스트
print(cell.text)

# 셀 내 문단 순회
for para in cell.paragraphs:
    print(para.text)
    for run in para.runs:
        print(f'  Run: {run.text}')

# 셀 내 중첩 테이블
for nested_table in cell.tables:
    print('Nested table found')
```

### 4.4 테이블 생성

```python
# 3행 4열 테이블 생성
table = doc.add_table(rows=3, cols=4)

# 스타일 적용
table.style = 'Table Grid'

# 셀에 값 입력
table.cell(0, 0).text = '헤더1'
table.cell(0, 1).text = '헤더2'

# 행 추가
row = table.add_row()
row.cells[0].text = '새 행 데이터'

# 열 추가
col = table.add_column(Inches(1.0))
```

### 4.5 병합 셀 처리

```python
# 셀 병합
# merge(other_cell)는 현재 셀과 다른 셀 사이의 모든 셀을 병합
a = table.cell(0, 0)
b = table.cell(0, 2)
merged_cell = a.merge(b)  # 0행의 0,1,2열 병합

# 병합된 셀 텍스트 설정
merged_cell.text = '병합된 셀'

# 병합된 셀 여부 확인 (간접적으로)
# python-docx는 병합된 셀을 동일한 _tc 요소를 공유하는 것으로 처리
# 병합된 영역에서 같은 셀을 여러 번 반환함

# 병합 정보 확인 방법
def get_merged_cells(table):
    """병합 정보 확인"""
    seen = set()
    merged_cells = []

    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            # 같은 셀이 여러 위치에서 반환되면 병합됨
            cell_id = id(cell._tc)
            if cell_id in seen:
                continue
            seen.add(cell_id)

            # 셀이 차지하는 실제 행/열 범위 확인
            # (셀 내부 XML에서 gridSpan, vMerge 속성 확인 필요)
            merged_cells.append((row_idx, col_idx, cell.text))

    return merged_cells
```

### 4.6 테이블 스타일

```python
# 테이블 스타일 설정
table.style = 'Table Grid'      # 격자
table.style = 'Light Shading'   # 밝은 음영
table.style = 'Light Grid'      # 밝은 격자

# 테이블 정렬
from docx.enum.table import WD_TABLE_ALIGNMENT
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# 자동 맞춤
table.autofit = True
```

### 4.7 셀 서식

```python
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt

cell = table.cell(0, 0)

# 셀 너비
cell.width = Inches(1.5)

# 수직 정렬
cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
# TOP, CENTER, BOTTOM, BOTH

# 셀 내 문단 정렬 (가운데 정렬)
for para in cell.paragraphs:
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

### 4.8 행/열 크기 설정

```python
from docx.shared import Inches, Pt, Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE

# 행 높이
row = table.rows[0]
row.height = Inches(0.5)
row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY  # 정확히
# AT_LEAST (최소), EXACTLY (정확히), AUTO (자동)

# 열 너비
for cell in table.columns[0].cells:
    cell.width = Inches(2.0)
```

---

## 5. 이미지/도형 처리

### 5.1 이미지 추가

```python
from docx.shared import Inches, Cm

# 이미지 추가 (새 문단에)
doc.add_picture('image.png')

# 크기 지정
doc.add_picture('image.png', width=Inches(2.0))
doc.add_picture('image.png', height=Cm(5.0))

# 너비와 높이 모두 지정 (비율 무시됨)
doc.add_picture('image.png', width=Inches(2.0), height=Inches(1.5))

# Run에 이미지 추가 (같은 문단에 텍스트와 함께)
para = doc.add_paragraph()
run = para.add_run()
run.add_picture('image.png', width=Inches(1.0))
run.add_text(' 이미지 옆 텍스트')
```

### 5.2 인라인 도형(InlineShapes) 읽기

```python
# 문서 내 인라인 도형 (이미지 등) 접근
inline_shapes = doc.inline_shapes

for shape in inline_shapes:
    # 도형 타입
    print(f'Type: {shape.type}')  # MSO_SHAPE_TYPE

    # 크기
    print(f'Width: {shape.width.inches} inches')
    print(f'Height: {shape.height.inches} inches')
```

### 5.3 이미지 데이터 추출 (ZIP 방식)

python-docx는 이미지 바이너리 데이터 직접 접근을 지원하지 않습니다.
DOCX 파일은 ZIP 형식이므로 직접 추출해야 합니다.

```python
import zipfile
import os

def extract_images_from_docx(docx_path, output_dir):
    """DOCX에서 이미지 추출"""
    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            # 이미지는 word/media/ 폴더에 저장됨
            if file_name.startswith('word/media/'):
                # 이미지 파일 추출
                image_name = os.path.basename(file_name)
                with zip_ref.open(file_name) as src:
                    with open(os.path.join(output_dir, image_name), 'wb') as dst:
                        dst.write(src.read())
                print(f'Extracted: {image_name}')

# 사용 예
extract_images_from_docx('document.docx', 'extracted_images')
```

### 5.4 이미지 제한사항

| 기능 | 지원 여부 |
|------|----------|
| 인라인 이미지 추가 | O |
| 이미지 크기 조절 | O |
| 이미지 크기 읽기 | O |
| 이미지 바이너리 추출 | X (ZIP 사용) |
| 플로팅 이미지 | X |
| 텍스트 배치 설정 | X |

---

## 6. 단위 변환

### 6.1 단위 클래스

```python
from docx.shared import Inches, Cm, Mm, Pt, Emu, Twips

# 1 inch = 914,400 EMU
# 1 inch = 2.54 cm
# 1 inch = 72 pt

# 단위 생성
width = Inches(2.0)
height = Cm(5.0)
font_size = Pt(12)

# 단위 변환 (읽기 전용 속성)
length = Inches(1.0)
print(f'{length.inches} inches')   # 1.0
print(f'{length.cm} cm')           # 2.54
print(f'{length.pt} pt')           # 72.0
print(f'{length.emu} emu')         # 914400
```

### 6.2 EMU (English Metric Units)

```python
# 모든 길이는 내부적으로 EMU로 저장
# 1 inch = 914,400 EMU
# 1 cm = 360,000 EMU
# 1 pt = 12,700 EMU

from docx.shared import Emu

# EMU 직접 사용
width = Emu(914400)  # 1 inch

# HWP 단위와 비교
# HWP: 1 inch = 7,200 HWPUNIT
# DOCX: 1 inch = 914,400 EMU
# 변환: EMU = HWPUNIT * 127
```

---

## 7. 섹션 및 페이지 설정

### 7.1 섹션 접근

```python
# 모든 섹션
sections = doc.sections

# 첫 번째 섹션
section = doc.sections[0]

# 새 섹션 추가
from docx.enum.section import WD_SECTION_START
new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
```

### 7.2 페이지 크기 및 방향

```python
from docx.shared import Inches, Cm
from docx.enum.section import WD_ORIENTATION

section = doc.sections[0]

# 페이지 크기
section.page_width = Inches(8.5)   # 너비
section.page_height = Inches(11)    # 높이

# 페이지 방향
section.orientation = WD_ORIENTATION.PORTRAIT   # 세로
section.orientation = WD_ORIENTATION.LANDSCAPE  # 가로

# 가로 방향으로 변경 시 너비/높이 교환 필요
if section.orientation == WD_ORIENTATION.LANDSCAPE:
    section.page_width, section.page_height = section.page_height, section.page_width
```

### 7.3 여백 설정

```python
from docx.shared import Inches, Cm

section = doc.sections[0]

# 여백 설정
section.left_margin = Inches(1.0)
section.right_margin = Inches(1.0)
section.top_margin = Inches(1.0)
section.bottom_margin = Inches(1.0)

# 제본 여백
section.gutter = Inches(0.5)

# 머리글/꼬리글 거리
section.header_distance = Inches(0.5)
section.footer_distance = Inches(0.5)
```

---

## 8. 스타일 관리

### 8.1 스타일 조회

```python
from docx.enum.style import WD_STYLE_TYPE

# 모든 스타일
styles = doc.styles

# 스타일 순회
for style in doc.styles:
    print(f'{style.name} ({style.type})')

# 특정 타입 스타일만
for style in doc.styles:
    if style.type == WD_STYLE_TYPE.PARAGRAPH:
        print(f'Paragraph style: {style.name}')
```

### 8.2 스타일 적용

```python
# 문단에 스타일 적용
para = doc.add_paragraph('스타일 적용', style='Heading 1')

# 또는
para.style = 'Normal'
para.style = doc.styles['Heading 2']

# Run에 문자 스타일 적용
run = para.add_run('강조 텍스트')
run.style = 'Emphasis'
```

### 8.3 주요 기본 스타일

**문단 스타일:**
- Normal, Heading 1-9, Title, Subtitle
- Body Text, Quote, Intense Quote
- List Bullet, List Number, List Paragraph

**문자 스타일:**
- Strong, Emphasis, Intense Emphasis
- Book Title, Subtle Reference

---

## 9. 문서 속성

### 9.1 코어 속성 조회/설정

```python
# 문서 속성 접근
core_props = doc.core_properties

# 읽기
print(f'Title: {core_props.title}')
print(f'Author: {core_props.author}')
print(f'Subject: {core_props.subject}')
print(f'Keywords: {core_props.keywords}')
print(f'Created: {core_props.created}')
print(f'Modified: {core_props.modified}')
print(f'Last modified by: {core_props.last_modified_by}')

# 쓰기
core_props.title = '문서 제목'
core_props.author = '작성자 이름'
core_props.subject = '문서 주제'
core_props.keywords = '키워드1, 키워드2'
```

---

## 10. 전체 예제 - 문서 읽기

```python
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def read_docx_complete(file_path):
    """DOCX 문서 전체 읽기"""
    doc = Document(file_path)

    print('='*50)
    print('문서 속성')
    print('='*50)
    print(f'Title: {doc.core_properties.title}')
    print(f'Author: {doc.core_properties.author}')

    print('\n' + '='*50)
    print('문서 내용')
    print('='*50)

    # 문단과 테이블을 순서대로 읽기
    for element in doc.iter_inner_content():
        if hasattr(element, 'runs'):
            # Paragraph
            para = element
            print(f'\n[문단] {para.text}')
            print(f'  정렬: {para.alignment}')
            print(f'  스타일: {para.style.name}')

            for run in para.runs:
                font = run.font
                print(f'  [Run] "{run.text}"')
                if run.bold: print('    - 굵게')
                if run.italic: print('    - 기울임')
                if font.name: print(f'    - 폰트: {font.name}')
                if font.size: print(f'    - 크기: {font.size.pt}pt')
        else:
            # Table
            table = element
            print(f'\n[테이블] {len(table.rows)}행 x {len(table.columns)}열')
            for row_idx, row in enumerate(table.rows):
                row_data = [cell.text for cell in row.cells]
                print(f'  Row {row_idx}: {row_data}')

    # 인라인 도형 (이미지 등)
    if doc.inline_shapes:
        print('\n' + '='*50)
        print('인라인 도형')
        print('='*50)
        for i, shape in enumerate(doc.inline_shapes):
            print(f'Shape {i}: {shape.type}')
            print(f'  크기: {shape.width.inches:.2f}" x {shape.height.inches:.2f}"')

if __name__ == '__main__':
    read_docx_complete('sample.docx')
```

---

## 11. 전체 예제 - 문서 작성

```python
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def create_sample_docx(output_path):
    """샘플 DOCX 문서 생성"""
    doc = Document()

    # 문서 속성
    doc.core_properties.title = 'python-docx 테스트 문서'
    doc.core_properties.author = 'Python Script'

    # 제목
    doc.add_heading('python-docx 라이브러리 테스트', level=0)

    # 일반 문단
    para = doc.add_paragraph('이것은 일반 문단입니다. ')

    # 서식 적용된 텍스트
    run = para.add_run('굵은 텍스트, ')
    run.bold = True

    run = para.add_run('기울임 텍스트, ')
    run.italic = True

    run = para.add_run('밑줄 텍스트, ')
    run.underline = True

    run = para.add_run('빨간 텍스트')
    run.font.color.rgb = RGBColor(255, 0, 0)

    # 가운데 정렬 문단
    para = doc.add_paragraph('가운데 정렬된 문단입니다.')
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 목록
    doc.add_heading('목록 예제', level=1)
    doc.add_paragraph('첫 번째 항목', style='List Bullet')
    doc.add_paragraph('두 번째 항목', style='List Bullet')
    doc.add_paragraph('세 번째 항목', style='List Bullet')

    # 테이블
    doc.add_heading('테이블 예제', level=1)
    table = doc.add_table(rows=3, cols=3)
    table.style = 'Table Grid'

    # 헤더 행
    header_cells = table.rows[0].cells
    header_cells[0].text = '이름'
    header_cells[1].text = '나이'
    header_cells[2].text = '직업'

    # 데이터 행
    data = [
        ['홍길동', '30', '개발자'],
        ['김영희', '25', '디자이너']
    ]

    for row_idx, row_data in enumerate(data, 1):
        cells = table.rows[row_idx].cells
        for col_idx, value in enumerate(row_data):
            cells[col_idx].text = value

    # 페이지 나누기
    doc.add_page_break()

    # 새 페이지 내용
    doc.add_heading('두 번째 페이지', level=1)
    doc.add_paragraph('페이지 나누기 후의 내용입니다.')

    # 저장
    doc.save(output_path)
    print(f'문서 저장 완료: {output_path}')

if __name__ == '__main__':
    create_sample_docx('sample_output.docx')
```

---

## 참고 자료

- [python-docx 공식 문서](https://python-docx.readthedocs.io/)
- [python-docx GitHub](https://github.com/python-openxml/python-docx)
- [python-docx PyPI](https://pypi.org/project/python-docx/)
- [API Reference - Document](https://python-docx.readthedocs.io/en/latest/api/document.html)
- [API Reference - Text](https://python-docx.readthedocs.io/en/latest/api/text.html)
- [Quickstart Guide](https://python-docx.readthedocs.io/en/latest/user/quickstart.html)
