# Microsoft Word win32com 자동화 가이드

최종 수정일: 2026-01-23

이 문서는 Python win32com을 사용하여 Microsoft Word를 자동화하는 방법을 정리합니다.

---

## 목차

1. [Word 애플리케이션 연결/생성](#1-word-애플리케이션-연결생성)
2. [문서 열기/저장](#2-문서-열기저장)
3. [텍스트 추출 (Range, Paragraphs)](#3-텍스트-추출-range-paragraphs)
4. [서식 조회 (Font, ParagraphFormat)](#4-서식-조회-font-paragraphformat)
5. [테이블 처리 (Tables, Rows, Cells)](#5-테이블-처리-tables-rows-cells)
6. [Selection 사용법](#6-selection-사용법)
7. [이미지 처리 (Shapes, InlineShapes)](#7-이미지-처리-shapes-inlineshapes)
8. [상수(Constants) 참조](#8-상수constants-참조)
9. [단위 변환](#9-단위-변환)
10. [HWP vs Word 비교](#10-hwp-vs-word-비교)

---

## 1. Word 애플리케이션 연결/생성

### 1.1 새 인스턴스 생성

```python
import win32com.client as win32

# 방법 1: Dispatch (late-binding, 상수 사용 불가)
word = win32.Dispatch('Word.Application')

# 방법 2: EnsureDispatch (early-binding, 상수 사용 가능)
word = win32.gencache.EnsureDispatch('Word.Application')
```

**참고:** `gencache.EnsureDispatch`를 사용하면 `win32com.client.constants`를 통해 Word 상수에 접근할 수 있습니다.

### 1.2 기본 설정

```python
# 화면 표시 설정
word.Visible = True    # Word 창 표시 (디버깅용)
word.Visible = False   # 백그라운드 실행 (자동화용)

# 경고 메시지 숨기기
word.DisplayAlerts = 0  # wdAlertsNone

# 새 문서 생성
doc = word.Documents.Add()
```

### 1.3 실행 중인 Word에 연결

```python
try:
    word = win32.GetActiveObject('Word.Application')
    print("기존 Word 인스턴스에 연결됨")
except:
    word = win32.Dispatch('Word.Application')
    print("새 Word 인스턴스 생성됨")
```

### 1.4 종료 처리

```python
# 문서 저장하지 않고 닫기
doc.Close(SaveChanges=0)  # 0=저장 안함, 1=저장, 2=사용자에게 물음

# Word 종료
word.Quit()
```

---

## 2. 문서 열기/저장

### 2.1 문서 열기

```python
# 기본 열기
doc = word.Documents.Open(r"C:\path\to\document.docx")

# 옵션과 함께 열기
doc = word.Documents.Open(
    FileName=r"C:\path\to\document.docx",
    ReadOnly=True,          # 읽기 전용
    AddToRecentFiles=False  # 최근 문서 목록에 추가 안함
)

# 활성 문서 접근
doc = word.ActiveDocument
```

### 2.2 문서 저장 (Save)

```python
# 현재 위치에 저장
doc.Save()
```

### 2.3 다른 이름으로 저장 (SaveAs/SaveAs2)

```python
# 기본 저장
doc.SaveAs(r"C:\path\to\new_document.docx")

# 형식 지정하여 저장
doc.SaveAs(
    FileName=r"C:\path\to\output.pdf",
    FileFormat=17  # wdFormatPDF
)

# SaveAs2 메서드 (Word 2010+)
doc.SaveAs2(
    FileName=r"C:\path\to\output.docx",
    FileFormat=16  # wdFormatDocumentDefault
)
```

### 2.4 주요 파일 형식 (WdSaveFormat)

| 상수명 | 값 | 확장자 | 설명 |
|--------|---|--------|------|
| wdFormatDocument | 0 | .doc | Word 97-2003 문서 |
| wdFormatTemplate | 1 | .dot | Word 97-2003 템플릿 |
| wdFormatText | 2 | .txt | 일반 텍스트 |
| wdFormatRTF | 6 | .rtf | 서식 있는 텍스트 |
| wdFormatHTML | 8 | .html | HTML 형식 |
| wdFormatFilteredHTML | 10 | .html | 필터링된 HTML |
| wdFormatXMLDocument | 12 | .docx | Word 2007+ XML 문서 |
| wdFormatDocumentDefault | 16 | .docx | 기본 문서 형식 |
| wdFormatPDF | 17 | .pdf | PDF 형식 |
| wdFormatXPS | 18 | .xps | XPS 형식 |
| wdFormatOpenDocumentText | 23 | .odt | OpenDocument 형식 |

---

## 3. 텍스트 추출 (Range, Paragraphs)

### 3.1 Range 객체

Range는 문서의 연속된 영역을 나타냅니다. 시작과 끝 위치를 지정하여 생성합니다.

```python
# 전체 문서 Range
full_range = doc.Content
full_text = full_range.Text

# 특정 범위 지정 (0-based 문자 위치)
partial_range = doc.Range(Start=0, End=100)
partial_text = partial_range.Text

# Range 속성
print(f"시작: {partial_range.Start}, 끝: {partial_range.End}")
```

### 3.2 Paragraphs 컬렉션

```python
# 전체 문단 수
para_count = doc.Paragraphs.Count

# 문단 순회
for para in doc.Paragraphs:
    text = para.Range.Text
    style = para.Style.NameLocal  # 스타일 이름
    print(f"[{style}] {text.strip()}")

# 특정 문단 접근 (1-based index)
first_para = doc.Paragraphs(1)
last_para = doc.Paragraphs(doc.Paragraphs.Count)

# 문단 Range의 텍스트
para_text = doc.Paragraphs(1).Range.Text
```

**주의:** 문단 텍스트 끝에는 줄바꿈 문자(`\r`)가 포함됩니다. 필요시 `.strip()` 또는 `.rstrip('\r')` 사용.

### 3.3 Words, Sentences, Characters

```python
# 단어 접근
word_count = doc.Words.Count
first_word = doc.Words(1).Text

# 문장 접근
sentence_count = doc.Sentences.Count
first_sentence = doc.Sentences(1).Text

# 문자 접근
char_count = doc.Characters.Count
first_char = doc.Characters(1).Text
```

### 3.4 텍스트 추출 예시

```python
def extract_all_text(doc):
    """문서의 모든 텍스트 추출"""
    return doc.Content.Text

def extract_paragraphs(doc):
    """문단별 텍스트 추출"""
    paragraphs = []
    for para in doc.Paragraphs:
        text = para.Range.Text.rstrip('\r')
        if text:  # 빈 문단 제외
            paragraphs.append(text)
    return paragraphs

def extract_with_style(doc):
    """스타일과 함께 텍스트 추출"""
    result = []
    for para in doc.Paragraphs:
        result.append({
            'text': para.Range.Text.rstrip('\r'),
            'style': para.Style.NameLocal,
            'start': para.Range.Start,
            'end': para.Range.End
        })
    return result
```

---

## 4. 서식 조회 (Font, ParagraphFormat)

### 4.1 Font 속성

Range 또는 Selection의 Font 객체를 통해 글자 서식에 접근합니다.

```python
# Range의 Font
font = doc.Paragraphs(1).Range.Font

# 주요 Font 속성
font_name = font.Name            # 폰트 이름 (예: "맑은 고딕")
font_name_ascii = font.NameAscii # 영문 폰트 이름
font_name_other = font.NameOther # 기타 언어 폰트
font_size = font.Size            # 폰트 크기 (pt)
font_color = font.Color          # 글자색 (RGB 정수값)

# 스타일 속성 (True/False 또는 wdUndefined=-9999999)
is_bold = font.Bold              # 굵게
is_italic = font.Italic          # 기울임
is_underline = font.Underline    # 밑줄 종류
is_strikethrough = font.StrikeThrough  # 취소선

# 기타 속성
font.Subscript      # 아래 첨자
font.Superscript    # 위 첨자
font.AllCaps        # 모두 대문자
font.SmallCaps      # 작은 대문자
font.Shadow         # 그림자
font.Emboss         # 양각
font.Engrave        # 음각
```

### 4.2 Font 속성값 해석

```python
# wdUndefined: 혼합된 서식 (선택 영역 내 서식이 다름)
wdUndefined = -9999999

def get_font_info(range_obj):
    """Range의 Font 정보 조회"""
    font = range_obj.Font

    info = {
        'name': font.Name,
        'size': font.Size if font.Size != wdUndefined else 'Mixed',
        'bold': None,
        'italic': None,
        'underline': font.Underline,
        'color': font.Color
    }

    # Bold/Italic: 0=False, -1=True, wdUndefined=Mixed
    if font.Bold == -1:
        info['bold'] = True
    elif font.Bold == 0:
        info['bold'] = False
    else:
        info['bold'] = 'Mixed'

    return info
```

### 4.3 색상 처리

```python
# Word 색상: RGB 정수값 (0x00BBGGRR 형식과 다름!)
# Word는 실제로 0x00RRGGBB 형식 사용

def word_color_to_rgb(color):
    """Word 색상값을 RGB 튜플로 변환"""
    if color < 0:  # 자동 색상 또는 undefined
        return None
    blue = (color >> 16) & 0xFF
    green = (color >> 8) & 0xFF
    red = color & 0xFF
    return (red, green, blue)

def rgb_to_word_color(r, g, b):
    """RGB를 Word 색상값으로 변환"""
    return r | (g << 8) | (b << 16)

# 예시
color = doc.Paragraphs(1).Range.Font.Color
rgb = word_color_to_rgb(color)
print(f"RGB: {rgb}")
```

### 4.4 ParagraphFormat 속성

```python
# 문단 서식 접근
pf = doc.Paragraphs(1).Format

# 정렬
alignment = pf.Alignment
# 0=wdAlignParagraphLeft (왼쪽)
# 1=wdAlignParagraphCenter (가운데)
# 2=wdAlignParagraphRight (오른쪽)
# 3=wdAlignParagraphJustify (양쪽)

# 들여쓰기 (pt 단위)
left_indent = pf.LeftIndent      # 왼쪽 들여쓰기
right_indent = pf.RightIndent    # 오른쪽 들여쓰기
first_line = pf.FirstLineIndent  # 첫 줄 들여쓰기 (양수=들여쓰기, 음수=내어쓰기)

# 문단 간격 (pt 단위)
space_before = pf.SpaceBefore    # 문단 위 간격
space_after = pf.SpaceAfter      # 문단 아래 간격

# 줄 간격
line_spacing = pf.LineSpacing         # 줄 간격 값
line_spacing_rule = pf.LineSpacingRule # 줄 간격 규칙
# 0=wdLineSpaceSingle (1줄)
# 1=wdLineSpace1pt5 (1.5줄)
# 2=wdLineSpaceDouble (2줄)
# 3=wdLineSpaceAtLeast (최소)
# 4=wdLineSpaceExactly (고정)
# 5=wdLineSpaceMultiple (배수)
```

### 4.5 서식 조회 예시

```python
def get_paragraph_format(para):
    """문단 서식 정보 조회"""
    pf = para.Format
    font = para.Range.Font

    alignment_names = {
        0: '왼쪽', 1: '가운데', 2: '오른쪽', 3: '양쪽'
    }

    return {
        # 글자 서식
        'font_name': font.Name,
        'font_size': font.Size,
        'bold': font.Bold == -1,
        'italic': font.Italic == -1,

        # 문단 서식
        'alignment': alignment_names.get(pf.Alignment, str(pf.Alignment)),
        'left_indent': pf.LeftIndent,
        'first_line_indent': pf.FirstLineIndent,
        'space_before': pf.SpaceBefore,
        'space_after': pf.SpaceAfter,
        'line_spacing': pf.LineSpacing,

        # 스타일
        'style': para.Style.NameLocal
    }
```

---

## 5. 테이블 처리 (Tables, Rows, Cells)

### 5.1 테이블 접근

```python
# 테이블 수
table_count = doc.Tables.Count

# 테이블 순회
for table in doc.Tables:
    row_count = table.Rows.Count
    col_count = table.Columns.Count
    print(f"테이블: {row_count}행 x {col_count}열")

# 특정 테이블 접근 (1-based index)
first_table = doc.Tables(1)
```

### 5.2 셀 접근

```python
table = doc.Tables(1)

# 방법 1: Cell(row, col) 메서드
cell = table.Cell(1, 1)  # 1행 1열 (1-based)
cell_text = cell.Range.Text

# 방법 2: Rows/Cells 컬렉션
row = table.Rows(1)
cell = row.Cells(1)
cell_text = cell.Range.Text

# 셀 텍스트에서 종료 마커 제거
clean_text = cell_text.rstrip('\r\x07')  # \x07 = 셀 종료 마커
```

**주의:** 셀 텍스트는 줄바꿈(`\r`)과 셀 종료 마커(`\x07`)를 포함합니다.

### 5.3 테이블 데이터 추출

```python
def extract_table_data(table):
    """테이블을 2D 리스트로 추출"""
    data = []
    for row in table.Rows:
        row_data = []
        for cell in row.Cells:
            # 셀 텍스트 정리
            text = cell.Range.Text.rstrip('\r\x07')
            row_data.append(text)
        data.append(row_data)
    return data

def extract_all_tables(doc):
    """문서의 모든 테이블 추출"""
    tables_data = []
    for i, table in enumerate(doc.Tables, 1):
        tables_data.append({
            'index': i,
            'rows': table.Rows.Count,
            'cols': table.Columns.Count,
            'data': extract_table_data(table)
        })
    return tables_data
```

### 5.4 테이블 생성

```python
# Range 위치에 테이블 생성
rang = doc.Range(Start=0, End=0)
table = doc.Tables.Add(
    Range=rang,
    NumRows=3,
    NumColumns=4,
    DefaultTableBehavior=1,  # wdWord9TableBehavior
    AutoFitBehavior=1        # wdAutoFitContent
)

# 셀에 텍스트 입력
table.Cell(1, 1).Range.Text = "헤더 1"
table.Cell(1, 2).Range.Text = "헤더 2"
```

### 5.5 테이블 서식

```python
# 셀 서식
cell = table.Cell(1, 1)
cell.Range.Bold = True
cell.Range.Font.Size = 12

# 셀 배경색
cell.Shading.BackgroundPatternColor = rgb_to_word_color(200, 200, 200)

# 셀 정렬
cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter

# 테두리
table.Borders.Enable = True
table.Borders.InsideLineStyle = 1   # wdLineStyleSingle
table.Borders.OutsideLineStyle = 1
```

### 5.6 행/열 추가 및 삭제

```python
# 행 추가
table.Rows.Add()  # 마지막에 추가
table.Rows.Add(BeforeRow=table.Rows(1))  # 첫 번째 행 앞에 추가

# 열 추가
table.Columns.Add()

# 행 삭제
table.Rows(1).Delete()

# 열 삭제
table.Columns(1).Delete()
```

### 5.7 셀 병합

```python
# 셀 병합
cell1 = table.Cell(1, 1)
cell2 = table.Cell(1, 3)
cell1.Merge(MergeTo=cell2)

# 병합 여부 확인 (간접적)
# RowIndex, ColumnIndex로 위치 확인 가능
```

---

## 6. Selection 사용법

### 6.1 Selection 객체

Selection은 현재 선택 영역 또는 커서 위치를 나타냅니다.

```python
# Selection 접근
sel = word.Selection

# 현재 Selection 정보
start = sel.Start
end = sel.End
text = sel.Text
```

### 6.2 커서 이동

```python
# 문서 시작/끝으로 이동
sel.HomeKey(Unit=6)  # wdStory = 6, 문서 시작
sel.EndKey(Unit=6)   # 문서 끝

# 줄 시작/끝으로 이동
sel.HomeKey(Unit=5)  # wdLine = 5
sel.EndKey(Unit=5)

# 방향 이동
sel.MoveLeft(Unit=1, Count=1)   # 한 문자 왼쪽
sel.MoveRight(Unit=1, Count=1)  # 한 문자 오른쪽
sel.MoveUp(Unit=5, Count=1)     # 한 줄 위
sel.MoveDown(Unit=5, Count=1)   # 한 줄 아래

# 단어 단위 이동
sel.MoveLeft(Unit=2, Count=1)   # wdWord = 2
sel.MoveRight(Unit=2, Count=1)

# 문단 단위 이동
sel.MoveUp(Unit=4, Count=1)     # wdParagraph = 4
sel.MoveDown(Unit=4, Count=1)
```

### 6.3 텍스트 선택

```python
# 확장하여 선택 (Extend=1)
sel.MoveRight(Unit=2, Count=3, Extend=1)  # 3단어 선택

# 전체 선택
sel.WholeStory()

# 현재 줄 선택
sel.HomeKey(Unit=5)
sel.EndKey(Unit=5, Extend=1)

# 현재 문단 선택
sel.HomeKey(Unit=4)  # 문단 시작
sel.EndKey(Unit=4, Extend=1)  # 문단 끝까지 선택

# 특정 범위로 이동
sel.SetRange(Start=0, End=100)
```

### 6.4 텍스트 입력

```python
# 현재 위치에 텍스트 입력
sel.TypeText("새로운 텍스트")

# 문단 삽입
sel.TypeParagraph()

# 선택 영역 대체
sel.Text = "대체할 텍스트"

# 줄바꿈 삽입
sel.InsertBreak(Type=6)  # wdLineBreak
sel.InsertBreak(Type=7)  # wdPageBreak
```

### 6.5 Find and Replace

```python
# EnsureDispatch 사용 시 (상수 사용 가능)
from win32com.client import constants as c

find = sel.Find
find.ClearFormatting()
find.Replacement.ClearFormatting()

find.Text = "찾을 텍스트"
find.Replacement.Text = "대체할 텍스트"
find.Forward = True
find.Wrap = c.wdFindContinue  # 1

# 찾기 실행
found = find.Execute()

# 모두 바꾸기
find.Execute(Replace=c.wdReplaceAll)  # 2
```

```python
# 상수 직접 사용 (Dispatch 사용 시)
find = sel.Find
find.ClearFormatting()
find.Replacement.ClearFormatting()

find.Text = "찾을 텍스트"
find.Replacement.Text = "대체할 텍스트"
find.Forward = True
find.Wrap = 1  # wdFindContinue

# 찾기
find.Execute()

# 모두 바꾸기
find.Execute(Replace=2)  # wdReplaceAll
```

### 6.6 Selection 서식 설정

```python
# 글자 서식
sel.Font.Bold = True
sel.Font.Size = 14
sel.Font.Name = "맑은 고딕"
sel.Font.Color = rgb_to_word_color(255, 0, 0)

# 문단 서식
sel.ParagraphFormat.Alignment = 1  # 가운데 정렬
sel.ParagraphFormat.LeftIndent = 36  # 0.5인치 (pt)
```

---

## 7. 이미지 처리 (Shapes, InlineShapes)

### 7.1 InlineShapes vs Shapes

- **InlineShapes**: 텍스트 흐름에 포함된 개체 (글자처럼 취급)
- **Shapes**: 자유롭게 배치된 개체 (떠 있는 개체)

### 7.2 InlineShapes 접근

```python
# InlineShapes 수
inline_count = doc.InlineShapes.Count

# InlineShapes 순회
for shape in doc.InlineShapes:
    print(f"Type: {shape.Type}, Width: {shape.Width}, Height: {shape.Height}")

# Type 값
# wdInlineShapePicture = 3
# wdInlineShapeLinkedPicture = 4
# wdInlineShapeEmbeddedOLEObject = 1
```

### 7.3 이미지 삽입

```python
# Selection 위치에 이미지 삽입
sel.InlineShapes.AddPicture(
    FileName=r"C:\path\to\image.jpg",
    LinkToFile=False,
    SaveWithDocument=True
)

# 특정 Range에 이미지 삽입
doc.InlineShapes.AddPicture(
    FileName=r"C:\path\to\image.png",
    LinkToFile=False,
    SaveWithDocument=True,
    Range=doc.Range(0, 0)
)
```

### 7.4 이미지 크기 조정

```python
# 인라인 이미지 크기 조정
shape = doc.InlineShapes(1)
shape.Width = 200   # 포인트 단위
shape.Height = 150

# 비율 유지하여 크기 조정
shape.LockAspectRatio = True  # -1 = True (msoTrue)
shape.Width = 200  # 높이 자동 조정
```

### 7.5 Shapes (떠 있는 개체)

```python
# Shapes 접근
for shape in doc.Shapes:
    print(f"Name: {shape.Name}, Type: {shape.Type}")
    print(f"Position: ({shape.Left}, {shape.Top})")
    print(f"Size: {shape.Width} x {shape.Height}")

# 도형 삽입
doc.Shapes.AddShape(
    Type=1,  # msoShapeRectangle
    Left=100,
    Top=100,
    Width=200,
    Height=100
)
```

---

## 8. 상수(Constants) 참조

### 8.1 상수 접근 방법

```python
# 방법 1: EnsureDispatch 사용 (권장)
import win32com.client as win32
word = win32.gencache.EnsureDispatch('Word.Application')
from win32com.client import constants as c

# 상수 사용
doc.SaveAs(FileName="output.pdf", FileFormat=c.wdFormatPDF)

# 방법 2: 직접 값 사용
doc.SaveAs(FileName="output.pdf", FileFormat=17)
```

### 8.2 주요 상수 목록

#### WdSaveFormat (저장 형식)
| 상수 | 값 |
|------|---|
| wdFormatDocument | 0 |
| wdFormatRTF | 6 |
| wdFormatHTML | 8 |
| wdFormatDocumentDefault | 16 |
| wdFormatPDF | 17 |

#### WdUnits (이동 단위)
| 상수 | 값 |
|------|---|
| wdCharacter | 1 |
| wdWord | 2 |
| wdSentence | 3 |
| wdParagraph | 4 |
| wdLine | 5 |
| wdStory | 6 |
| wdCell | 12 |
| wdColumn | 9 |
| wdRow | 10 |
| wdTable | 15 |

#### WdParagraphAlignment (정렬)
| 상수 | 값 |
|------|---|
| wdAlignParagraphLeft | 0 |
| wdAlignParagraphCenter | 1 |
| wdAlignParagraphRight | 2 |
| wdAlignParagraphJustify | 3 |

#### WdFindWrap (찾기 옵션)
| 상수 | 값 |
|------|---|
| wdFindStop | 0 |
| wdFindContinue | 1 |
| wdFindAsk | 2 |

#### WdReplace (바꾸기 옵션)
| 상수 | 값 |
|------|---|
| wdReplaceNone | 0 |
| wdReplaceOne | 1 |
| wdReplaceAll | 2 |

#### WdLineSpacing (줄 간격)
| 상수 | 값 |
|------|---|
| wdLineSpaceSingle | 0 |
| wdLineSpace1pt5 | 1 |
| wdLineSpaceDouble | 2 |
| wdLineSpaceAtLeast | 3 |
| wdLineSpaceExactly | 4 |
| wdLineSpaceMultiple | 5 |

#### WdCellVerticalAlignment (셀 세로 정렬)
| 상수 | 값 |
|------|---|
| wdCellAlignVerticalTop | 0 |
| wdCellAlignVerticalCenter | 1 |
| wdCellAlignVerticalBottom | 3 |

---

## 9. 단위 변환

### 9.1 Word 기본 단위

Word API에서 대부분의 크기/위치 값은 **포인트(pt)** 단위를 사용합니다.

### 9.2 단위 변환 공식

```python
# 포인트 변환
def pt_to_inch(pt): return pt / 72
def inch_to_pt(inch): return inch * 72

def pt_to_cm(pt): return pt / 72 * 2.54
def cm_to_pt(cm): return cm / 2.54 * 72

def pt_to_mm(pt): return pt / 72 * 25.4
def mm_to_pt(mm): return mm / 25.4 * 72

# Twips 변환 (일부 API에서 사용)
def pt_to_twips(pt): return pt * 20
def twips_to_pt(twips): return twips / 20
```

### 9.3 Word Application 변환 메서드

```python
# InchesToPoints, CentimetersToPoints, etc.
pt_value = word.InchesToPoints(1)     # 1인치 = 72pt
pt_value = word.CentimetersToPoints(2.54)  # 2.54cm = 72pt
pt_value = word.MillimetersToPoints(25.4)  # 25.4mm = 72pt

# 역변환
inch_value = word.PointsToInches(72)
cm_value = word.PointsToCentimeters(72)
```

---

## 10. HWP vs Word 비교

### 10.1 객체 모델 비교

| 기능 | HWP | Word |
|------|-----|------|
| Application | `HWPFrame.HwpObject` | `Word.Application` |
| 문서 | `hwp` (직접) | `word.Documents` |
| 현재 위치 | `hwp.GetPos()` | `word.Selection` |
| 텍스트 추출 | `hwp.GetText()` | `doc.Content.Text` |
| 문단 | `hwp.Paragraphs` (제한적) | `doc.Paragraphs` |
| 표 | `hwp.HeadCtrl` 순회 | `doc.Tables` |
| 서식 | `hwp.CharShape`, `hwp.ParaShape` | `Range.Font`, `Range.ParagraphFormat` |

### 10.2 단위 비교

| 단위 | HWP (HWPUNIT) | Word (pt) |
|------|--------------|-----------|
| 1 pt | 100 | 1 |
| 1 inch | 7,200 | 72 |
| 1 cm | 약 2,834.6 | 약 28.35 |

### 10.3 색상 형식 비교

| 항목 | HWP | Word |
|------|-----|------|
| 형식 | COLORREF (0x00BBGGRR) | RGB 정수 |
| 빨강 | `0x000000FF` | `255` 또는 `0x0000FF` |
| 초록 | `0x0000FF00` | `65280` 또는 `0x00FF00` |
| 파랑 | `0x00FF0000` | `16711680` 또는 `0xFF0000` |

### 10.4 변환 함수

```python
def hwpunit_to_word_pt(hwpunit):
    """HWPUNIT을 Word 포인트로 변환"""
    return hwpunit / 100

def word_pt_to_hwpunit(pt):
    """Word 포인트를 HWPUNIT으로 변환"""
    return pt * 100

def hwp_color_to_word(hwp_color):
    """HWP COLORREF (BGR) to Word RGB"""
    r = hwp_color & 0xFF
    g = (hwp_color >> 8) & 0xFF
    b = (hwp_color >> 16) & 0xFF
    return r | (g << 8) | (b << 16)

def word_color_to_hwp(word_color):
    """Word RGB to HWP COLORREF (BGR)"""
    r = word_color & 0xFF
    g = (word_color >> 8) & 0xFF
    b = (word_color >> 16) & 0xFF
    return r | (g << 8) | (b << 16)
```

---

## 참고 자료

- [Python and Microsoft Office - Using PyWin32](https://www.blog.pythonlibrary.org/2010/07/16/python-and-microsoft-office-using-pywin32/)
- [Automate Word Document (.docx) With Python-docx And pywin32](https://pythoninoffice.com/automate-docx-with-python/)
- [Microsoft Word VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/word)
- [Document.SaveAs2 method (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.saveas2)
- [WdSaveFormat enumeration (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat)
