"""
한글(HWP) 파일 생성 및 텍스트/테이블 입력 예제
- HWP Automation API를 사용하여 새 문서를 생성하고 텍스트/테이블을 입력합니다.
- Windows 환경에서 한글(한컴오피스)이 설치되어 있어야 합니다.
"""

import win32com.client as win32
from typing import List, Optional


class HwpDocument:
    """한글 문서 자동화 클래스"""

    # HWPUNIT: 1mm = 283.465 HWPUNIT
    HWPUNIT_PER_MM = 283.465

    def __init__(self, visible: bool = True):
        """
        한글 문서 객체를 초기화합니다.

        Args:
            visible: 한글 창 표시 여부
        """
        self.hwp = None
        self.visible = visible

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def open(self):
        """한글 애플리케이션을 시작합니다."""
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # 보안 모듈 등록
        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        # 창 표시 설정
        self.hwp.XHwpWindows.Item(0).Visible = self.visible

    def close(self):
        """한글 애플리케이션을 종료합니다."""
        if self.hwp:
            self.hwp.Quit()
            self.hwp = None

    def new_document(self):
        """새 문서를 생성합니다."""
        self.hwp.HAction.Run("FileNew")

    def insert_text(self, text: str):
        """
        현재 캐럿 위치에 텍스트를 삽입합니다.

        Args:
            text: 삽입할 텍스트
        """
        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
        self.hwp.HParameterSet.HInsertText.Text = text
        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)

    def insert_paragraph(self):
        """문단을 나눕니다 (Enter 키)."""
        self.hwp.HAction.Run("BreakPara")

    def create_table(self, rows: int, cols: int,
                     col_widths: Optional[List[int]] = None,
                     row_heights: Optional[List[int]] = None):
        """
        테이블을 생성합니다.

        Args:
            rows: 행 수
            cols: 열 수
            col_widths: 각 열의 너비 (HWPUNIT 단위, 생략 시 자동)
            row_heights: 각 행의 높이 (HWPUNIT 단위, 생략 시 자동)
        """
        self.hwp.HAction.GetDefault("TableCreation", self.hwp.HParameterSet.HTableCreation.HSet)
        self.hwp.HParameterSet.HTableCreation.Rows = rows
        self.hwp.HParameterSet.HTableCreation.Cols = cols

        # 열 너비 설정
        if col_widths:
            col_width_array = self.hwp.HParameterSet.HTableCreation.ColWidth
            for i, width in enumerate(col_widths):
                col_width_array.SetItem(i, width)

        # 행 높이 설정
        if row_heights:
            row_height_array = self.hwp.HParameterSet.HTableCreation.RowHeight
            for i, height in enumerate(row_heights):
                row_height_array.SetItem(i, height)

        self.hwp.HAction.Execute("TableCreation", self.hwp.HParameterSet.HTableCreation.HSet)

    def move_to_next_cell(self):
        """다음 셀로 이동합니다 (Tab 키)."""
        self.hwp.HAction.Run("TableRightCell")

    def move_to_prev_cell(self):
        """이전 셀로 이동합니다 (Shift+Tab)."""
        self.hwp.HAction.Run("TableLeftCell")

    def move_to_next_row(self):
        """다음 행으로 이동합니다."""
        self.hwp.HAction.Run("TableLowerCell")

    def move_to_prev_row(self):
        """이전 행으로 이동합니다."""
        self.hwp.HAction.Run("TableUpperCell")

    def fill_table(self, data: List[List[str]]):
        """
        테이블에 데이터를 채웁니다.
        테이블이 이미 생성되어 있고 첫 번째 셀에 커서가 있어야 합니다.

        Args:
            data: 2차원 리스트 형태의 테이블 데이터
        """
        for row_idx, row in enumerate(data):
            for col_idx, cell_text in enumerate(row):
                self.insert_text(cell_text)
                # 마지막 열이 아니면 다음 셀로 이동
                if col_idx < len(row) - 1:
                    self.move_to_next_cell()
            # 마지막 행이 아니면 다음 행 첫 번째 셀로 이동
            if row_idx < len(data) - 1:
                self.move_to_next_cell()

    def create_table_with_data(self, data: List[List[str]],
                                col_widths: Optional[List[int]] = None):
        """
        데이터가 채워진 테이블을 생성합니다.

        Args:
            data: 2차원 리스트 형태의 테이블 데이터
            col_widths: 각 열의 너비 (HWPUNIT 단위)
        """
        if not data:
            return

        rows = len(data)
        cols = max(len(row) for row in data)

        self.create_table(rows, cols, col_widths)
        self.fill_table(data)

    def save(self, path: str, format: str = "HWP"):
        """
        문서를 저장합니다.

        Args:
            path: 저장할 파일 경로
            format: 파일 형식 (HWP, HWPX, PDF 등)
        """
        self.hwp.SaveAs(path, format)

    @classmethod
    def mm_to_hwpunit(cls, mm: float) -> int:
        """밀리미터를 HWPUNIT으로 변환합니다."""
        return int(mm * cls.HWPUNIT_PER_MM)


def create_hwp_with_text(save_path: str, text: str) -> bool:
    """
    한글 문서를 생성하고 텍스트를 입력한 후 저장합니다.

    Args:
        save_path: 저장할 파일 경로 (예: "C:/documents/hello.hwp")
        text: 입력할 텍스트

    Returns:
        성공 여부
    """
    hwp = None
    try:
        # 1. 한글 COM 객체 생성 (한글이 자동으로 실행됨)
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        # 2. 보안 모듈 등록 (자동화 사용 시 필요)
        # RegisterModule 메서드로 보안 모듈을 등록해야 파일 저장 등이 가능
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        # 3. 한글 창 표시 설정 (True: 보이기, False: 숨기기)
        hwp.XHwpWindows.Item(0).Visible = True

        # 4. 새 문서 생성
        hwp.HAction.Run("FileNew")

        # 5. 텍스트 삽입
        # InsertText 액션을 사용하여 텍스트 입력
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 6. 파일 저장
        hwp.SaveAs(save_path, "HWP")

        print(f"파일이 성공적으로 저장되었습니다: {save_path}")
        return True

    except Exception as e:
        print(f"오류 발생: {e}")
        return False

    finally:
        # 7. 한글 종료 (필요한 경우)
        if hwp:
            hwp.Quit()


def create_hwp_simple(save_path: str, text: str) -> bool:
    """
    간단한 방식으로 한글 문서를 생성합니다.
    (gencache 대신 Dispatch 사용)

    Args:
        save_path: 저장할 파일 경로
        text: 입력할 텍스트

    Returns:
        성공 여부
    """
    hwp = None
    try:
        # COM 객체 생성
        hwp = win32.Dispatch("HWPFrame.HwpObject")

        # 보안 모듈 등록
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        # 새 문서 생성
        hwp.HAction.Run("FileNew")

        # 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Text", text)
        act.Execute(pset)

        # 파일 저장
        hwp.SaveAs(save_path, "HWP")

        print(f"파일이 성공적으로 저장되었습니다: {save_path}")
        return True

    except Exception as e:
        print(f"오류 발생: {e}")
        return False

    finally:
        if hwp:
            hwp.Quit()


if __name__ == "__main__":
    # 저장할 파일 경로 설정
    output_path = r"C:\hwp_docs\hello.hwp"

    # 예제 1: 간단한 텍스트 문서 생성
    # create_hwp_with_text(output_path, "안녕하세요")

    # 기존 hello.hwp 파일 열고 테이블 추가
    doc = HwpDocument(visible=True)
    doc.open()

    # 기존 파일 열기
    doc.hwp.Open(output_path)

    # 문서 끝으로 이동
    doc.hwp.HAction.Run("MoveDocEnd")
    doc.insert_paragraph()
    doc.insert_paragraph()

    # 5x5 테이블 생성
    doc.create_table(5, 5)

    # 파일 저장
    doc.save(output_path)
    print(f"문서가 저장되었습니다: {output_path}")
    print("한글이 열려 있습니다. 직접 확인하고 닫아주세요.")
