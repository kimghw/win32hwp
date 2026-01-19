# 테이블 필드 관리 모듈
# 필드 등록, 조회, 삭제, 변경 + 셀 좌표 연동
import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any
from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL


@dataclass
class FieldInfo:
    """필드 정보를 저장하는 데이터 클래스"""
    name: str                    # 필드 이름
    list_id: int                 # 필드가 위치한 셀의 list_id
    row: int = -1                # 셀 행 좌표 (0-based)
    col: int = -1                # 셀 열 좌표 (0-based)
    text: str = ""               # 필드에 입력된 텍스트
    direction: str = ""          # 안내문/지시문
    memo: str = ""               # 도움말/설명
    field_type: str = "cell"     # 필드 타입: "cell", "clickhere", "textbox"


class TableField:
    """테이블 필드 관리 클래스

    테이블 내 필드의 CRUD 작업과 셀 좌표 연동을 제공합니다.
    TableInfo와 연동하여 필드의 (row, col) 좌표를 자동으로 매핑합니다.
    """

    # 필드 옵션 상수
    FIELD_CELL = 1          # 셀 필드
    FIELD_CLICKHERE = 2     # 누름틀 필드
    FIELD_SELECTION = 4     # 선택 영역 내 필드
    FIELD_TEXTBOX = 8       # 글상자 필드

    # 필드 번호 옵션
    FIELD_PLAIN = 0         # 번호 없이 이름만
    FIELD_NUMBER = 1        # {{#}} 형식 일련번호
    FIELD_COUNT = 2         # {{#}} 형식 개수

    def __init__(self, hwp=None, debug: bool = False):
        self.hwp = hwp or get_hwp_instance()
        self.debug = debug
        self.table_info = TableInfo(self.hwp, debug=debug)
        self._fields: Dict[str, List[FieldInfo]] = {}  # 필드명 → FieldInfo 리스트
        self._coord_map: Dict[Tuple[int, int], int] = {}  # (row, col) → list_id
        self._list_id_to_coord: Dict[int, Tuple[int, int]] = {}  # list_id → (row, col)

    def _log(self, msg: str):
        """디버그 로그 출력"""
        if self.debug:
            print(f"[TableField] {msg}")

    # =========================================================================
    # 테이블 진입 및 좌표 맵 초기화
    # =========================================================================

    def enter_table(self, table_index: int = 0) -> bool:
        """테이블에 진입하고 좌표 맵을 초기화"""
        if not self.table_info.enter_table(table_index):
            self._log(f"테이블 {table_index} 진입 실패")
            return False

        # 셀 정보 수집 및 좌표 맵 생성
        self.table_info.collect_cells_bfs()
        self._coord_map = self.table_info.build_coordinate_map()

        # 역방향 맵 = 대표 좌표 사용 (병합 셀의 경우 가장 위-왼쪽 좌표)
        self._list_id_to_coord = dict(self.table_info._representative_coords)

        self._log(f"테이블 {table_index} 진입 완료, {len(self._coord_map)}개 좌표 매핑")
        return True

    def get_coord_from_list_id(self, list_id: int) -> Optional[Tuple[int, int]]:
        """list_id로부터 대표 좌표 반환 (병합 셀은 가장 위-왼쪽)"""
        return self._list_id_to_coord.get(list_id)

    def get_representative_coord(self, list_id: int) -> Optional[Tuple[int, int]]:
        """병합 셀의 대표 좌표 반환 (가장 위, 가장 왼쪽)"""
        return self.table_info.get_representative_coord(list_id)

    def get_cell_coords(self, list_id: int) -> List[Tuple[int, int]]:
        """셀이 차지하는 모든 좌표 반환"""
        return self.table_info.get_cell_coords(list_id)

    def get_merge_info(self, list_id: int) -> Dict:
        """셀의 병합 정보 반환"""
        return self.table_info.get_merge_info(list_id)

    def get_list_id_from_coord(self, row: int, col: int) -> Optional[int]:
        """(row, col) 좌표로부터 list_id 반환"""
        return self._coord_map.get((row, col))

    # =========================================================================
    # 필드 조회 (Read)
    # =========================================================================

    def get_all_fields(self, option: int = 0) -> List[FieldInfo]:
        """문서 내 모든 필드 조회

        Args:
            option: 필드 옵션 (0=모두, 1=셀, 2=누름틀, 8=글상자)

        Returns:
            FieldInfo 리스트
        """
        field_list_str = self.hwp.GetFieldList(self.FIELD_NUMBER, option)
        if not field_list_str:
            return []

        field_names = field_list_str.split('\x02')
        fields = []

        for name in field_names:
            if not name:
                continue

            # 필드로 이동하여 위치 정보 획득
            if self.hwp.MoveToField(name, True, True, False):
                pos = self.hwp.GetPos()
                list_id = pos[0]
                coord = self.get_coord_from_list_id(list_id)

                # 필드 텍스트 획득
                text = self.hwp.GetFieldText(name) or ""

                field_info = FieldInfo(
                    name=name,
                    list_id=list_id,
                    row=coord[0] if coord else -1,
                    col=coord[1] if coord else -1,
                    text=text.strip('\x02')
                )
                fields.append(field_info)
                self._log(f"필드 발견: {name} at ({field_info.row}, {field_info.col})")

        return fields

    def get_cell_fields(self) -> List[FieldInfo]:
        """셀 필드만 조회"""
        return self.get_all_fields(self.FIELD_CELL)

    def get_clickhere_fields(self) -> List[FieldInfo]:
        """누름틀 필드만 조회"""
        return self.get_all_fields(self.FIELD_CLICKHERE)

    def get_field_by_name(self, name: str) -> Optional[FieldInfo]:
        """이름으로 필드 조회"""
        if not self.hwp.FieldExist(name):
            return None

        if self.hwp.MoveToField(name, True, True, False):
            pos = self.hwp.GetPos()
            list_id = pos[0]
            coord = self.get_coord_from_list_id(list_id)
            text = self.hwp.GetFieldText(name) or ""

            return FieldInfo(
                name=name,
                list_id=list_id,
                row=coord[0] if coord else -1,
                col=coord[1] if coord else -1,
                text=text.strip('\x02')
            )
        return None

    def get_fields_at_coord(self, row: int, col: int) -> List[FieldInfo]:
        """특정 셀 좌표의 필드들 조회"""
        list_id = self.get_list_id_from_coord(row, col)
        if list_id is None:
            return []

        all_fields = self.get_all_fields()
        return [f for f in all_fields if f.list_id == list_id]

    def get_field_text(self, name: str) -> str:
        """필드 텍스트 조회"""
        text = self.hwp.GetFieldText(name)
        return text.strip('\x02') if text else ""

    def field_exists(self, name: str) -> bool:
        """필드 존재 여부 확인"""
        return bool(self.hwp.FieldExist(name))

    # =========================================================================
    # 필드 등록 (Create)
    # =========================================================================

    def create_field_at_cursor(self, name: str, direction: str = "", memo: str = "") -> bool:
        """현재 커서 위치에 누름틀 필드 생성

        Args:
            name: 필드 이름
            direction: 안내문 (누름틀 내 표시 텍스트)
            memo: 도움말/메모

        Returns:
            성공 여부
        """
        try:
            result = self.hwp.CreateField(direction, memo, name)
            self._log(f"필드 생성: {name} (direction={direction}, memo={memo})")
            return bool(result)
        except Exception as e:
            self._log(f"필드 생성 실패: {e}")
            return False

    def create_field_at_coord(self, row: int, col: int, name: str,
                               direction: str = "", memo: str = "") -> bool:
        """특정 셀 좌표에 필드 생성

        Args:
            row: 행 좌표 (0-based)
            col: 열 좌표 (0-based)
            name: 필드 이름
            direction: 안내문
            memo: 도움말

        Returns:
            성공 여부
        """
        list_id = self.get_list_id_from_coord(row, col)
        if list_id is None:
            self._log(f"좌표 ({row}, {col})에 해당하는 셀 없음")
            return False

        # 셀로 이동
        self.hwp.SetPos(list_id, 0, 0)
        return self.create_field_at_cursor(name, direction, memo)

    def set_cell_field_name(self, row: int, col: int, name: str) -> bool:
        """특정 셀에 필드 이름 설정 (셀 자체를 필드로)

        Args:
            row: 행 좌표
            col: 열 좌표
            name: 필드 이름

        Returns:
            성공 여부
        """
        list_id = self.get_list_id_from_coord(row, col)
        if list_id is None:
            return False

        self.hwp.SetPos(list_id, 0, 0)
        try:
            # 셀 블록 선택 후 필드명 설정
            self.hwp.HAction.Run("TableCellBlock")
            result = self.hwp.SetCurFieldName(name, 1, '', '')  # option=1 (셀 필드)
            self.hwp.HAction.Run("Cancel")
            self._log(f"셀 ({row}, {col}) 필드 이름 설정: {name}")
            return bool(result)
        except Exception as e:
            self.hwp.HAction.Run("Cancel")
            self._log(f"셀 필드 이름 설정 실패: {e}")
            return False

    def set_structured_field_names(self, caption: str = None,
                                    header_rows: int = 1,
                                    footer_rows: int = 0) -> Dict[str, List[str]]:
        """테이블 전체에 구조화된 필드 이름 설정

        필드 이름 형식: {캡션}_{영역}_{row}_{col}
        예: 표1_head_0_0, 표1_body_1_2, 표1_foot_3_0

        Args:
            caption: 캡션 이름 (None이면 테이블 캡션 자동 추출, 없으면 "TBL")
            header_rows: 헤더 행 수 (기본 1)
            footer_rows: 푸터 행 수 (기본 0)

        Returns:
            {'head': [...], 'body': [...], 'foot': [...]} 생성된 필드명 목록
        """
        # 캡션 가져오기
        if caption is None:
            caption = self.table_info.get_table_caption()
            if not caption:
                caption = "TBL"
            else:
                # 캡션에서 특수문자 제거, 공백을 _로
                caption = caption.strip().replace(' ', '_')
                caption = ''.join(c for c in caption if c.isalnum() or c == '_')
                if not caption:
                    caption = "TBL"

        # 테이블 크기
        size = self.table_info.get_table_size()
        total_rows = size['rows']
        total_cols = size['cols']

        if total_rows == 0:
            return {'head': [], 'body': [], 'foot': []}

        # 영역 구분
        # head: 0 ~ header_rows-1
        # foot: total_rows-footer_rows ~ total_rows-1
        # body: 나머지
        result = {'head': [], 'body': [], 'foot': []}

        # 첫 번째 셀로 이동하여 테이블 컨텍스트 확보
        first_list_id = min(self.table_info.cells.keys())
        self.hwp.SetPos(first_list_id, 0, 0)

        for list_id in sorted(self.table_info.cells.keys()):
            rep_coord = self.table_info.get_representative_coord(list_id)
            if rep_coord is None:
                continue

            row, col = rep_coord

            # 영역 판단
            if row < header_rows:
                area = 'head'
            elif footer_rows > 0 and row >= total_rows - footer_rows:
                area = 'foot'
            else:
                area = 'body'

            # 필드 이름 생성
            field_name = f"{caption}_{area}_{row}_{col}"

            # 셀로 이동하여 필드 설정
            try:
                self.hwp.SetPos(list_id, 0, 0)
                self.hwp.HAction.Run("TableCellBlock")
                self.hwp.SetCurFieldName(field_name, 1, '', '')
                self.hwp.HAction.Run("Cancel")
                result[area].append(field_name)
                self._log(f"필드 설정: {field_name}")
            except Exception as e:
                self._log(f"필드 설정 실패 ({field_name}): {e}")
                self.hwp.HAction.Run("Cancel")

        return result

    def set_structured_field_values(self, caption: str = None,
                                     header_rows: int = 1,
                                     footer_rows: int = 0,
                                     show_coords: bool = True) -> Dict[str, List[str]]:
        """테이블 전체에 구조화된 필드 이름 설정 + 좌표값 입력

        Args:
            caption: 캡션 이름
            header_rows: 헤더 행 수
            footer_rows: 푸터 행 수
            show_coords: True면 좌표를 셀 값으로 입력

        Returns:
            {'head': [...], 'body': [...], 'foot': [...]} 생성된 필드명 목록
        """
        # 먼저 필드 이름 설정
        result = self.set_structured_field_names(caption, header_rows, footer_rows)

        if not show_coords:
            return result

        # 좌표 값 입력
        for list_id in sorted(self.table_info.cells.keys()):
            rep_coord = self.table_info.get_representative_coord(list_id)
            if rep_coord is None:
                continue

            row, col = rep_coord
            all_coords = self.table_info.get_cell_coords(list_id)

            # 영역 판단
            size = self.table_info.get_table_size()
            total_rows = size['rows']
            if row < header_rows:
                area = 'head'
            elif footer_rows > 0 and row >= total_rows - footer_rows:
                area = 'foot'
            else:
                area = 'body'

            # 캡션 가져오기
            if caption is None:
                cap = self.table_info.get_table_caption() or "TBL"
                cap = cap.strip().replace(' ', '_')
                cap = ''.join(c for c in cap if c.isalnum() or c == '_') or "TBL"
            else:
                cap = caption

            field_name = f"{cap}_{area}_{row}_{col}"

            # 좌표 문자열
            if len(all_coords) > 1:
                min_r = min(r for r, c in all_coords)
                max_r = max(r for r, c in all_coords)
                min_c = min(c for r, c in all_coords)
                max_c = max(c for r, c in all_coords)
                coord_str = f"({min_r},{min_c})~({max_r},{max_c})"
            else:
                coord_str = f"({row},{col})"

            self.hwp.PutFieldText(field_name, coord_str)

        return result

    def set_table_fields_json(self, table_name: str = None) -> Dict[str, Any]:
        """테이블 셀에 JSON 필드 이름 설정

        필드 이름 형식: {"table":"표이름","coord":[row,col],"desc":"","target":"","author":""}

        Args:
            table_name: 표 이름 (None이면 캡션 자동 추출, 없으면 "TBL")

        Returns:
            {'table': 표이름, 'fields': [필드명JSON, ...]}
        """
        # 표 이름 결정
        if table_name is None:
            table_name = self.table_info.get_table_caption()
            if not table_name:
                table_name = "TBL"
            else:
                # 캡션 정리 (공백/특수문자 유지, 앞뒤 공백만 제거)
                table_name = table_name.strip()
                if not table_name:
                    table_name = "TBL"

        result = {'table': table_name, 'fields': []}

        if not self.table_info.cells:
            return result

        first_list_id = min(self.table_info.cells.keys())
        self.hwp.SetPos(first_list_id, 0, 0)

        for list_id in sorted(self.table_info.cells.keys()):
            rep_coord = self.table_info.get_representative_coord(list_id)
            if rep_coord is None:
                continue

            row, col = rep_coord

            # JSON 필드 이름 생성
            field_data = {"table": table_name, "coord": [row, col], "desc": "", "target": "", "author": ""}
            field_name = json.dumps(field_data, ensure_ascii=False, separators=(',', ':'))

            try:
                self.hwp.SetPos(list_id, 0, 0)
                self.hwp.HAction.Run("TableCellBlock")
                self.hwp.SetCurFieldName(field_name, 1, '', '')
                self.hwp.HAction.Run("Cancel")

                result['fields'].append(field_name)
                self._log(f"JSON 필드 설정: {table_name} ({row},{col})")
            except Exception as e:
                self._log(f"필드 설정 실패: {e}")
                self.hwp.HAction.Run("Cancel")

        return result

    def get_json_fields(self) -> List[Dict]:
        """JSON 형식 필드 이름들을 파싱하여 반환

        Returns:
            [{"table": 표이름, "coord": [r,c], "desc": "", "target": "", "author": ""}, ...]
        """
        field_list = self.hwp.GetFieldList(1, 1)  # 셀 필드
        if not field_list:
            return []

        result = []
        for f in field_list.split('\x02'):
            # {{인덱스}} 제거
            name = f.split('{{')[0] if '{{' in f else f
            try:
                data = json.loads(name)
                if 'table' in data and 'coord' in data:
                    result.append(data)
            except json.JSONDecodeError:
                pass

        return result

    def get_fields_by_table(self, table_name: str) -> List[Dict]:
        """특정 표의 JSON 필드들 조회

        Args:
            table_name: 표 이름

        Returns:
            [{"table": 표이름, "coord": [r,c]}, ...]
        """
        all_fields = self.get_json_fields()
        return [f for f in all_fields if f.get('table') == table_name]

    def make_field_name(self, table: str, row: int, col: int,
                        desc: str = "", target: str = "", author: str = "") -> str:
        """table, coord로 JSON 필드명 생성

        Args:
            table: 표 이름
            row: 행 좌표 (0-based)
            col: 열 좌표 (0-based)
            desc: 설명 (선택)
            target: 타겟 (선택)
            author: 작성자 (선택)

        Returns:
            JSON 형식 필드명
        """
        data = {
            "table": table,
            "coord": [row, col],
            "desc": desc,
            "target": target,
            "author": author
        }
        return json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    def find_field_by_coord(self, table: str, row: int, col: int) -> Optional[str]:
        """table과 coord만으로 필드명 찾기

        문서에서 해당 table, coord를 가진 필드를 검색하여 전체 필드명 반환

        Args:
            table: 표 이름
            row: 행 좌표 (0-based)
            col: 열 좌표 (0-based)

        Returns:
            전체 JSON 필드명 (없으면 None)
        """
        for f in self.get_json_fields():
            if f.get('table') == table and f.get('coord') == [row, col]:
                return json.dumps(f, ensure_ascii=False, separators=(',', ':'))
        return None

    def get_field_value_by_coord(self, table: str, row: int, col: int) -> Optional[str]:
        """table과 coord로 필드 텍스트 조회

        Args:
            table: 표 이름
            row: 행 좌표
            col: 열 좌표

        Returns:
            필드 텍스트 (없으면 None)
        """
        field_name = self.find_field_by_coord(table, row, col)
        if field_name:
            return self.get_field_text(field_name)
        return None

    def set_field_value_by_coord(self, table: str, row: int, col: int, text: str) -> bool:
        """table과 coord로 필드 텍스트 입력

        Args:
            table: 표 이름
            row: 행 좌표
            col: 열 좌표
            text: 입력할 텍스트

        Returns:
            성공 여부
        """
        field_name = self.find_field_by_coord(table, row, col)
        if field_name:
            self.put_field_text(field_name, text)
            return True
        return False

    # =========================================================================
    # 필드 수정 (Update)
    # =========================================================================

    def put_field_text(self, name: str, text: str) -> None:
        """필드에 텍스트 입력

        Args:
            name: 필드 이름 (동일 이름 모든 필드에 적용)
            text: 입력할 텍스트
        """
        self.hwp.PutFieldText(name, text)
        self._log(f"필드 텍스트 입력: {name} = {text}")

    def put_field_text_by_index(self, name: str, index: int, text: str) -> None:
        """특정 인덱스의 필드에 텍스트 입력

        Args:
            name: 필드 이름
            index: 필드 인덱스 (0-based)
            text: 입력할 텍스트
        """
        field_ref = f"{name}{{{{{index}}}}}"
        self.hwp.PutFieldText(field_ref, text)
        self._log(f"필드 텍스트 입력: {field_ref} = {text}")

    def put_fields_text(self, field_text_map: Dict[str, str]) -> None:
        """여러 필드에 텍스트 일괄 입력

        Args:
            field_text_map: {필드이름: 텍스트} 딕셔너리
        """
        if not field_text_map:
            return

        field_list = '\x02'.join(field_text_map.keys())
        text_list = '\x02'.join(field_text_map.values())
        self.hwp.PutFieldText(field_list, text_list)
        self._log(f"필드 일괄 입력: {len(field_text_map)}개")

    def rename_field(self, old_name: str, new_name: str) -> None:
        """필드 이름 변경

        Args:
            old_name: 기존 이름
            new_name: 새 이름
        """
        self.hwp.RenameField(old_name, new_name)
        self._log(f"필드 이름 변경: {old_name} → {new_name}")

    def rename_fields(self, rename_map: Dict[str, str]) -> None:
        """여러 필드 이름 일괄 변경

        Args:
            rename_map: {기존이름: 새이름} 딕셔너리
        """
        if not rename_map:
            return

        old_names = '\x02'.join(rename_map.keys())
        new_names = '\x02'.join(rename_map.values())
        self.hwp.RenameField(old_names, new_names)
        self._log(f"필드 이름 일괄 변경: {len(rename_map)}개")

    def modify_field_editable(self, name: str, editable: bool) -> int:
        """필드의 편집 가능 속성 변경

        Args:
            name: 필드 이름
            editable: True=편집가능, False=읽기전용

        Returns:
            결과 코드 (음수=에러)
        """
        if editable:
            result = self.hwp.ModifyFieldProperties(name, 0, 1)
        else:
            result = self.hwp.ModifyFieldProperties(name, 1, 0)
        self._log(f"필드 편집 속성 변경: {name}, editable={editable}, result={result}")
        return result

    # =========================================================================
    # 필드 삭제 (Delete)
    # =========================================================================

    def delete_field_at_cursor(self) -> bool:
        """현재 커서 위치의 필드 삭제 (내용은 유지)"""
        try:
            self.hwp.HAction.Run("DeleteField")
            self._log("현재 위치 필드 삭제")
            return True
        except Exception as e:
            self._log(f"필드 삭제 실패: {e}")
            return False

    def delete_field_by_name(self, name: str) -> bool:
        """이름으로 필드 삭제"""
        if not self.hwp.MoveToField(name, True, True, False):
            self._log(f"필드 {name} 찾을 수 없음")
            return False
        return self.delete_field_at_cursor()

    def delete_field_at_coord(self, row: int, col: int) -> bool:
        """특정 좌표의 필드 삭제"""
        fields = self.get_fields_at_coord(row, col)
        if not fields:
            return False

        for f in fields:
            self.delete_field_by_name(f.name)
        return True

    def clear_field_text(self, name: str) -> None:
        """필드 텍스트 비우기"""
        self.put_field_text(name, "")

    # =========================================================================
    # 필드 이동 (Navigation)
    # =========================================================================

    def move_to_field(self, name: str, select: bool = False) -> bool:
        """필드로 이동

        Args:
            name: 필드 이름
            select: True=내용 선택, False=커서만 이동

        Returns:
            성공 여부
        """
        return bool(self.hwp.MoveToField(name, True, True, select))

    def move_to_field_by_coord(self, row: int, col: int) -> bool:
        """좌표의 첫 번째 필드로 이동"""
        fields = self.get_fields_at_coord(row, col)
        if not fields:
            return False
        return self.move_to_field(fields[0].name)

    # =========================================================================
    # 유틸리티
    # =========================================================================

    def get_current_field_name(self) -> str:
        """현재 커서 위치의 필드 이름 반환"""
        name = self.hwp.GetCurFieldName()
        return name if name else ""

    def get_current_field_state(self) -> Dict:
        """현재 커서 위치의 필드 상태 반환"""
        state = self.hwp.GetCurFieldState
        field_type_map = {0: "none", 1: "cell", 2: "clickhere"}

        return {
            'has_name': bool(state & 0x10),
            'field_type': field_type_map.get(state & 0x0F, "unknown"),
            'raw': state
        }

    def set_field_view_option(self, option: int) -> int:
        """필드 표시 옵션 설정

        Args:
            option: 1=숨김, 2=빨간색(기본), 3=흰색

        Returns:
            설정된 옵션
        """
        return self.hwp.SetFieldViewOption(option)

    def get_table_size(self) -> Dict[str, int]:
        """현재 테이블 크기 반환"""
        return self.table_info.get_table_size()

    def get_field_summary(self) -> Dict:
        """필드 요약 정보 반환"""
        all_fields = self.get_all_fields()
        cell_fields = [f for f in all_fields if f.row >= 0]

        return {
            'total': len(all_fields),
            'in_table': len(cell_fields),
            'outside_table': len(all_fields) - len(cell_fields),
            'fields': all_fields
        }

    def print_field_map(self):
        """테이블 필드 맵 출력 (디버그용)"""
        table_size = self.get_table_size()
        rows, cols = table_size['rows'], table_size['cols']

        # 좌표별 필드 수집
        field_map = {}
        for f in self.get_all_fields():
            if f.row >= 0:
                key = (f.row, f.col)
                if key not in field_map:
                    field_map[key] = []
                field_map[key].append(f.name)

        print(f"\n테이블 필드 맵 ({rows}x{cols}):")
        print("-" * 60)
        for r in range(rows):
            row_str = f"행{r}: "
            for c in range(cols):
                fields = field_map.get((r, c), [])
                if fields:
                    row_str += f"[{','.join(fields)}] "
                else:
                    row_str += "[    ] "
            print(row_str)
        print("-" * 60)


# =============================================================================
# 모듈 테스트
# =============================================================================
if __name__ == "__main__":
    print("TableField 모듈 테스트")
    print("=" * 60)

    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
        sys.exit(1)

    tf = TableField(hwp, debug=True)

    # 첫 번째 테이블 진입
    if tf.enter_table(0):
        print("\n[테이블 정보]")
        size = tf.get_table_size()
        print(f"  크기: {size['rows']}행 x {size['cols']}열")

        print("\n[필드 조회]")
        summary = tf.get_field_summary()
        print(f"  총 필드: {summary['total']}개")
        print(f"  테이블 내: {summary['in_table']}개")

        for f in summary['fields']:
            print(f"  - {f.name}: ({f.row}, {f.col}), text='{f.text}'")

        tf.print_field_map()
    else:
        print("테이블을 찾을 수 없습니다.")
