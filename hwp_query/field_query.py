# -*- coding: utf-8 -*-
"""
필드 조회 모듈

HWP 문서 내 필드(셀 필드, 누름틀 필드 등)를 조회합니다.
"""

import json
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Any


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


# 필드 옵션 상수
FIELD_CELL = 1          # 셀 필드
FIELD_CLICKHERE = 2     # 누름틀 필드
FIELD_SELECTION = 4     # 선택 영역 내 필드
FIELD_TEXTBOX = 8       # 글상자 필드

FIELD_PLAIN = 0         # 번호 없이 이름만
FIELD_NUMBER = 1        # {{#}} 형식 일련번호
FIELD_COUNT = 2         # {{#}} 형식 개수


def get_all_fields(hwp, option: int = 0, list_id_to_coord: Dict = None) -> List[FieldInfo]:
    """
    문서 내 모든 필드 조회

    Args:
        hwp: HWP COM 객체
        option: 필드 옵션 (0=모두, 1=셀, 2=누름틀, 8=글상자)
        list_id_to_coord: list_id → (row, col) 맵 (좌표 조회용)

    Returns:
        FieldInfo 리스트
    """
    field_list_str = hwp.GetFieldList(FIELD_NUMBER, option)
    if not field_list_str:
        return []

    field_names = field_list_str.split('\x02')
    fields = []

    for name in field_names:
        if not name:
            continue

        # 필드로 이동하여 위치 정보 획득
        if hwp.MoveToField(name, True, True, False):
            pos = hwp.GetPos()
            list_id = pos[0]

            # 좌표 조회
            coord = None
            if list_id_to_coord:
                coord = list_id_to_coord.get(list_id)

            # 필드 텍스트 획득
            text = hwp.GetFieldText(name) or ""

            field_info = FieldInfo(
                name=name,
                list_id=list_id,
                row=coord[0] if coord else -1,
                col=coord[1] if coord else -1,
                text=text.strip('\x02')
            )
            fields.append(field_info)

    return fields


def get_cell_fields(hwp, list_id_to_coord: Dict = None) -> List[FieldInfo]:
    """셀 필드만 조회"""
    return get_all_fields(hwp, FIELD_CELL, list_id_to_coord)


def get_clickhere_fields(hwp, list_id_to_coord: Dict = None) -> List[FieldInfo]:
    """누름틀 필드만 조회"""
    return get_all_fields(hwp, FIELD_CLICKHERE, list_id_to_coord)


def get_field_by_name(hwp, name: str, list_id_to_coord: Dict = None) -> Optional[FieldInfo]:
    """
    이름으로 필드 조회

    Args:
        hwp: HWP COM 객체
        name: 필드 이름
        list_id_to_coord: list_id → (row, col) 맵

    Returns:
        FieldInfo 또는 None
    """
    if not hwp.FieldExist(name):
        return None

    if hwp.MoveToField(name, True, True, False):
        pos = hwp.GetPos()
        list_id = pos[0]
        coord = list_id_to_coord.get(list_id) if list_id_to_coord else None
        text = hwp.GetFieldText(name) or ""

        return FieldInfo(
            name=name,
            list_id=list_id,
            row=coord[0] if coord else -1,
            col=coord[1] if coord else -1,
            text=text.strip('\x02')
        )
    return None


def get_fields_at_coord(hwp, row: int, col: int, coord_to_list_id: Dict = None) -> List[FieldInfo]:
    """
    특정 셀 좌표의 필드들 조회

    Args:
        hwp: HWP COM 객체
        row: 행 좌표 (0-based)
        col: 열 좌표 (0-based)
        coord_to_list_id: (row, col) → list_id 맵

    Returns:
        FieldInfo 리스트
    """
    if not coord_to_list_id:
        return []

    list_id = coord_to_list_id.get((row, col))
    if list_id is None:
        return []

    # 역방향 맵 생성
    list_id_to_coord = {v: k for k, v in coord_to_list_id.items()}

    all_fields = get_all_fields(hwp, list_id_to_coord=list_id_to_coord)
    return [f for f in all_fields if f.list_id == list_id]


def get_field_text(hwp, name: str) -> str:
    """
    필드 텍스트 조회

    Args:
        hwp: HWP COM 객체
        name: 필드 이름

    Returns:
        str: 필드 텍스트
    """
    text = hwp.GetFieldText(name)
    return text.strip('\x02') if text else ""


def field_exists(hwp, name: str) -> bool:
    """
    필드 존재 여부 확인

    Args:
        hwp: HWP COM 객체
        name: 필드 이름

    Returns:
        bool: 존재하면 True
    """
    return bool(hwp.FieldExist(name))


def get_json_fields(hwp) -> List[Dict]:
    """
    JSON 형식 필드 이름들을 파싱하여 반환

    필드 이름이 JSON 형식인 경우 파싱하여 딕셔너리로 반환합니다.
    예: {"table":"표이름","coord":[row,col],"desc":"","target":"","author":""}

    Args:
        hwp: HWP COM 객체

    Returns:
        list: [{"table": 표이름, "coord": [r,c], ...}, ...]
    """
    field_list = hwp.GetFieldList(FIELD_NUMBER, FIELD_CELL)
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


def get_fields_by_table(hwp, table_name: str) -> List[Dict]:
    """
    특정 표의 JSON 필드들 조회

    Args:
        hwp: HWP COM 객체
        table_name: 표 이름

    Returns:
        list: [{"table": 표이름, "coord": [r,c]}, ...]
    """
    all_fields = get_json_fields(hwp)
    return [f for f in all_fields if f.get('table') == table_name]


def find_field_by_coord(hwp, table: str, row: int, col: int) -> Optional[str]:
    """
    table과 coord만으로 필드명 찾기

    Args:
        hwp: HWP COM 객체
        table: 표 이름
        row: 행 좌표 (0-based)
        col: 열 좌표 (0-based)

    Returns:
        str: 전체 JSON 필드명 (없으면 None)
    """
    for f in get_json_fields(hwp):
        if f.get('table') == table and f.get('coord') == [row, col]:
            return json.dumps(f, ensure_ascii=False, separators=(',', ':'))
    return None


def get_field_value_by_coord(hwp, table: str, row: int, col: int) -> Optional[str]:
    """
    table과 coord로 필드 텍스트 조회

    Args:
        hwp: HWP COM 객체
        table: 표 이름
        row: 행 좌표
        col: 열 좌표

    Returns:
        str: 필드 텍스트 (없으면 None)
    """
    field_name = find_field_by_coord(hwp, table, row, col)
    if field_name:
        return get_field_text(hwp, field_name)
    return None


def get_current_field_name(hwp) -> str:
    """
    현재 커서 위치의 필드 이름 반환

    Args:
        hwp: HWP COM 객체

    Returns:
        str: 필드 이름 (없으면 빈 문자열)
    """
    name = hwp.GetCurFieldName()
    return name if name else ""


def get_current_field_state(hwp) -> Dict:
    """
    현재 커서 위치의 필드 상태 반환

    Args:
        hwp: HWP COM 객체

    Returns:
        dict: {
            'has_name': 필드 이름 존재 여부,
            'field_type': 필드 타입 ('none', 'cell', 'clickhere'),
            'raw': 원본 상태값
        }
    """
    state = hwp.GetCurFieldState
    field_type_map = {0: "none", 1: "cell", 2: "clickhere"}

    return {
        'has_name': bool(state & 0x10),
        'field_type': field_type_map.get(state & 0x0F, "unknown"),
        'raw': state
    }


def get_field_summary(hwp, list_id_to_coord: Dict = None) -> Dict:
    """
    필드 요약 정보 반환

    Args:
        hwp: HWP COM 객체
        list_id_to_coord: list_id → (row, col) 맵

    Returns:
        dict: {
            'total': 전체 필드 수,
            'in_table': 테이블 내 필드 수,
            'outside_table': 테이블 외부 필드 수,
            'fields': FieldInfo 리스트
        }
    """
    all_fields = get_all_fields(hwp, list_id_to_coord=list_id_to_coord)
    cell_fields = [f for f in all_fields if f.row >= 0]

    return {
        'total': len(all_fields),
        'in_table': len(cell_fields),
        'outside_table': len(all_fields) - len(cell_fields),
        'fields': all_fields
    }
