# -*- coding: utf-8 -*-
"""
컨트롤 탐색 모듈

HWP 문서 내 컨트롤(표, 그림, 수식 등)을 탐색하고 정보를 조회합니다.
"""

from typing import Dict, List, Any, Optional


# 컨트롤 ID → 이름 매핑
CTRL_NAMES = {
    'cold': '단',
    'secd': '구역',
    'fn': '각주',
    'en': '미주',
    'tbl': '표',
    'eqed': '수식',
    'atno': '번호넣기',
    'head': '머리말',
    'foot': '꼬리말',
    '%dte': '현재 날짜/시간',
    '%pat': '문서 경로',
    '%mmg': '메일 머지',
    '%xrf': '상호 참조',
    '%clk': '누름틀',
    '%hlk': '하이퍼링크',
    'bokm': '책갈피',
    'idxm': '찾아보기',
    '$pic': '그림',
    'gso': '그리기 개체',
    'form': '양식 개체',
    '+pbt': '명령 단추',
    '+rbt': '라디오 단추',
    '+cbt': '선택 상자',
    '+cob': '콤보 상자',
    '+edt': '입력 상자'
}


def find_ctrl(hwp) -> str:
    """
    현재 커서 위치의 컨트롤 ID 조회

    Args:
        hwp: HWP COM 객체

    Returns:
        str: 컨트롤 ID (없으면 빈 문자열)
    """
    try:
        ctrl_id = hwp.FindCtrl()
        return ctrl_id if ctrl_id else ''
    except Exception as e:
        print(f"FindCtrl 실패: {e}")
        return ''


def get_ctrl_name(ctrl_id: str) -> str:
    """
    컨트롤 ID에 해당하는 이름 반환

    Args:
        ctrl_id: 컨트롤 ID (예: 'tbl', 'gso')

    Returns:
        str: 컨트롤 이름 (예: '표', '그리기 개체')
    """
    return CTRL_NAMES.get(ctrl_id, '알 수 없는 컨트롤')


def get_ctrls_in_cell(hwp, target_list_id: int, target_para_id: int = None) -> List[Dict]:
    """
    특정 셀(list_id)에 속한 컨트롤 찾기

    HeadCtrl에서 시작하여 Next로 순회하며 GetAnchorPos로 위치를 확인합니다.

    Args:
        hwp: HWP COM 객체
        target_list_id: 대상 list_id
        target_para_id: 특정 문단만 필터링 (None=전체)

    Returns:
        list: [{
            'ctrl': 컨트롤 객체,
            'id': 컨트롤 ID,
            'desc': 설명,
            'para': 문단 ID
        }, ...]
    """
    ctrls = []
    ctrl = hwp.HeadCtrl

    while ctrl:
        try:
            anchor = ctrl.GetAnchorPos(0)
            ctrl_list_id = anchor.Item("List")
            ctrl_para_id = anchor.Item("Para")

            if ctrl_list_id == target_list_id:
                if target_para_id is None or ctrl_para_id == target_para_id:
                    ctrl_id = ctrl.CtrlID
                    # secd(구역), cold(단) 제외
                    if ctrl_id and ctrl_id not in ("secd", "cold"):
                        ctrls.append({
                            'ctrl': ctrl,
                            'id': ctrl_id,
                            'desc': getattr(ctrl, 'UserDesc', ctrl_id),
                            'para': ctrl_para_id
                        })
        except:
            pass
        ctrl = ctrl.Next

    return ctrls


def iterate_all_ctrls(hwp) -> List[Dict]:
    """
    문서 내 모든 컨트롤 순회

    Args:
        hwp: HWP COM 객체

    Returns:
        list: [{
            'ctrl': 컨트롤 객체,
            'id': 컨트롤 ID,
            'desc': 설명,
            'list_id': 위치 list_id,
            'para_id': 위치 para_id
        }, ...]
    """
    ctrls = []
    ctrl = hwp.HeadCtrl

    while ctrl:
        try:
            ctrl_id = ctrl.CtrlID
            if ctrl_id and ctrl_id not in ("secd", "cold"):
                anchor = ctrl.GetAnchorPos(0)
                ctrls.append({
                    'ctrl': ctrl,
                    'id': ctrl_id,
                    'desc': getattr(ctrl, 'UserDesc', ctrl_id),
                    'list_id': anchor.Item("List"),
                    'para_id': anchor.Item("Para")
                })
        except:
            pass
        ctrl = ctrl.Next

    return ctrls


class CtrlInfo:
    """컨트롤 정보 클래스"""

    def __init__(self, hwp):
        self.hwp = hwp

    def find_ctrl(self) -> str:
        """현재 커서 위치의 컨트롤 ID 조회"""
        return find_ctrl(self.hwp)

    def get_info(self) -> Dict[str, Any]:
        """현재 위치의 컨트롤 정보 조회"""
        ctrl_id = self.find_ctrl()

        if ctrl_id:
            ctrl_name = CTRL_NAMES.get(ctrl_id, '알 수 없는 컨트롤')
            return {
                'ctrl_id': ctrl_id,
                'ctrl_name': ctrl_name,
                'has_ctrl': True
            }
        else:
            return {
                'ctrl_id': '',
                'ctrl_name': '',
                'has_ctrl': False
            }

    def select_ctrl(self, ctrl_list: str, option: int = 1) -> bool:
        """특정 컨트롤 선택"""
        try:
            self.hwp.SelectCtrl(ctrl_list, option)
            return True
        except Exception as e:
            print(f"SelectCtrl 실패: {e}")
            return False

    def print_info(self):
        """컨트롤 정보 출력"""
        info = self.get_info()

        print("=== 컨트롤 정보 ===")
        if info['has_ctrl']:
            print(f"컨트롤 ID: {info['ctrl_id']}")
            print(f"컨트롤 이름: {info['ctrl_name']}")
        else:
            print("컨트롤 없음")
