# -*- coding: utf-8 -*-
"""
한글(HWP) 테이블 관리 클래스

기능:
1. 문서 내 모든 표 조회
2. 표에 고유 ID 부여 (3행 1열 셀에 필드로 삽입)
3. 표 정보를 YAML로 저장/로드
"""

import win32com.client as win32
import yaml
import os
import uuid


class TableManager:
    """한글 테이블 관리 클래스"""

    def __init__(self, hwp):
        """
        Args:
            hwp: 한글 인스턴스
        """
        self.hwp = hwp
        self.tables = {}  # {table_id: table_info}

    def _save_pos(self):
        """현재 위치 저장"""
        return self.hwp.GetPos()

    def _restore_pos(self, pos):
        """위치 복원"""
        self.hwp.SetPos(pos[0], pos[1], pos[2])

    def scan_tables(self):
        """
        문서 내 모든 표 조회

        Returns:
            list: 표 정보 목록
        """
        tables = []
        ctrl = self.hwp.HeadCtrl
        table_index = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                anchor = ctrl.GetAnchorPos(0)
                para = anchor.Item("Para")
                pos = anchor.Item("Pos")

                # 표 속성 가져오기
                table_info = {
                    'index': table_index,
                    'para': para,
                    'pos': pos,
                    'user_desc': ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else '',
                    'ctrl': ctrl  # 컨트롤 참조 (저장 시 제외)
                }
                tables.append(table_info)
                table_index += 1

            ctrl = ctrl.Next

        return tables

    def get_table_by_index(self, index):
        """
        인덱스로 표 컨트롤 가져오기

        Args:
            index: 표 순번 (0부터)

        Returns:
            ctrl: 표 컨트롤 또는 None
        """
        ctrl = self.hwp.HeadCtrl
        table_index = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                if table_index == index:
                    return ctrl
                table_index += 1
            ctrl = ctrl.Next

        return None

    def move_to_cell(self, table_ctrl, row, col):
        """
        표의 특정 셀로 이동

        Args:
            table_ctrl: 표 컨트롤
            row: 행 번호 (0부터)
            col: 열 번호 (0부터)

        Returns:
            bool: 성공 여부
        """
        # 표 위치로 이동
        anchor = table_ctrl.GetAnchorPos(0)
        self.hwp.SetPos(0, anchor.Item("Para"), anchor.Item("Pos"))

        # 표 안으로 진입 (첫 셀로 이동됨)
        self.hwp.HAction.Run("ShapeObjTableSelCell")

        # 첫 행으로 이동 (충분히 많이 위로)
        for _ in range(100):
            self.hwp.HAction.Run("TableUpperCell")

        # 첫 열로 이동 (충분히 많이 왼쪽으로)
        for _ in range(100):
            self.hwp.HAction.Run("TableLeftCell")

        # 지정된 행으로 이동
        for _ in range(row):
            self.hwp.HAction.Run("TableLowerCell")

        # 지정된 열로 이동
        for _ in range(col):
            self.hwp.HAction.Run("TableRightCell")

        return True

    def insert_field_id(self, table_ctrl, table_id, row=2, col=0):
        """
        표의 특정 셀에 필드 ID 삽입 (누름틀 사용)

        Args:
            table_ctrl: 표 컨트롤
            table_id: 부여할 ID
            row: 행 번호 (0부터, 기본값 2 = 3행)
            col: 열 번호 (0부터, 기본값 0 = 1열)

        Returns:
            bool: 성공 여부
        """
        saved = self._save_pos()

        try:
            # 해당 셀로 이동
            self.move_to_cell(table_ctrl, row, col)

            # 셀 시작으로 이동
            self.hwp.HAction.Run("MoveParaBegin")

            # 누름틀 생성 (CreateField 메서드 사용)
            # CreateField(direction, memo, name)
            field_name = f"TBL_{table_id}"
            result = self.hwp.CreateField(
                "",           # direction (안내문)
                "",           # memo (설명)
                field_name    # name (필드명)
            )

            # 표 밖으로 나가기
            self.hwp.HAction.Run("Cancel")
            self.hwp.HAction.Run("Cancel")

            return result

        except Exception as e:
            print(f"필드 삽입 실패: {e}")
            return False

        finally:
            self._restore_pos(saved)

    def get_next_id(self):
        """
        다음 순차 ID 반환 (기존 테이블 중 최대값 + 1)

        Returns:
            int: 다음 ID
        """
        if not self.tables:
            return 1

        max_id = 0
        for table_id in self.tables.keys():
            try:
                num = int(table_id)
                if num > max_id:
                    max_id = num
            except ValueError:
                continue

        return max_id + 1

    def register_all_tables(self):
        """
        모든 표에 고유 ID 부여 및 등록

        Returns:
            dict: 등록된 표 정보
        """
        tables = self.scan_tables()

        # 기존 테이블 정보 유지, 다음 ID부터 시작
        next_id = self.get_next_id()

        for i, table in enumerate(tables):
            table_id = str(next_id + i)  # 순차 번호

            # 필드 삽입
            success = self.insert_field_id(table['ctrl'], table_id)

            if success:
                self.tables[table_id] = {
                    'id': table_id,
                    'index': table['index'],
                    'para': table['para'],
                    'pos': table['pos'],
                    'user_desc': table['user_desc']
                }
                print(f"표 {table_id} 등록 완료 (Para={table['para']})")

        return self.tables

    def save_to_yaml(self, filepath):
        """
        표 정보를 YAML로 저장

        Args:
            filepath: 저장 경로
        """
        data = {
            'tables': self.tables,
            'count': len(self.tables)
        }

        with open(filepath, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)

        print(f"저장 완료: {filepath}")

    def load_from_yaml(self, filepath):
        """
        YAML에서 표 정보 로드

        Args:
            filepath: 파일 경로

        Returns:
            dict: 표 정보
        """
        if not os.path.exists(filepath):
            print(f"파일 없음: {filepath}")
            return {}

        with open(filepath, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)

        self.tables = data.get('tables', {})
        print(f"로드 완료: {len(self.tables)}개 표")
        return self.tables

    def find_table_by_id(self, table_id):
        """
        ID로 표 정보 찾기

        Args:
            table_id: 표 ID

        Returns:
            dict: 표 정보 또는 None
        """
        return self.tables.get(table_id)

    def get_cell_text(self, table_ctrl, row, col):
        """
        표의 특정 셀 텍스트 가져오기

        Args:
            table_ctrl: 표 컨트롤
            row: 행 번호 (0부터)
            col: 열 번호 (0부터)

        Returns:
            str: 셀 텍스트
        """
        saved = self._save_pos()

        try:
            # 표 위치로 이동 후 진입
            anchor = table_ctrl.GetAnchorPos(0)
            self.hwp.SetPos(0, anchor.Item("Para"), anchor.Item("Pos"))

            # 표 안으로 진입
            self.hwp.HAction.Run("ShapeObjTableSelCell")

            # 표의 첫 셀(1,1)로 이동
            # moveStartOfCell(104) = 행의 시작, moveTopOfCell(106) = 열의 시작
            self.hwp.MovePos(104)  # 행의 시작 (첫 열로)
            self.hwp.MovePos(106)  # 열의 시작 (첫 행으로)

            # 지정된 행으로 이동
            for _ in range(row):
                self.hwp.MovePos(103)  # moveDownOfCell

            # 지정된 열로 이동
            for _ in range(col):
                self.hwp.MovePos(101)  # moveRightOfCell

            # 셀 전체 선택
            self.hwp.HAction.Run("MoveParaBegin")
            self.hwp.HAction.Run("MoveSelParaEnd")

            text = self.hwp.GetTextFile("TEXT", "saveblock")
            self.hwp.HAction.Run("Cancel")

            if text:
                text = text.replace('\r\n', '').replace('\r', '').replace('\n', '')

            return text or ""

        finally:
            self._restore_pos(saved)


if __name__ == "__main__":
    # 테스트
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.Open(r"d:\hwp_docs\test.hwp", "HWP", "forceopen:true")
    hwp.EditMode = 1

    manager = TableManager(hwp)

    print("=" * 50)
    print("  테이블 관리자 테스트")
    print("=" * 50)
    print()

    # 1. 표 스캔
    print("1. 문서 내 표 스캔...")
    tables = manager.scan_tables()
    print(f"   발견된 표: {len(tables)}개")
    for t in tables:
        print(f"   - 표 {t['index']}: Para={t['para']}, Pos={t['pos']}")
    print()

    if tables:
        # 2. 모든 표에 ID 부여
        print("2. 표에 ID 필드 삽입...")
        manager.register_all_tables()
        print()

        # 3. YAML로 저장
        yaml_path = r"d:\hwp_docs\tables.yaml"
        print(f"3. YAML 저장: {yaml_path}")
        manager.save_to_yaml(yaml_path)
        print()

        # 4. 저장된 내용 출력
        print("4. 저장된 표 정보:")
        for table_id, info in manager.tables.items():
            print(f"   TABLE_ID_{table_id}: Para={info['para']}, Pos={info['pos']}")

    print("\n테스트 완료")
