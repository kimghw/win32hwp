# -*- coding: utf-8 -*-
"""HWP → Excel 변환 설정 관리 모듈"""

import os
from dataclasses import dataclass, field as dataclass_field
from typing import Optional, Dict, Any

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False


# =============================================================================
# 설정 데이터클래스
# =============================================================================

@dataclass
class PageMarginsConfig:
    enabled: bool = True
    left: bool = True
    right: bool = True
    top: bool = True
    bottom: bool = True


@dataclass
class FitToPageConfig:
    enabled: bool = True
    width: int = 1
    height: int = 0


@dataclass
class PageConfig:
    enabled: bool = True
    margins: PageMarginsConfig = dataclass_field(default_factory=PageMarginsConfig)
    orientation: bool = True
    fit_to_page: FitToPageConfig = dataclass_field(default_factory=FitToPageConfig)


@dataclass
class BackgroundConfig:
    enabled: bool = True


@dataclass
class BorderConfig:
    enabled: bool = True
    style: str = "thin"


@dataclass
class FontConfig:
    enabled: bool = True
    name: bool = True
    size: bool = True
    bold: bool = True
    italic: bool = True
    color: bool = True
    underline: bool = False
    strikeout: bool = False


@dataclass
class AlignmentConfig:
    enabled: bool = True
    horizontal: bool = True
    vertical: bool = True
    wrap_text: bool = True


@dataclass
class StyleConfig:
    enabled: bool = True
    background: BackgroundConfig = dataclass_field(default_factory=BackgroundConfig)
    border: BorderConfig = dataclass_field(default_factory=BorderConfig)
    font: FontConfig = dataclass_field(default_factory=FontConfig)
    alignment: AlignmentConfig = dataclass_field(default_factory=AlignmentConfig)


@dataclass
class RowHeightConfig:
    enabled: bool = True
    min: int = 12
    max: int = 409


@dataclass
class ColWidthConfig:
    enabled: bool = True
    min: int = 2
    max: int = 255


@dataclass
class SizeConfig:
    enabled: bool = True
    row_height: RowHeightConfig = dataclass_field(default_factory=RowHeightConfig)
    col_width: ColWidthConfig = dataclass_field(default_factory=ColWidthConfig)


@dataclass
class FieldPatternConfig:
    enabled: bool = True
    left_format: str = "{A}_"
    top_format: str = "_{B}"
    both_format: str = "{A}_{B}"
    tolerance: int = 50


@dataclass
class FieldRandomConfig:
    length: int = 12


@dataclass
class FieldNamingConfig:
    use_bookmark: bool = True
    pattern: FieldPatternConfig = dataclass_field(default_factory=FieldPatternConfig)
    random: FieldRandomConfig = dataclass_field(default_factory=FieldRandomConfig)


@dataclass
class FieldTargetConfig:
    no_background_only: bool = True


@dataclass
class FieldConfig:
    enabled: bool = True
    set_hwp_field: bool = True
    naming: FieldNamingConfig = dataclass_field(default_factory=FieldNamingConfig)
    target: FieldTargetConfig = dataclass_field(default_factory=FieldTargetConfig)


@dataclass
class LockRulesConfig:
    with_background: bool = True
    without_background: bool = False


@dataclass
class ProtectionAllowConfig:
    format_cells: bool = True
    format_columns: bool = True
    format_rows: bool = True
    insert_columns: bool = False
    insert_rows: bool = False
    delete_columns: bool = False
    delete_rows: bool = False


@dataclass
class ProtectionConfig:
    enabled: bool = True
    lock_rules: LockRulesConfig = dataclass_field(default_factory=LockRulesConfig)
    allow: ProtectionAllowConfig = dataclass_field(default_factory=ProtectionAllowConfig)


@dataclass
class OutputSheetsConfig:
    main: bool = True
    page_info: bool = True
    cell_info: bool = True
    size_info: bool = True


@dataclass
class OutputSuffixConfig:
    page: str = "_page"
    cells: str = "_cells"
    sizes: str = "_sizes"


@dataclass
class OutputConfig:
    sheets: OutputSheetsConfig = dataclass_field(default_factory=OutputSheetsConfig)
    suffix: OutputSuffixConfig = dataclass_field(default_factory=OutputSuffixConfig)


@dataclass
class ExportConfig:
    """전체 변환 설정"""
    page: PageConfig = dataclass_field(default_factory=PageConfig)
    style: StyleConfig = dataclass_field(default_factory=StyleConfig)
    size: SizeConfig = dataclass_field(default_factory=SizeConfig)
    field: FieldConfig = dataclass_field(default_factory=FieldConfig)
    protection: ProtectionConfig = dataclass_field(default_factory=ProtectionConfig)
    output: OutputConfig = dataclass_field(default_factory=OutputConfig)


# =============================================================================
# 설정 로드/저장 함수
# =============================================================================

def get_default_config() -> ExportConfig:
    """기본 설정 반환"""
    return ExportConfig()


def load_config(config_path: str) -> ExportConfig:
    """YAML 설정 파일 로드

    Args:
        config_path: 설정 파일 경로

    Returns:
        ExportConfig 객체
    """
    if not HAS_YAML:
        print("[경고] PyYAML이 설치되지 않았습니다. 기본 설정을 사용합니다.")
        return get_default_config()

    if not os.path.exists(config_path):
        print(f"[경고] 설정 파일이 없습니다: {config_path}. 기본 설정을 사용합니다.")
        return get_default_config()

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)

        if not data:
            return get_default_config()

        return _dict_to_config(data)

    except Exception as e:
        print(f"[오류] 설정 파일 로드 실패: {e}. 기본 설정을 사용합니다.")
        return get_default_config()


def _dict_to_config(data: Dict[str, Any]) -> ExportConfig:
    """딕셔너리를 ExportConfig로 변환"""
    config = ExportConfig()

    # page
    if 'page' in data:
        p = data['page']
        config.page.enabled = p.get('enabled', True)
        config.page.orientation = p.get('orientation', True)

        if 'margins' in p:
            m = p['margins']
            config.page.margins.enabled = m.get('enabled', True)
            config.page.margins.left = m.get('left', True)
            config.page.margins.right = m.get('right', True)
            config.page.margins.top = m.get('top', True)
            config.page.margins.bottom = m.get('bottom', True)

        if 'fit_to_page' in p:
            f = p['fit_to_page']
            config.page.fit_to_page.enabled = f.get('enabled', True)
            config.page.fit_to_page.width = f.get('width', 1)
            config.page.fit_to_page.height = f.get('height', 0)

    # style
    if 'style' in data:
        s = data['style']
        config.style.enabled = s.get('enabled', True)

        if 'background' in s:
            config.style.background.enabled = s['background'].get('enabled', True)

        if 'border' in s:
            b = s['border']
            config.style.border.enabled = b.get('enabled', True)
            config.style.border.style = b.get('style', 'thin')

        if 'font' in s:
            f = s['font']
            config.style.font.enabled = f.get('enabled', True)
            config.style.font.name = f.get('name', True)
            config.style.font.size = f.get('size', True)
            config.style.font.bold = f.get('bold', True)
            config.style.font.italic = f.get('italic', True)
            config.style.font.color = f.get('color', True)
            config.style.font.underline = f.get('underline', False)
            config.style.font.strikeout = f.get('strikeout', False)

        if 'alignment' in s:
            a = s['alignment']
            config.style.alignment.enabled = a.get('enabled', True)
            config.style.alignment.horizontal = a.get('horizontal', True)
            config.style.alignment.vertical = a.get('vertical', True)
            config.style.alignment.wrap_text = a.get('wrap_text', True)

    # size
    if 'size' in data:
        sz = data['size']
        config.size.enabled = sz.get('enabled', True)

        if 'row_height' in sz:
            r = sz['row_height']
            config.size.row_height.enabled = r.get('enabled', True)
            config.size.row_height.min = r.get('min', 12)
            config.size.row_height.max = r.get('max', 409)

        if 'col_width' in sz:
            c = sz['col_width']
            config.size.col_width.enabled = c.get('enabled', True)
            config.size.col_width.min = c.get('min', 2)
            config.size.col_width.max = c.get('max', 255)

    # field
    if 'field' in data:
        fd = data['field']
        config.field.enabled = fd.get('enabled', True)
        config.field.set_hwp_field = fd.get('set_hwp_field', True)

        if 'naming' in fd:
            n = fd['naming']
            config.field.naming.use_bookmark = n.get('use_bookmark', True)

            if 'pattern' in n:
                pt = n['pattern']
                config.field.naming.pattern.enabled = pt.get('enabled', True)
                config.field.naming.pattern.left_format = pt.get('left_format', '{A}_')
                config.field.naming.pattern.top_format = pt.get('top_format', '_{B}')
                config.field.naming.pattern.both_format = pt.get('both_format', '{A}_{B}')
                config.field.naming.pattern.tolerance = pt.get('tolerance', 50)

            if 'random' in n:
                config.field.naming.random.length = n['random'].get('length', 12)

        if 'target' in fd:
            config.field.target.no_background_only = fd['target'].get('no_background_only', True)

    # protection
    if 'protection' in data:
        pr = data['protection']
        config.protection.enabled = pr.get('enabled', True)

        if 'lock_rules' in pr:
            lr = pr['lock_rules']
            config.protection.lock_rules.with_background = lr.get('with_background', True)
            config.protection.lock_rules.without_background = lr.get('without_background', False)

        if 'allow' in pr:
            al = pr['allow']
            config.protection.allow.format_cells = al.get('format_cells', True)
            config.protection.allow.format_columns = al.get('format_columns', True)
            config.protection.allow.format_rows = al.get('format_rows', True)
            config.protection.allow.insert_columns = al.get('insert_columns', False)
            config.protection.allow.insert_rows = al.get('insert_rows', False)
            config.protection.allow.delete_columns = al.get('delete_columns', False)
            config.protection.allow.delete_rows = al.get('delete_rows', False)

    # output
    if 'output' in data:
        o = data['output']

        if 'sheets' in o:
            sh = o['sheets']
            config.output.sheets.main = sh.get('main', True)
            config.output.sheets.page_info = sh.get('page_info', True)
            config.output.sheets.cell_info = sh.get('cell_info', True)
            config.output.sheets.size_info = sh.get('size_info', True)

        if 'suffix' in o:
            sf = o['suffix']
            config.output.suffix.page = sf.get('page', '_page')
            config.output.suffix.cells = sf.get('cells', '_cells')
            config.output.suffix.sizes = sf.get('sizes', '_sizes')

    return config


def save_default_config(config_path: str):
    """기본 설정을 YAML 파일로 저장"""
    if not HAS_YAML:
        print("[오류] PyYAML이 설치되지 않았습니다: pip install pyyaml")
        return

    # config_template.yaml 복사
    template_path = os.path.join(os.path.dirname(__file__), 'config_template.yaml')
    if os.path.exists(template_path):
        import shutil
        shutil.copy(template_path, config_path)
        print(f"설정 파일 생성: {config_path}")
    else:
        print(f"[오류] 템플릿 파일이 없습니다: {template_path}")


# =============================================================================
# 테스트
# =============================================================================

if __name__ == "__main__":
    # 기본 설정 출력
    config = get_default_config()
    print("=== 기본 설정 ===")
    print(f"page.enabled: {config.page.enabled}")
    print(f"style.font.enabled: {config.style.font.enabled}")
    print(f"field.enabled: {config.field.enabled}")
    print(f"protection.enabled: {config.protection.enabled}")

    # YAML 로드 테스트
    test_path = os.path.join(os.path.dirname(__file__), 'config_template.yaml')
    if os.path.exists(test_path):
        print(f"\n=== YAML 로드 테스트: {test_path} ===")
        config = load_config(test_path)
        print(f"page.margins.enabled: {config.page.margins.enabled}")
        print(f"field.naming.pattern.both_format: {config.field.naming.pattern.both_format}")
