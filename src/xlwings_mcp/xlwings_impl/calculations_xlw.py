"""
xlwings implementation for formula calculations
Phase 1: apply_formula
"""

import os
import logging
from typing import Dict, Any

import xlwings as xw

logger = logging.getLogger(__name__)

def apply_formula_xlw_with_wb(
    wb,
    sheet_name: str,
    cell: str,
    formula: str
) -> Dict[str, Any]:
    """Apply formula using existing workbook object (session-based).
    
    Args:
        wb: Workbook object from session
        sheet_name: Sheet name
        cell: Target cell (e.g., A1)
        formula: Formula to apply
        
    Returns:
        Dictionary with result and calculated value
    """
    try:
        # Check sheet exists
        if sheet_name not in [s.name for s in wb.sheets]:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        ws = wb.sheets[sheet_name]
        
        # Normalize formula
        if not formula.startswith('='):
            formula = f'={formula}'
        
        # Get cell
        cell_range = ws.range(cell)
        
        # Apply formula
        try:
            cell_range.formula = formula
        except Exception as e:
            return {
                "error": f"Formula error in cell {cell}: {str(e)}",
                "formula": formula,
                "cell": cell
            }
        
        # Get calculated result
        try:
            calculated_value = cell_range.value
            display_value = cell_range.api.Text
        except Exception as e:
            logger.warning(f"Failed to read calculated value: {e}")
            calculated_value = None
            display_value = None
        
        # Save workbook
        wb.save()
        
        return {
            "message": f"Formula applied to {cell}",
            "cell": cell,
            "formula": formula,
            "calculated_value": calculated_value,
            "display_value": display_value
        }
        
    except Exception as e:
        logger.error(f"Failed to apply formula: {e}")
        return {"error": f"Failed to apply formula: {str(e)}"}

def apply_formula_xlw(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> Dict[str, Any]:
    """xlwings를 사용한 수식 적용
    
    Args:
        filepath: Excel 파일 경로
        sheet_name: 시트명
        cell: 대상 셀 (예: A1)
        formula: 적용할 수식 (= 포함 또는 미포함)
        
    Returns:
        작업 결과와 계산된 값을 포함한 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 시트 존재 확인
        if sheet_name not in [s.name for s in wb.sheets]:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        ws = wb.sheets[sheet_name]
        
        # 수식 정규화 (= 접두사 확인)
        if not formula.startswith('='):
            formula = f'={formula}'
        
        # 셀 객체 가져오기
        cell_range = ws.range(cell)
        
        # 수식 적용 (Excel이 자동으로 구문 검증)
        try:
            cell_range.formula = formula
        except Exception as e:
            # Excel에서 수식 오류 발생 시
            return {
                "error": f"Formula error in cell {cell}: {str(e)}",
                "formula": formula,
                "cell": cell
            }
        
        # 수식 계산 결과 확인
        try:
            calculated_value = cell_range.value
            display_value = cell_range.api.Text  # Excel에 표시되는 텍스트
        except Exception as e:
            logger.warning(f"계산 결과 읽기 실패: {e}")
            calculated_value = None
            display_value = None
        
        # 파일 저장
        wb.save()
        
        return {
            "message": f"Formula applied to {cell}",
            "cell": cell,
            "formula": formula,
            "calculated_value": calculated_value,
            "display_value": display_value
        }
        
    except Exception as e:
        logger.error(f"xlwings 수식 적용 실패: {e}")
        return {"error": f"Failed to apply formula: {str(e)}"}
        
    finally:
        # 리소스 정리
        if wb:
            try:
                wb.close()
            except Exception as e:
                logger.warning(f"워크북 닫기 실패: {e}")
        
        if app:
            try:
                app.quit()
            except Exception as e:
                logger.warning(f"Excel 앱 종료 실패: {e}")

def validate_formula_syntax_xlw(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str
) -> Dict[str, Any]:
    """xlwings를 사용한 수식 문법 검증
    
    Args:
        filepath: Excel 파일 경로
        sheet_name: 시트명
        cell: 대상 셀
        formula: 검증할 수식
        
    Returns:
        검증 결과 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        if not os.path.exists(filepath):
            return {"error": f"File not found: {filepath}"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 시트 존재 확인
        if sheet_name not in [s.name for s in wb.sheets]:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        ws = wb.sheets[sheet_name]
        
        # 수식 정규화
        if not formula.startswith('='):
            formula = f'={formula}'
        
        # 백업용 원본 값 저장
        cell_range = ws.range(cell)
        original_value = cell_range.value
        original_formula = cell_range.formula
        
        try:
            # 임시로 수식 적용해서 검증
            cell_range.formula = formula
            
            # 계산 결과 확인
            preview_value = cell_range.value
            
            # 원래 값으로 복원
            if original_formula and original_formula.startswith('='):
                cell_range.formula = original_formula
            else:
                cell_range.value = original_value
            
            return {
                "valid": True,
                "message": "Formula syntax is valid",
                "formula": formula,
                "preview_value": preview_value
            }
            
        except Exception as e:
            # 원래 값으로 복원 시도
            try:
                if original_formula and original_formula.startswith('='):
                    cell_range.formula = original_formula
                else:
                    cell_range.value = original_value
            except Exception:
                pass
                
            return {
                "valid": False,
                "message": f"Invalid formula syntax: {str(e)}",
                "formula": formula
            }
        
    except Exception as e:
        logger.error(f"xlwings 수식 검증 실패: {e}")
        return {"error": f"Failed to validate formula: {str(e)}"}
        
    finally:
        # 리소스 정리 (저장하지 않음 - 검증만)
        if wb:
            try:
                wb.close()
            except Exception as e:
                logger.warning(f"워크북 닫기 실패: {e}")
        
        if app:
            try:
                app.quit()
            except Exception as e:
                logger.warning(f"Excel 앱 종료 실패: {e}")