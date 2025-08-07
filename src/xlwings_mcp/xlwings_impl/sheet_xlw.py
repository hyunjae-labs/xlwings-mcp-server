"""
xlwings implementation for sheet operations
Phase 2: create_worksheet, delete_worksheet, rename_worksheet, copy_worksheet
"""

import os
import logging
from typing import Dict, Any
from pathlib import Path

import xlwings as xw

logger = logging.getLogger(__name__)

def create_worksheet_xlw(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """xlwings를 사용한 워크시트 생성
    
    Args:
        filepath: Excel 파일 경로
        sheet_name: 생성할 시트명
        
    Returns:
        생성 결과 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        file_path = Path(filepath)
        if not file_path.exists():
            return {"error": f"File not found: {filepath}"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 시트 이름 중복 체크
        existing_sheets = [sheet.name for sheet in wb.sheets]
        if sheet_name in existing_sheets:
            return {"error": f"Sheet '{sheet_name}' already exists"}
        
        # 새 시트 추가
        wb.sheets.add(name=sheet_name)
        
        # 파일 저장
        wb.save()
        
        return {"message": f"Sheet '{sheet_name}' created successfully"}
        
    except Exception as e:
        logger.error(f"xlwings 워크시트 생성 실패: {e}")
        return {"error": f"Failed to create worksheet: {str(e)}"}
        
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

def delete_worksheet_xlw(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """xlwings를 사용한 워크시트 삭제
    
    Args:
        filepath: Excel 파일 경로
        sheet_name: 삭제할 시트명
        
    Returns:
        삭제 결과 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        file_path = Path(filepath)
        if not file_path.exists():
            return {"error": f"File not found: {filepath}"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 시트 존재 확인
        existing_sheets = [sheet.name for sheet in wb.sheets]
        if sheet_name not in existing_sheets:
            return {"error": f"Sheet '{sheet_name}' not found"}
        
        # 시트가 1개만 있으면 삭제 불가
        if len(wb.sheets) == 1:
            return {"error": "Cannot delete the only sheet in workbook"}
        
        # 시트 삭제
        wb.sheets[sheet_name].delete()
        
        # 파일 저장
        wb.save()
        
        return {"message": f"Sheet '{sheet_name}' deleted successfully"}
        
    except Exception as e:
        logger.error(f"xlwings 워크시트 삭제 실패: {e}")
        return {"error": f"Failed to delete worksheet: {str(e)}"}
        
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

def rename_worksheet_xlw(filepath: str, old_name: str, new_name: str) -> Dict[str, Any]:
    """xlwings를 사용한 워크시트 이름 변경
    
    Args:
        filepath: Excel 파일 경로
        old_name: 기존 시트명
        new_name: 새 시트명
        
    Returns:
        이름 변경 결과 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        file_path = Path(filepath)
        if not file_path.exists():
            return {"error": f"File not found: {filepath}"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 기존 시트 확인
        existing_sheets = [sheet.name for sheet in wb.sheets]
        if old_name not in existing_sheets:
            return {"error": f"Sheet '{old_name}' not found"}
        
        # 새 이름 중복 확인
        if new_name in existing_sheets:
            return {"error": f"Sheet '{new_name}' already exists"}
        
        # 시트 이름 변경
        wb.sheets[old_name].name = new_name
        
        # 파일 저장
        wb.save()
        
        return {"message": f"Sheet renamed from '{old_name}' to '{new_name}'"}
        
    except Exception as e:
        logger.error(f"xlwings 워크시트 이름 변경 실패: {e}")
        return {"error": f"Failed to rename worksheet: {str(e)}"}
        
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

def copy_worksheet_xlw(filepath: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]:
    """xlwings를 사용한 워크시트 복사
    
    Args:
        filepath: Excel 파일 경로
        source_sheet: 원본 시트명
        target_sheet: 대상 시트명
        
    Returns:
        복사 결과 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        file_path = Path(filepath)
        if not file_path.exists():
            return {"error": f"File not found: {filepath}"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 원본 시트 확인
        existing_sheets = [sheet.name for sheet in wb.sheets]
        if source_sheet not in existing_sheets:
            return {"error": f"Source sheet '{source_sheet}' not found"}
        
        # 대상 이름 중복 확인
        if target_sheet in existing_sheets:
            return {"error": f"Target sheet '{target_sheet}' already exists"}
        
        # 시트 복사
        source = wb.sheets[source_sheet]
        
        # xlwings에서 시트 복사하기 - COM API 사용
        try:
            # COM API를 통한 시트 복사
            source.api.Copy(After=source.api)
            
            # 복사된 시트는 보통 마지막에 추가됨
            # 복사된 시트 찾기
            new_sheets = [sheet.name for sheet in wb.sheets]
            copied_sheet_name = None
            
            for sheet_name in new_sheets:
                if sheet_name not in existing_sheets:
                    copied_sheet_name = sheet_name
                    break
            
            if copied_sheet_name:
                # 복사된 시트 이름 변경
                wb.sheets[copied_sheet_name].name = target_sheet
            else:
                # 대안 방법: 수동으로 시트 생성 후 데이터 복사
                new_sheet = wb.sheets.add(name=target_sheet)
                source_range = source.used_range
                if source_range:
                    new_sheet.range("A1").value = source_range.value
                    
        except Exception as copy_error:
            logger.warning(f"COM API 복사 실패, 대안 방법 사용: {copy_error}")
            # 대안 방법: 새 시트를 만들고 데이터를 복사
            new_sheet = wb.sheets.add(name=target_sheet)
            source_range = source.used_range
            if source_range:
                new_sheet.range("A1").value = source_range.value
        
        # 파일 저장
        wb.save()
        
        return {"message": f"Sheet '{source_sheet}' copied to '{target_sheet}'"}
        
    except Exception as e:
        logger.error(f"xlwings 워크시트 복사 실패: {e}")
        return {"error": f"Failed to copy worksheet: {str(e)}"}
        
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