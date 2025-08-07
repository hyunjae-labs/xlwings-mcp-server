"""
xlwings implementation for workbook operations
Phase 1: get_workbook_metadata
"""

import os
import logging
from typing import Dict, Any, Optional
from pathlib import Path

import xlwings as xw
from .base import excel_context, validate_file_path, validate_sheet_exists

logger = logging.getLogger(__name__)

def get_workbook_metadata_xlw(
    filepath: str,
    include_ranges: bool = False
) -> Dict[str, Any]:
    """xlwings를 사용한 워크북 메타데이터 조회
    
    Args:
        filepath: Excel 파일 경로
        include_ranges: 각 시트의 사용 범위 포함 여부
        
    Returns:
        워크북 메타데이터 딕셔너리
    """
    try:
        # 파일 경로 검증
        file_path = validate_file_path(filepath, must_exist=True)
        
        # Excel context로 워크북 열기
        with excel_context(filepath) as wb:
            # 기본 메타데이터 수집
            metadata = {
                "filename": file_path.name,
                "full_path": str(file_path.absolute()),
                "sheets": [sheet.name for sheet in wb.sheets],
                "sheet_count": len(wb.sheets),
                "size": file_path.stat().st_size,
                "modified": file_path.stat().st_mtime
            }
            
            # 워크북 속성 추가
            try:
                # COM 객체를 통해 추가 속성 접근 (가능한 경우)
                wb_props = wb.api.BuiltinDocumentProperties
                
                # 작성자 정보
                try:
                    metadata["author"] = wb_props("Author").Value
                except Exception:
                    metadata["author"] = "Unknown"
                
                # 생성 날짜
                try:
                    metadata["created"] = wb_props("Creation Date").Value
                except Exception:
                    metadata["created"] = None
                    
                # 마지막 저장자
                try:
                    metadata["last_saved_by"] = wb_props("Last Save Time").Value
                except Exception:
                    metadata["last_saved_by"] = None
                    
            except Exception as e:
                logger.debug(f"워크북 속성 읽기 부분적 실패: {e}")
            
            # 활성 시트 정보
            if wb.sheets:
                try:
                    # xlwings에서 활성 시트는 첫 번째 시트로 가정
                    metadata["active_sheet"] = wb.sheets[0].name
                except Exception:
                    metadata["active_sheet"] = None
            
            # 시트별 범위 정보 (요청된 경우)
            if include_ranges:
                sheet_info = {}
                for sheet in wb.sheets:
                    try:
                        # 사용된 범위 확인
                        used_range = sheet.used_range
                        if used_range:
                            sheet_info[sheet.name] = {
                                "used_range": str(used_range.address),
                                "rows": used_range.rows.count,
                                "columns": used_range.columns.count,
                                "first_cell": used_range.offset(0, 0).resize(1, 1).address,
                                "last_cell": used_range.offset(
                                    used_range.rows.count - 1,
                                    used_range.columns.count - 1
                                ).resize(1, 1).address
                            }
                        else:
                            # 빈 시트
                            sheet_info[sheet.name] = {
                                "used_range": "Empty",
                                "rows": 0,
                                "columns": 0,
                                "first_cell": "A1",
                                "last_cell": "A1"
                            }
                            
                        # 시트 보호 상태 확인
                        try:
                            sheet_info[sheet.name]["protected"] = sheet.api.ProtectContents
                        except Exception:
                            sheet_info[sheet.name]["protected"] = False
                            
                    except Exception as e:
                        logger.warning(f"시트 '{sheet.name}' 정보 수집 실패: {e}")
                        sheet_info[sheet.name] = {"error": str(e)}
                
                metadata["sheet_info"] = sheet_info
            
            return metadata
        
    except Exception as e:
        logger.error(f"xlwings 워크북 메타데이터 조회 실패: {e}")
        return {"error": f"Failed to get workbook metadata: {str(e)}"}

def create_workbook_xlw(
    filepath: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """xlwings를 사용한 새 워크북 생성
    
    Args:
        filepath: 생성할 파일 경로
        sheet_name: 기본 시트명 (optional, defaults to Excel's default)
        
    Returns:
        생성 결과 딕셔너리
    """
    try:
        # Use Excel's default sheet name if not provided
        if not sheet_name:
            sheet_name = "Sheet1"  # Excel's default
        
        # Excel context로 새 워크북 생성
        with excel_context(filepath, create_if_not_exists=True, sheet_name=sheet_name) as wb:
            # 워크북이 이미 저장되었으므로 추가 작업 없음
            return {
                "message": f"Created workbook: {filepath}",
                "filepath": filepath,
                "active_sheet": sheet_name
            }
        
    except Exception as e:
        logger.error(f"xlwings 워크북 생성 실패: {e}")
        return {"error": f"Failed to create workbook: {str(e)}"}

def get_sheet_list_xlw(filepath: str) -> Dict[str, Any]:
    """xlwings를 사용한 시트 목록 조회
    
    Args:
        filepath: Excel 파일 경로
        
    Returns:
        시트 목록이 포함된 딕셔너리
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
        
        # 시트 정보 수집
        sheets_info = []
        for i, sheet in enumerate(wb.sheets):
            sheet_info = {
                "index": i,
                "name": sheet.name
            }
            
            # 시트 상태 정보 추가
            try:
                sheet_info["visible"] = sheet.api.Visible != 0  # xlSheetHidden = 0
                sheet_info["protected"] = sheet.api.ProtectContents
            except Exception:
                sheet_info["visible"] = True
                sheet_info["protected"] = False
            
            sheets_info.append(sheet_info)
        
        return {
            "sheets": sheets_info,
            "count": len(sheets_info)
        }
        
    except Exception as e:
        logger.error(f"xlwings 시트 목록 조회 실패: {e}")
        return {"error": f"Failed to get sheet list: {str(e)}"}
        
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