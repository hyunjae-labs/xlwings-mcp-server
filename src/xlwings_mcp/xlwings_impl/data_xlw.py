"""
xlwings implementation for data operations
Phase 1: read_data_from_excel, write_data_to_excel
"""

import os
import json
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path

import xlwings as xw
from .helpers import ExcelHelper

logger = logging.getLogger(__name__)

def read_data_from_excel_xlw(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> str:
    """xlwings를 사용한 데이터 읽기
    
    Args:
        filepath: Excel 파일 경로
        sheet_name: 시트명
        start_cell: 시작 셀 (기본값: A1)
        end_cell: 종료 셀 (선택사항, 자동 확장)
        preview_only: 미리보기 모드 (현재 미사용)
        
    Returns:
        JSON 형식의 문자열 - 셀 메타데이터와 함께 구조화된 데이터
    """
    app = None
    wb = None
    
    try:
        # Excel 앱 시작 (백그라운드에서)
        app = xw.App(visible=False, add_book=False)
        
        # 파일 경로 검증
        if not os.path.exists(filepath):
            return json.dumps({"error": f"File not found: {filepath}"}, indent=2)
        
        # 워크북 열기
        wb = app.books.open(filepath)
        
        # 시트 존재 확인
        if sheet_name not in [s.name for s in wb.sheets]:
            return json.dumps({"error": f"Sheet '{sheet_name}' not found"}, indent=2)
        
        ws = wb.sheets[sheet_name]
        
        # Set default start_cell if not provided
        if not start_cell:
            # Find first non-empty cell or default to A1
            used_range = ws.used_range
            if used_range:
                start_cell = used_range.address.split(":")[0].replace("$", "")
            else:
                start_cell = "A1"
        
        # 범위 결정
        if end_cell:
            # 명시적 범위 사용
            data_range = ws.range(f"{start_cell}:{end_cell}")
        else:
            # 시작 셀부터 자동 확장
            try:
                data_range = ws.range(start_cell).expand()
            except Exception:
                # 빈 시트이거나 단일 셀인 경우
                data_range = ws.range(start_cell)
        
        # 데이터 읽기
        values = data_range.value
        
        # 결과 구조 생성
        result = {
            "range": str(data_range.address),
            "sheet_name": sheet_name,
            "cells": []
        }
        
        # 셀별 데이터 변환
        if values is None:
            # 단일 빈 셀
            result["cells"].append({
                "address": data_range.address,
                "value": None,
                "row": data_range.row,
                "column": data_range.column
            })
        elif isinstance(values, list):
            # 다차원 배열
            for i, row in enumerate(values):
                if isinstance(row, list):
                    for j, val in enumerate(row):
                        cell_range = data_range.offset(i, j).resize(1, 1)
                        result["cells"].append({
                            "address": cell_range.address,
                            "value": val,
                            "row": cell_range.row,
                            "column": cell_range.column
                        })
                else:
                    # 단일 열의 경우
                    cell_range = data_range.offset(i, 0).resize(1, 1)
                    result["cells"].append({
                        "address": cell_range.address,
                        "value": row,
                        "row": cell_range.row,
                        "column": cell_range.column
                    })
        else:
            # 단일 값
            result["cells"].append({
                "address": data_range.address,
                "value": values,
                "row": data_range.row,
                "column": data_range.column
            })
        
        return json.dumps(result, indent=2, default=str, ensure_ascii=False)
        
    except Exception as e:
        logger.error(f"xlwings 데이터 읽기 실패: {e}")
        return json.dumps({"error": f"Failed to read data: {str(e)}"}, indent=2)
        
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

def write_data_to_excel_xlw(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: Optional[str] = None
) -> Dict[str, str]:
    """xlwings를 사용한 데이터 쓰기
    
    Args:
        filepath: Excel 파일 경로
        sheet_name: 시트명
        data: 쓸 데이터 (2차원 리스트)
        start_cell: 시작 셀 (기본값: A1)
        
    Returns:
        작업 결과 메시지 딕셔너리
    """
    app = None
    wb = None
    
    try:
        # 데이터 검증
        if not data:
            return {"error": "No data provided to write"}
        
        # Excel 앱 시작
        app = xw.App(visible=False, add_book=False)
        
        # 파일이 존재하는지 확인
        if os.path.exists(filepath):
            wb = app.books.open(filepath)
        else:
            # 파일이 없으면 새로 생성
            wb = app.books.add()
            # 디렉토리 생성
            Path(filepath).parent.mkdir(parents=True, exist_ok=True)
            wb.save(filepath)
        
        # 시트 확인/생성
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            wb.sheets.add(sheet_name)
        
        ws = wb.sheets[sheet_name]
        
        # Set default start_cell if not provided
        if not start_cell:
            # Find appropriate location for writing
            used_range = ws.used_range
            if used_range:
                # If sheet has data, find empty area below it
                last_row = used_range.last_cell.row
                start_cell = f"A{last_row + 2}"  # Leave one row gap
            else:
                start_cell = "A1"
        
        # 데이터 쓰기
        range_obj = ws.range(start_cell)
        range_obj.value = data
        
        # 파일 저장
        wb.save()
        
        return {"message": f"Data written to {sheet_name} starting from {start_cell}"}
        
    except Exception as e:
        logger.error(f"xlwings 데이터 쓰기 실패: {e}")
        return {"error": f"Failed to write data: {str(e)}"}
        
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

def read_data_from_excel_xlw_with_wb(
    wb,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> str:
    """xlwings 세션 기반 데이터 읽기
    
    Args:
        wb: 워크북 객체 (세션에서 전달)
        sheet_name: 시트명
        start_cell: 시작 셀 (기본값: A1)
        end_cell: 종료 셀 (선택사항, 자동 확장)
        preview_only: 미리보기 모드 (현재 미사용)
        
    Returns:
        JSON 형식의 문자열 - 셀 메타데이터와 함께 구조화된 데이터
    """
    try:
        # 시트 존재 확인
        if sheet_name not in [s.name for s in wb.sheets]:
            return json.dumps({"error": f"Sheet '{sheet_name}' not found"}, indent=2)
        
        ws = wb.sheets[sheet_name]
        
        # Set default start_cell if not provided
        if not start_cell:
            # Find first non-empty cell or default to A1
            used_range = ws.used_range
            if used_range:
                start_cell = used_range.address.split(":")[0].replace("$", "")
            else:
                start_cell = "A1"
        
        # 범위 결정
        if end_cell:
            # 명시적 범위 사용
            data_range = ws.range(f"{start_cell}:{end_cell}")
        else:
            # 시작 셀부터 자동 확장
            try:
                data_range = ws.range(start_cell).expand()
            except Exception:
                # 빈 시트이거나 단일 셀인 경우
                data_range = ws.range(start_cell)
        
        # 데이터 읽기
        values = data_range.value
        
        # 결과 구조 생성
        result = {
            "range": str(data_range.address),
            "sheet_name": sheet_name,
            "cells": []
        }
        
        # 셀별 데이터 변환
        if values is None:
            # 단일 빈 셀
            result["cells"].append({
                "address": data_range.address,
                "value": None,
                "row": data_range.row,
                "column": data_range.column
            })
        elif isinstance(values, list):
            # 다차원 배열
            for i, row in enumerate(values):
                if isinstance(row, list):
                    for j, val in enumerate(row):
                        cell_range = data_range.offset(i, j).resize(1, 1)
                        result["cells"].append({
                            "address": cell_range.address,
                            "value": val,
                            "row": cell_range.row,
                            "column": cell_range.column
                        })
                else:
                    # 단일 열의 경우
                    cell_range = data_range.offset(i, 0).resize(1, 1)
                    result["cells"].append({
                        "address": cell_range.address,
                        "value": row,
                        "row": cell_range.row,
                        "column": cell_range.column
                    })
        else:
            # 단일 값
            result["cells"].append({
                "address": data_range.address,
                "value": values,
                "row": data_range.row,
                "column": data_range.column
            })
        
        return json.dumps(result, indent=2, default=str, ensure_ascii=False)
        
    except Exception as e:
        logger.error(f"xlwings 데이터 읽기 실패: {e}")
        return json.dumps({"error": f"Failed to read data: {str(e)}"}, indent=2)

def write_data_to_excel_xlw_with_wb(
    wb,
    sheet_name: str,
    data: List[List],
    start_cell: Optional[str] = None
) -> Dict[str, str]:
    """xlwings 세션 기반 데이터 쓰기
    
    Args:
        wb: 워크북 객체 (세션에서 전달)
        sheet_name: 시트명
        data: 쓸 데이터 (2차원 리스트)
        start_cell: 시작 셀 (기본값: A1)
        
    Returns:
        작업 결과 메시지 딕셔너리
    """
    try:
        # 데이터 검증
        if not data:
            return {"error": "No data provided to write"}
        
        # 시트 확인/생성
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name not in sheet_names:
            wb.sheets.add(sheet_name)
        
        ws = wb.sheets[sheet_name]
        
        # Set default start_cell if not provided
        if not start_cell:
            # Find appropriate location for writing
            used_range = ws.used_range
            if used_range:
                # If sheet has data, find empty area below it
                last_row = used_range.last_cell.row
                start_cell = f"A{last_row + 2}"  # Leave one row gap
            else:
                start_cell = "A1"
        
        # 데이터 쓰기 (성능 최적화를 위해 calc_state_context 사용)
        with ExcelHelper.calc_state_context(wb):
            range_obj = ws.range(start_cell)
            range_obj.value = data
        
        # 파일 저장
        wb.save()
        
        return {"message": f"Data written to {sheet_name} starting from {start_cell}"}
        
    except Exception as e:
        logger.error(f"xlwings 데이터 쓰기 실패: {e}")
        return {"error": f"Failed to write data: {str(e)}"}