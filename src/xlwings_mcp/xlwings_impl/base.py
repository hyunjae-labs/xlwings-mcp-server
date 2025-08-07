"""
xlwings 공통 유틸리티 모듈
Excel 앱과 워크북 관리를 위한 context manager 및 공통 기능 제공
"""

import os
import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Optional, Generator

import xlwings as xw

logger = logging.getLogger(__name__)


@contextmanager
def excel_context(
    filepath: str, 
    visible: bool = False,
    create_if_not_exists: bool = False,
    sheet_name: str = "Sheet1"
) -> Generator[xw.Book, None, None]:
    """Excel 앱과 워크북을 관리하는 context manager
    
    Args:
        filepath: Excel 파일 경로
        visible: Excel 앱 가시성 (기본값: False)
        create_if_not_exists: 파일이 없을 경우 생성 여부 (기본값: False)
        sheet_name: 새 파일 생성 시 기본 시트명 (기본값: "Sheet1")
        
    Yields:
        xw.Book: xlwings 워크북 객체
        
    Raises:
        FileNotFoundError: 파일이 없고 create_if_not_exists=False인 경우
        Exception: Excel 앱/워크북 관련 오류
        
    Example:
        # 기존 파일 열기
        with excel_context("/path/to/file.xlsx") as wb:
            sheet = wb.sheets["Sheet1"]
            data = sheet.range("A1:C3").value
            
        # 새 파일 생성
        with excel_context("/path/to/new.xlsx", create_if_not_exists=True) as wb:
            wb.sheets[0].range("A1").value = "Hello World"
            wb.save()
    """
    app = None
    wb = None
    
    try:
        # 파일 경로 검증
        file_path = Path(filepath)
        
        if not file_path.exists() and not create_if_not_exists:
            raise FileNotFoundError(f"File not found: {filepath}")
        
        # Excel 앱 시작
        logger.debug(f"Starting Excel app (visible={visible})")
        app = xw.App(visible=visible, add_book=False)
        
        if file_path.exists():
            # 기존 파일 열기
            logger.debug(f"Opening existing workbook: {filepath}")
            wb = app.books.open(filepath)
        else:
            # 새 워크북 생성
            logger.debug(f"Creating new workbook: {filepath}")
            # 디렉토리 생성
            file_path.parent.mkdir(parents=True, exist_ok=True)
            
            wb = app.books.add()
            
            # 기본 시트명 설정
            if wb.sheets and sheet_name != wb.sheets[0].name:
                wb.sheets[0].name = sheet_name
            
            # 파일 저장 (경로 설정)
            wb.save(filepath)
        
        logger.debug(f"Successfully opened/created workbook: {filepath}")
        yield wb
        
    except Exception as e:
        logger.error(f"Excel context error: {e}")
        raise
        
    finally:
        # 리소스 정리
        if wb:
            try:
                wb.close()
                logger.debug("Workbook closed successfully")
            except Exception as e:
                logger.warning(f"Failed to close workbook: {e}")
        
        if app:
            try:
                app.quit()
                logger.debug("Excel app quit successfully")
            except Exception as e:
                logger.warning(f"Failed to quit Excel app: {e}")


@contextmanager
def excel_app_context(visible: bool = False) -> Generator[xw.App, None, None]:
    """Excel 앱만을 관리하는 context manager
    
    워크북을 직접 생성하거나 여러 워크북을 다룰 때 사용
    
    Args:
        visible: Excel 앱 가시성 (기본값: False)
        
    Yields:
        xw.App: xlwings 앱 객체
        
    Example:
        with excel_app_context() as app:
            wb1 = app.books.add()
            wb2 = app.books.add()
            # 작업...
            wb1.close()
            wb2.close()
    """
    app = None
    
    try:
        logger.debug(f"Starting Excel app context (visible={visible})")
        app = xw.App(visible=visible, add_book=False)
        yield app
        
    except Exception as e:
        logger.error(f"Excel app context error: {e}")
        raise
        
    finally:
        if app:
            try:
                app.quit()
                logger.debug("Excel app quit successfully")
            except Exception as e:
                logger.warning(f"Failed to quit Excel app: {e}")


def validate_file_path(filepath: str, must_exist: bool = True) -> Path:
    """파일 경로 유효성 검증
    
    Args:
        filepath: 검증할 파일 경로
        must_exist: 파일 존재 여부 확인 (기본값: True)
        
    Returns:
        Path: 검증된 Path 객체
        
    Raises:
        FileNotFoundError: 파일이 없고 must_exist=True인 경우
        ValueError: 잘못된 경로 형식인 경우
    """
    if not filepath or not isinstance(filepath, str):
        raise ValueError(f"Invalid filepath: {filepath}")
    
    file_path = Path(filepath)
    
    if must_exist and not file_path.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    return file_path


def validate_sheet_exists(wb: xw.Book, sheet_name: str) -> xw.Sheet:
    """시트 존재 여부 확인 및 반환
    
    Args:
        wb: xlwings 워크북 객체
        sheet_name: 확인할 시트명
        
    Returns:
        xw.Sheet: 시트 객체
        
    Raises:
        ValueError: 시트가 존재하지 않는 경우
    """
    if sheet_name not in [s.name for s in wb.sheets]:
        available_sheets = [s.name for s in wb.sheets]
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available sheets: {available_sheets}"
        )
    
    return wb.sheets[sheet_name]


class ExcelOperationError(Exception):
    """Excel 작업 관련 커스텀 예외"""
    pass


class ExcelResourceError(Exception):
    """Excel 리소스 관리 관련 커스텀 예외"""
    pass