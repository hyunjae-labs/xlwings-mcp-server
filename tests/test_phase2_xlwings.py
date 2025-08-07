"""
Phase 2 xlwings implementation tests
워크북/시트 관리 5개 함수 테스트

테스트 대상 함수:
1. create_workbook
2. create_worksheet  
3. delete_worksheet
4. rename_worksheet
5. copy_worksheet
"""

import os
import pytest
import tempfile
from pathlib import Path

# xlwings 구현 테스트

from src.excel_mcp.workbook import create_workbook, create_sheet
from src.excel_mcp.sheet import copy_sheet, delete_sheet, rename_sheet

class TestPhase2Xlwings:
    """Phase 2 xlwings 구현 테스트"""
    
    def setup_method(self):
        """각 테스트 실행 전 임시 파일 설정"""
        self.temp_dir = Path(tempfile.mkdtemp())
        self.test_file = self.temp_dir / "test_phase2.xlsx"
    
    def teardown_method(self):
        """각 테스트 실행 후 정리"""
        if self.test_file.exists():
            self.test_file.unlink()
        if self.temp_dir.exists():
            self.temp_dir.rmdir()

    def test_create_workbook_basic(self):
        """기본 워크북 생성 테스트"""
        result = create_workbook(str(self.test_file))
        
        assert "message" in result
        assert "Created workbook" in result["message"]
        assert self.test_file.exists()
    
    def test_create_workbook_custom_sheet(self):
        """커스텀 시트명으로 워크북 생성 테스트"""
        result = create_workbook(str(self.test_file), "CustomSheet")
        
        assert "message" in result
        assert "active_sheet" in result
        assert result["active_sheet"] == "CustomSheet"
        assert self.test_file.exists()

    def test_create_worksheet_basic(self):
        """기본 워크시트 생성 테스트"""
        # 먼저 워크북 생성
        create_workbook(str(self.test_file))
        
        # 새 시트 생성
        result = create_sheet(str(self.test_file), "NewSheet")
        
        assert "message" in result
        assert "NewSheet created successfully" in result["message"]
    
    def test_create_worksheet_duplicate_error(self):
        """중복 시트명 생성 시 에러 테스트"""
        # 워크북 생성
        create_workbook(str(self.test_file), "TestSheet")
        
        # 동일한 이름으로 시트 생성 시도
        with pytest.raises(Exception) as exc_info:
            create_sheet(str(self.test_file), "TestSheet")
        
        assert "already exists" in str(exc_info.value)

    def test_create_worksheet_nonexistent_file(self):
        """존재하지 않는 파일에 시트 생성 시 에러 테스트"""
        nonexistent_file = self.temp_dir / "nonexistent.xlsx"
        
        with pytest.raises(Exception):
            create_sheet(str(nonexistent_file), "NewSheet")

    def test_delete_worksheet_basic(self):
        """기본 워크시트 삭제 테스트"""
        # 워크북 생성 및 시트 2개 생성
        create_workbook(str(self.test_file))
        create_sheet(str(self.test_file), "DeleteMe")
        
        # 시트 삭제
        result = delete_sheet(str(self.test_file), "DeleteMe")
        
        assert "message" in result
        assert "deleted successfully" in result["message"]
    
    def test_delete_worksheet_only_sheet_error(self):
        """유일한 시트 삭제 시 에러 테스트"""
        # 워크북 생성 (시트 1개)
        create_workbook(str(self.test_file), "OnlySheet")
        
        # 유일한 시트 삭제 시도
        with pytest.raises(Exception) as exc_info:
            delete_sheet(str(self.test_file), "OnlySheet")
        
        assert "Cannot delete the only sheet" in str(exc_info.value)
    
    def test_delete_worksheet_nonexistent_error(self):
        """존재하지 않는 시트 삭제 시 에러 테스트"""
        create_workbook(str(self.test_file))
        
        with pytest.raises(Exception) as exc_info:
            delete_sheet(str(self.test_file), "NonexistentSheet")
        
        assert "not found" in str(exc_info.value)

    def test_rename_worksheet_basic(self):
        """기본 워크시트 이름 변경 테스트"""
        create_workbook(str(self.test_file), "OldName")
        
        result = rename_sheet(str(self.test_file), "OldName", "NewName")
        
        assert "message" in result
        assert "renamed from 'OldName' to 'NewName'" in result["message"]
    
    def test_rename_worksheet_nonexistent_error(self):
        """존재하지 않는 시트 이름 변경 시 에러 테스트"""
        create_workbook(str(self.test_file))
        
        with pytest.raises(Exception) as exc_info:
            rename_sheet(str(self.test_file), "NonexistentSheet", "NewName")
        
        assert "not found" in str(exc_info.value)
    
    def test_rename_worksheet_duplicate_error(self):
        """중복 이름으로 변경 시 에러 테스트"""
        create_workbook(str(self.test_file), "Sheet1")
        create_sheet(str(self.test_file), "Sheet2")
        
        with pytest.raises(Exception) as exc_info:
            rename_sheet(str(self.test_file), "Sheet1", "Sheet2")
        
        assert "already exists" in str(exc_info.value)

    def test_copy_worksheet_basic(self):
        """기본 워크시트 복사 테스트"""
        create_workbook(str(self.test_file), "SourceSheet")
        
        result = copy_sheet(str(self.test_file), "SourceSheet", "CopiedSheet")
        
        assert "message" in result
        assert "copied to 'CopiedSheet'" in result["message"]
    
    def test_copy_worksheet_nonexistent_source_error(self):
        """존재하지 않는 원본 시트 복사 시 에러 테스트"""
        create_workbook(str(self.test_file))
        
        with pytest.raises(Exception) as exc_info:
            copy_sheet(str(self.test_file), "NonexistentSheet", "CopiedSheet")
        
        assert "not found" in str(exc_info.value)
    
    def test_copy_worksheet_duplicate_target_error(self):
        """중복 대상 이름으로 복사 시 에러 테스트"""
        create_workbook(str(self.test_file), "SourceSheet")
        create_sheet(str(self.test_file), "TargetSheet")
        
        with pytest.raises(Exception) as exc_info:
            copy_sheet(str(self.test_file), "SourceSheet", "TargetSheet")
        
        assert "already exists" in str(exc_info.value)

    def test_workflow_multiple_operations(self):
        """여러 시트 작업의 워크플로우 테스트"""
        # 1. 워크북 생성
        create_workbook(str(self.test_file), "MainSheet")
        
        # 2. 추가 시트들 생성
        create_sheet(str(self.test_file), "DataSheet")
        create_sheet(str(self.test_file), "TempSheet")
        
        # 3. 시트 복사
        copy_sheet(str(self.test_file), "DataSheet", "BackupSheet")
        
        # 4. 시트 이름 변경
        rename_sheet(str(self.test_file), "TempSheet", "ProcessedSheet")
        
        # 5. 불필요한 시트 삭제
        delete_sheet(str(self.test_file), "ProcessedSheet")
        
        # 모든 작업이 성공하면 테스트 통과
        assert True

class TestPhase2EdgeCases:
    """Phase 2 엣지 케이스 테스트"""
    
    def setup_method(self):
        """각 테스트 실행 전 설정"""
        self.temp_dir = Path(tempfile.mkdtemp())
    
    def teardown_method(self):
        """각 테스트 실행 후 정리"""
        for file in self.temp_dir.glob("*.xlsx"):
            file.unlink()
        if self.temp_dir.exists():
            self.temp_dir.rmdir()

    def test_long_sheet_names(self):
        """긴 시트 이름 처리 테스트"""
        test_file = self.temp_dir / "long_names.xlsx"
        long_name = "VeryLongSheetNameThatExceedsNormalLength"
        
        create_workbook(str(test_file), long_name)
        
        # 시트 생성, 이름 변경, 복사 테스트
        create_sheet(str(test_file), "ShortName")
        rename_sheet(str(test_file), "ShortName", "AnotherLongName")
        copy_sheet(str(test_file), long_name, "CopyOfLongName")

    def test_special_characters_in_names(self):
        """특수 문자가 포함된 시트 이름 테스트"""
        test_file = self.temp_dir / "special_chars.xlsx"
        
        create_workbook(str(test_file), "기본시트")
        
        # 한글, 숫자, 공백이 포함된 시트명
        create_sheet(str(test_file), "데이터 시트 2024")
        copy_sheet(str(test_file), "기본시트", "복사본 시트")

    def test_directory_creation(self):
        """존재하지 않는 디렉토리에 워크북 생성 테스트"""
        nested_dir = self.temp_dir / "nested" / "directory"
        test_file = nested_dir / "test.xlsx"
        
        # 디렉토리가 자동으로 생성되어야 함
        result = create_workbook(str(test_file))
        
        assert "message" in result
        assert test_file.exists()

if __name__ == "__main__":
    # 개별 테스트 실행을 위한 코드
    print("Phase 2 xlwings 테스트를 실행합니다...")
    
    # 기본 테스트 실행
    test_instance = TestPhase2Xlwings()
    test_instance.setup_method()
    
    try:
        test_instance.test_create_workbook_basic()
        print("✅ 워크북 생성 테스트 통과")
        
        test_instance.test_create_worksheet_basic()
        print("✅ 워크시트 생성 테스트 통과")
        
        print("모든 기본 테스트가 통과했습니다!")
        
    except Exception as e:
        print(f"❌ 테스트 실패: {e}")
    
    finally:
        test_instance.teardown_method()