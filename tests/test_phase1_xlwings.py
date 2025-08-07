"""
Phase 1 xlwings 구현 테스트
5개 함수: read_data_from_excel, write_data_to_excel, apply_formula, validate_excel_range, get_workbook_metadata
"""

import os
import tempfile
import unittest
import json
from pathlib import Path

# 테스트를 위해 xlwings 모드 활성화
os.environ["USE_XLWINGS"] = "true"

try:
    from excel_mcp.xlwings_impl.data_xlw import read_data_from_excel_xlw, write_data_to_excel_xlw
    from excel_mcp.xlwings_impl.calculations_xlw import apply_formula_xlw, validate_formula_syntax_xlw
    from excel_mcp.xlwings_impl.validation_xlw import validate_excel_range_xlw
    from excel_mcp.xlwings_impl.workbook_xlw import get_workbook_metadata_xlw, create_workbook_xlw
    XLWINGS_AVAILABLE = True
except ImportError as e:
    print(f"xlwings를 사용할 수 없습니다: {e}")
    XLWINGS_AVAILABLE = False

class TestPhase1Xlwings(unittest.TestCase):
    """Phase 1 xlwings 구현 테스트 클래스"""
    
    def setUp(self):
        """테스트 준비"""
        if not XLWINGS_AVAILABLE:
            self.skipTest("xlwings가 설치되지 않았거나 Excel이 없습니다")
        
        # 임시 디렉토리 생성
        self.test_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.test_dir, "test_workbook.xlsx")
        
        # 테스트용 워크북 생성
        result = create_workbook_xlw(self.test_file, "TestSheet")
        if "error" in result:
            self.fail(f"테스트 워크북 생성 실패: {result['error']}")
    
    def tearDown(self):
        """테스트 정리"""
        # 임시 파일들 정리
        import shutil
        if os.path.exists(self.test_dir):
            try:
                shutil.rmtree(self.test_dir)
            except Exception as e:
                print(f"임시 디렉토리 정리 실패: {e}")
    
    def test_write_and_read_data(self):
        """데이터 쓰기/읽기 테스트"""
        print("\\n=== 데이터 쓰기/읽기 테스트 ===")
        
        # 테스트 데이터
        test_data = [
            ["이름", "나이", "점수"],
            ["김철수", 25, 90],
            ["이영희", 30, 85],
            ["박민수", 28, 92]
        ]
        
        # 데이터 쓰기
        write_result = write_data_to_excel_xlw(
            self.test_file, 
            "TestSheet", 
            test_data, 
            "A1"
        )
        print(f"쓰기 결과: {write_result}")
        self.assertNotIn("error", write_result, "데이터 쓰기 실패")
        
        # 데이터 읽기
        read_result = read_data_from_excel_xlw(
            self.test_file, 
            "TestSheet", 
            "A1", 
            "C4"
        )
        print(f"읽기 결과 (처음 100자): {read_result[:100]}...")
        
        # JSON 파싱 확인
        try:
            data = json.loads(read_result)
            self.assertIn("cells", data, "셀 데이터가 없습니다")
            self.assertEqual(len(data["cells"]), 12, "셀 개수가 맞지 않습니다")  # 3x4 = 12개
        except json.JSONDecodeError as e:
            self.fail(f"JSON 파싱 실패: {e}")
    
    def test_apply_formula(self):
        """수식 적용 테스트"""
        print("\\n=== 수식 적용 테스트 ===")
        
        # 먼저 숫자 데이터 준비
        number_data = [
            [10, 20, 30],
            [5, 15, 25]
        ]
        
        write_result = write_data_to_excel_xlw(
            self.test_file, 
            "TestSheet", 
            number_data, 
            "A1"
        )
        self.assertNotIn("error", write_result, "테스트 데이터 쓰기 실패")
        
        # 수식 적용 (합계)
        formula_result = apply_formula_xlw(
            self.test_file,
            "TestSheet", 
            "D1",
            "=A1+B1+C1"
        )
        print(f"수식 적용 결과: {formula_result}")
        self.assertNotIn("error", formula_result, "수식 적용 실패")
        
        # 수식 검증
        validation_result = validate_formula_syntax_xlw(
            self.test_file,
            "TestSheet",
            "D2", 
            "=SUM(A1:C1)"
        )
        print(f"수식 검증 결과: {validation_result}")
        self.assertTrue(validation_result.get("valid", False), "수식 검증 실패")
    
    def test_validate_range(self):
        """범위 검증 테스트"""
        print("\\n=== 범위 검증 테스트 ===")
        
        # 유효한 범위 검증
        valid_result = validate_excel_range_xlw(
            self.test_file,
            "TestSheet",
            "A1",
            "C5"
        )
        print(f"유효한 범위 검증: {valid_result}")
        self.assertTrue(valid_result.get("valid", False), "유효한 범위 검증 실패")
        
        # 잘못된 시트명 테스트
        invalid_sheet_result = validate_excel_range_xlw(
            self.test_file,
            "NonExistentSheet",
            "A1"
        )
        print(f"잘못된 시트 검증: {invalid_sheet_result}")
        self.assertFalse(invalid_sheet_result.get("valid", True), "잘못된 시트 검증이 성공했습니다")
    
    def test_workbook_metadata(self):
        """워크북 메타데이터 테스트"""
        print("\\n=== 워크북 메타데이터 테스트 ===")
        
        # 기본 메타데이터
        basic_metadata = get_workbook_metadata_xlw(self.test_file, include_ranges=False)
        print(f"기본 메타데이터: {basic_metadata}")
        self.assertNotIn("error", basic_metadata, "메타데이터 조회 실패")
        self.assertIn("sheets", basic_metadata, "시트 정보 없음")
        self.assertIn("TestSheet", basic_metadata["sheets"], "테스트 시트 없음")
        
        # 범위 포함 메타데이터
        detailed_metadata = get_workbook_metadata_xlw(self.test_file, include_ranges=True)
        print(f"상세 메타데이터: {detailed_metadata}")
        self.assertNotIn("error", detailed_metadata, "상세 메타데이터 조회 실패")
        self.assertIn("sheet_info", detailed_metadata, "시트 상세 정보 없음")
    
    def test_error_handling(self):
        """오류 처리 테스트"""
        print("\\n=== 오류 처리 테스트 ===")
        
        # 존재하지 않는 파일
        nonexistent_file = os.path.join(self.test_dir, "nonexistent.xlsx")
        
        read_result = read_data_from_excel_xlw(nonexistent_file, "Sheet1")
        print(f"존재하지 않는 파일 읽기: {json.loads(read_result)}")
        self.assertIn("error", json.loads(read_result), "파일 없음 오류가 처리되지 않았습니다")
        
        # 잘못된 수식
        formula_result = apply_formula_xlw(
            self.test_file,
            "TestSheet",
            "A1",
            "=INVALID_FUNCTION()"
        )
        print(f"잘못된 수식 적용: {formula_result}")
        self.assertIn("error", formula_result, "잘못된 수식 오류가 처리되지 않았습니다")

def run_tests():
    """테스트 실행 함수"""
    print("=== xlwings-mcp-server Phase 1 테스트 시작 ===")
    
    if not XLWINGS_AVAILABLE:
        print("xlwings를 사용할 수 없습니다. 테스트를 건너뜁니다.")
        return False
    
    # 테스트 실행
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestPhase1Xlwings)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    print(f"\\n=== 테스트 결과 ===")
    print(f"실행된 테스트: {result.testsRun}")
    print(f"실패: {len(result.failures)}")
    print(f"오류: {len(result.errors)}")
    
    if result.failures:
        print("\\n실패한 테스트:")
        for test, traceback in result.failures:
            print(f"- {test}: {traceback}")
    
    if result.errors:
        print("\\n오류가 발생한 테스트:")
        for test, traceback in result.errors:
            print(f"- {test}: {traceback}")
    
    success = len(result.failures) == 0 and len(result.errors) == 0
    print(f"\\n전체 결과: {'성공' if success else '실패'}")
    
    return success

if __name__ == "__main__":
    run_tests()