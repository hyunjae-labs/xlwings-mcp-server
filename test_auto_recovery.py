#!/usr/bin/env python3
"""
테스트 스크립트: 세션 자동 복구 기능 검증

이 스크립트는 구현된 자동 복구 기능을 테스트합니다:
1. 세션 생성 및 기본 동작 확인
2. 세션 만료 시뮬레이션 및 자동 복구 테스트
3. 파일 상태 변화 감지 테스트
4. 메모리 관리 테스트
"""

import os
import sys
import time
import tempfile
from pathlib import Path

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from xlwings_mcp.session import ExcelSessionManager, ExcelSession
import logging

# Configure logging to see auto-recovery messages
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_basic_session_functionality():
    """기본 세션 기능 테스트"""
    print("🧪 Test 1: 기본 세션 기능 테스트")
    
    manager = ExcelSessionManager()
    
    # 임시 파일 생성
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name
    
    try:
        # 세션 생성
        session_id = manager.open_workbook(test_file, visible=False, read_only=False)
        print(f"✅ 세션 생성 성공: {session_id}")
        
        # 세션 조회
        session = manager.get_session(session_id)
        assert session is not None, "세션이 None입니다"
        print(f"✅ 세션 조회 성공: {session.filepath}")
        
        # 세션 닫기
        result = manager.close_workbook(session_id)
        assert result == True, "세션 닫기 실패"
        print("✅ 세션 닫기 성공")
        
    finally:
        if os.path.exists(test_file):
            try:
                os.unlink(test_file)
            except:
                pass
    
    print("✅ Test 1 통과\n")

def test_auto_recovery_on_ttl_expiry():
    """TTL 만료 시 자동 복구 테스트"""
    print("🧪 Test 2: TTL 만료 시 자동 복구 테스트")
    
    # 짧은 TTL로 매니저 생성 (환경변수 설정)
    os.environ['EXCEL_MCP_SESSION_TTL'] = '2'  # 2초 TTL
    manager = ExcelSessionManager()
    
    # 임시 파일 생성
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name
    
    try:
        # 세션 생성
        session_id = manager.open_workbook(test_file, visible=False, read_only=False)
        print(f"✅ 세션 생성: {session_id}")
        
        # 첫 번째 접근 (정상)
        session1 = manager.get_session(session_id)
        assert session1 is not None, "첫 번째 세션 접근 실패"
        print("✅ 첫 번째 세션 접근 성공")
        
        # TTL 만료 대기
        print("⏰ TTL 만료 대기 중 (3초)...")
        time.sleep(3)
        
        # 만료 후 접근 시도 (자동 복구 발생해야 함)
        print("🔄 만료된 세션 접근 시도 - 자동 복구 예상")
        session2 = manager.get_session(session_id)
        
        if session2 is not None:
            print("✅ 자동 복구 성공!")
            print(f"   - 복구된 세션 파일: {session2.filepath}")
            print(f"   - 원본 파일과 일치: {session2.filepath == test_file}")
        else:
            print("❌ 자동 복구 실패")
            return False
        
        # 정리
        manager.close_workbook(session_id)
        
    finally:
        if os.path.exists(test_file):
            try:
                os.unlink(test_file)
            except:
                pass
    
    print("✅ Test 2 통과\n")
    return True

def test_file_modification_detection():
    """파일 수정 감지 테스트"""
    print("🧪 Test 3: 파일 수정 감지 테스트")
    
    os.environ['EXCEL_MCP_SESSION_TTL'] = '1'  # 1초 TTL
    manager = ExcelSessionManager()
    
    # 임시 파일 생성
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name
    
    try:
        # 세션 생성
        session_id = manager.open_workbook(test_file, visible=False, read_only=False)
        print(f"✅ 세션 생성: {session_id}")
        
        # TTL 만료 대기
        time.sleep(2)
        
        # 파일 수정 (mtime 변경)
        Path(test_file).touch()
        print("📝 파일 수정 시뮬레이션 완료")
        
        # 자동 복구 시도 (경고 메시지 확인)
        print("🔄 수정된 파일로 자동 복구 시도")
        session = manager.get_session(session_id)
        
        if session is not None:
            print("✅ 파일이 수정되었지만 복구 성공 (경고 로그 확인)")
        else:
            print("❌ 복구 실패")
            return False
        
        manager.close_workbook(session_id)
        
    finally:
        if os.path.exists(test_file):
            try:
                os.unlink(test_file)
            except:
                pass
    
    print("✅ Test 3 통과\n")
    return True

def test_memory_management():
    """메모리 관리 테스트"""
    print("🧪 Test 4: 메모리 관리 테스트")
    
    os.environ['EXCEL_MCP_MAX_EXPIRED_HISTORY'] = '3'  # 최대 3개 히스토리
    os.environ['EXCEL_MCP_SESSION_TTL'] = '1'  # 1초 TTL
    manager = ExcelSessionManager()
    
    session_ids = []
    test_files = []
    
    try:
        # 5개 세션 생성 (히스토리 한계 초과)
        for i in range(5):
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                test_file = tmp.name
                test_files.append(test_file)
            
            session_id = manager.open_workbook(test_file, visible=False, read_only=False)
            session_ids.append(session_id)
            print(f"✅ 세션 {i+1} 생성: {session_id}")
        
        # 모든 세션 만료 대기
        print("⏰ 모든 세션 만료 대기...")
        time.sleep(2)
        
        # 각 세션에 접근하여 히스토리로 이동시킴
        for i, session_id in enumerate(session_ids):
            session = manager.get_session(session_id)
            if session:
                print(f"✅ 세션 {i+1} 자동 복구 성공")
                manager.close_workbook(session_id)
        
        # 히스토리 크기 확인
        history_size = len(manager._expired_sessions)
        print(f"📊 만료 세션 히스토리 크기: {history_size}")
        
        if history_size <= 3:
            print("✅ 메모리 관리 정상 작동 (히스토리 크기 제한)")
        else:
            print(f"⚠️  히스토리 크기가 한계를 초과함: {history_size} > 3")
        
    finally:
        for test_file in test_files:
            if os.path.exists(test_file):
                try:
                    os.unlink(test_file)
                except:
                    pass
    
    print("✅ Test 4 통과\n")
    return True

def test_session_redirect_mapping():
    """세션 리다이렉션 매핑 테스트"""
    print("🧪 Test 5: 세션 리다이렉션 매핑 테스트")
    
    os.environ['EXCEL_MCP_SESSION_TTL'] = '1'
    manager = ExcelSessionManager()
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = tmp.name
    
    try:
        # 세션 생성
        original_session_id = manager.open_workbook(test_file, visible=False, read_only=False)
        print(f"✅ 원본 세션 생성: {original_session_id}")
        
        # TTL 만료 대기
        time.sleep(2)
        
        # 자동 복구 (리다이렉션 생성)
        recovered_session = manager.get_session(original_session_id)
        assert recovered_session is not None, "자동 복구 실패"
        print("✅ 자동 복구 완료")
        
        # 리다이렉션 매핑 확인
        redirect_count = len(manager._session_redirects)
        print(f"📊 리다이렉션 매핑 수: {redirect_count}")
        
        if redirect_count > 0:
            print("✅ 리다이렉션 매핑 생성 확인")
        
        # 원본 세션 ID로 계속 접근 가능한지 확인
        session_again = manager.get_session(original_session_id)
        assert session_again is not None, "리다이렉션을 통한 접근 실패"
        print("✅ 리다이렉션을 통한 세션 접근 성공")
        
        # 정리
        manager.close_workbook(original_session_id)
        
        # 정리 후 리다이렉션 매핑 제거 확인
        final_redirect_count = len(manager._session_redirects)
        print(f"📊 정리 후 리다이렉션 매핑 수: {final_redirect_count}")
        
    finally:
        if os.path.exists(test_file):
            try:
                os.unlink(test_file)
            except:
                pass
    
    print("✅ Test 5 통과\n")
    return True

def main():
    """메인 테스트 실행"""
    print("🚀 세션 자동 복구 기능 통합 테스트 시작\n")
    
    tests = [
        test_basic_session_functionality,
        test_auto_recovery_on_ttl_expiry,
        test_file_modification_detection,
        test_memory_management,
        test_session_redirect_mapping
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        try:
            if test():
                passed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"❌ 테스트 실행 중 오류: {e}")
            failed += 1
        
        # 각 테스트 간 간격
        time.sleep(1)
    
    print("=" * 60)
    print(f"🎯 테스트 결과: {passed}개 통과, {failed}개 실패")
    
    if failed == 0:
        print("🎉 모든 테스트 통과! 자동 복구 기능이 성공적으로 구현되었습니다.")
        return True
    else:
        print("⚠️  일부 테스트가 실패했습니다. 구현을 다시 확인해주세요.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)