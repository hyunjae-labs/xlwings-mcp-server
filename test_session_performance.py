#!/usr/bin/env python3
"""
Test script to demonstrate performance improvements with session management.
"""

import time
import os
import tempfile
from pathlib import Path

# Import the MCP server components
import sys
sys.path.insert(0, 'src')

from xlwings_mcp.session import SESSION_MANAGER
import xlwings as xw


def test_old_approach(filepath, iterations=5):
    """Test the old approach: open/close every time"""
    print(f"\n🐌 Testing OLD approach (open/close every time)...")
    start_time = time.time()
    
    for i in range(iterations):
        # Old approach: create new app every time
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(filepath)
        ws = wb.sheets[0]
        
        # Do some work
        ws.range(f'A{i+1}').value = f"Old approach - iteration {i+1}"
        ws.range(f'B{i+1}').formula = f"=A{i+1}"
        
        # Close everything
        wb.save()
        wb.close()
        app.quit()
        
        print(f"  Iteration {i+1} completed")
    
    elapsed = time.time() - start_time
    print(f"⏱️  OLD approach took: {elapsed:.2f} seconds")
    print(f"   Average per operation: {elapsed/iterations:.2f} seconds")
    return elapsed


def test_new_approach(filepath, iterations=5):
    """Test the new approach: session management"""
    print(f"\n🚀 Testing NEW approach (session management)...")
    start_time = time.time()
    
    # Open workbook once
    session_id = SESSION_MANAGER.open_workbook(filepath, visible=False)
    session = SESSION_MANAGER.get_session(session_id)
    
    for i in range(iterations):
        with session.lock:
            ws = session.workbook.sheets[0]
            
            # Do some work
            ws.range(f'D{i+1}').value = f"New approach - iteration {i+1}"
            ws.range(f'E{i+1}').formula = f"=D{i+1}"
            
            # Save but don't close
            session.workbook.save()
        
        print(f"  Iteration {i+1} completed")
    
    # Close once at the end
    SESSION_MANAGER.close_workbook(session_id)
    
    elapsed = time.time() - start_time
    print(f"⏱️  NEW approach took: {elapsed:.2f} seconds")
    print(f"   Average per operation: {elapsed/iterations:.2f} seconds")
    return elapsed


def main():
    """Main test function"""
    print("=" * 60)
    print("Excel MCP Server - Session Management Performance Test")
    print("=" * 60)
    
    # Create test file
    test_dir = tempfile.mkdtemp()
    test_file = os.path.join(test_dir, "test_performance.xlsx")
    
    print(f"\n📁 Creating test file: {test_file}")
    
    # Create initial file
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()
    wb.sheets[0].range('A1').value = "Performance Test"
    wb.save(test_file)
    wb.close()
    app.quit()
    
    # Number of iterations
    iterations = 10
    print(f"\n🔄 Running {iterations} iterations for each approach...")
    
    # Run tests
    old_time = test_old_approach(test_file, iterations)
    new_time = test_new_approach(test_file, iterations)
    
    # Calculate improvement
    improvement = old_time / new_time
    time_saved = old_time - new_time
    
    print("\n" + "=" * 60)
    print("📊 RESULTS SUMMARY")
    print("=" * 60)
    print(f"Old approach (open/close every time): {old_time:.2f} seconds")
    print(f"New approach (session management):    {new_time:.2f} seconds")
    print(f"")
    print(f"🎯 Performance improvement: {improvement:.1f}x faster")
    print(f"⏱️  Time saved: {time_saved:.2f} seconds ({(time_saved/old_time)*100:.1f}%)")
    print("=" * 60)
    
    # Cleanup
    print(f"\n🧹 Cleaning up test file...")
    try:
        os.remove(test_file)
        os.rmdir(test_dir)
    except:
        pass
    
    # Make sure all sessions are closed
    SESSION_MANAGER.close_all_sessions()


if __name__ == "__main__":
    main()