"""
Excel Session Manager for xlwings MCP Server
Manages Excel application instances and workbook sessions with TTL and LRU policies.
"""

import os
import uuid
import time
import threading
import logging
from typing import Dict, Optional, Any
from pathlib import Path
from datetime import datetime

import xlwings as xw

logger = logging.getLogger(__name__)


def is_file_locked(filepath: str) -> bool:
    """
    Check if a file is locked by another process.
    
    Args:
        filepath: Path to the file to check
        
    Returns:
        True if file is locked, False otherwise
    """
    try:
        import psutil
        abs_path = os.path.abspath(filepath)
        
        # Get all processes
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # Check if process has the file open
                for item in proc.open_files():
                    if item.path == abs_path:
                        logger.info(f"FILE_LOCKED: {filepath} is locked by {proc.info['name']} (PID: {proc.info['pid']})")
                        return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    except ImportError:
        # If psutil is not available, try to open file exclusively
        try:
            with open(filepath, 'r+b') as f:
                import fcntl
                fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
            return False
        except (IOError, OSError):
            return True
        except ImportError:
            # Windows fallback
            try:
                import msvcrt
                with open(filepath, 'r+b') as f:
                    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                return False
            except:
                return True
    
    return False


class ExcelSession:
    """Represents an Excel workbook session"""
    
    def __init__(self, session_id: str, filepath: str, app: Any, workbook: Any, 
                 visible: bool = False, read_only: bool = False):
        self.id = session_id
        self.filepath = os.path.abspath(filepath)
        self.app = app
        self.workbook = workbook
        self.visible = visible
        self.read_only = read_only
        self.created_at = time.time()
        self.last_accessed = time.time()
        self.lock = threading.RLock()
        
    def touch(self):
        """Update last access time"""
        self.last_accessed = time.time()
        
    def get_info(self) -> Dict[str, Any]:
        """Get session information"""
        return {
            "session_id": self.id,
            "filepath": self.filepath,
            "visible": self.visible,
            "read_only": self.read_only,
            "created_at": datetime.fromtimestamp(self.created_at).isoformat(),
            "last_access": datetime.fromtimestamp(self.last_accessed).isoformat(),
            "sheets": [sheet.name for sheet in self.workbook.sheets] if self.workbook else []
        }


class ExcelSessionManager:
    """Singleton manager for Excel sessions"""
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, '_initialized'):
            self._initialized = True
            self._sessions: Dict[str, ExcelSession] = {}
            self._sessions_lock = threading.RLock()
            
            # Configuration from environment
            self._ttl = int(os.getenv('EXCEL_MCP_SESSION_TTL', '600'))  # 10 minutes default
            self._max_sessions = int(os.getenv('EXCEL_MCP_MAX_OPEN', '8'))  # 8 sessions max
            
            # Start cleanup thread
            self._cleanup_thread = threading.Thread(target=self._cleanup_worker, daemon=True)
            self._cleanup_thread.start()
            
            logger.info(f"ExcelSessionManager initialized: TTL={self._ttl}s, MAX={self._max_sessions}")
    
    def open_workbook(self, filepath: str, visible: bool = False, 
                     read_only: bool = False) -> str:
        """Open a workbook and create a new session"""
        
        # Generate session ID
        session_id = str(uuid.uuid4())
        
        # Check if we need to evict old sessions (LRU)
        with self._sessions_lock:
            if len(self._sessions) >= self._max_sessions:
                self._evict_lru_session()
        
        try:
            # Log session creation
            logger.debug(f"Creating session {session_id} for {filepath} (visible={visible}, read_only={read_only})")
            
            # Create Excel app instance
            app = xw.App(visible=visible, add_book=False)
            app.display_alerts = False
            app.screen_updating = not visible  # Disable screen updating for hidden instances
            
            # Open workbook
            abs_path = os.path.abspath(filepath)
            
            if os.path.exists(abs_path):
                # Check if file is locked before trying to open
                if not read_only and is_file_locked(abs_path):
                    app.quit()  # Clean up the app we just created
                    raise IOError(f"FILE_ACCESS_ERROR: '{abs_path}' is locked by another process. Use force_close_workbook_by_path() to force close it first.")
                
                wb = app.books.open(abs_path, read_only=read_only)
                logger.debug(f"Opened existing workbook: {abs_path}")
            else:
                # Create new workbook if doesn't exist
                wb = app.books.add()
                Path(abs_path).parent.mkdir(parents=True, exist_ok=True)
                wb.save(abs_path)
                logger.debug(f"Created new workbook: {abs_path}")
            
            # Create session
            session = ExcelSession(session_id, abs_path, app, wb, visible, read_only)
            
            # Store session
            with self._sessions_lock:
                self._sessions[session_id] = session
                logger.info(f"Session {session_id} created for {filepath} (total sessions: {len(self._sessions)})")
            
            return session_id
            
        except Exception as e:
            logger.error(f"Failed to create session for {filepath}: {e}")
            # Clean up on failure
            if 'app' in locals():
                try:
                    app.quit()
                except:
                    pass
            raise
    
    def get_session(self, session_id: str) -> Optional[ExcelSession]:
        """Get a session by ID"""
        with self._sessions_lock:
            session = self._sessions.get(session_id)
            if session:
                # Check if session is expired
                if hasattr(session, 'last_accessed'):
                    time_since_access = time.time() - session.last_accessed
                    if time_since_access > self._ttl:
                        logger.warning(f"SESSION_TIMEOUT: Session '{session_id}' expired (last accessed {time_since_access:.0f}s ago, TTL={self._ttl}s)")
                        # Clean up expired session
                        try:
                            if session.workbook:
                                session.workbook.close()
                            if session.app:
                                session.app.quit()
                        except:
                            pass
                        del self._sessions[session_id]
                        return None
                
                session.touch()
                logger.debug(f"Session {session_id} accessed")
            else:
                logger.warning(f"SESSION_NOT_FOUND: Session '{session_id}' not found. It may have expired or been closed.")
            return session
    
    def close_workbook(self, session_id: str, save: bool = True) -> bool:
        """Close a workbook and remove session"""
        with self._sessions_lock:
            session = self._sessions.get(session_id)
            if not session:
                logger.warning(f"Cannot close: session {session_id} not found")
                return False
            
            try:
                with session.lock:
                    logger.debug(f"Closing session {session_id}")
                    
                    # Save and close workbook
                    if session.workbook:
                        if save and not session.read_only:
                            session.workbook.save()
                        session.workbook.close()
                    
                    # Quit Excel app
                    if session.app:
                        session.app.quit()
                    
                    # Remove from sessions
                    del self._sessions[session_id]
                    logger.info(f"Session {session_id} closed (remaining sessions: {len(self._sessions)})")
                    return True
                    
            except Exception as e:
                logger.error(f"Error closing session {session_id}: {e}")
                # Force remove from sessions even on error
                if session_id in self._sessions:
                    del self._sessions[session_id]
                return False
    
    def list_sessions(self) -> list:
        """List all active sessions"""
        with self._sessions_lock:
            return [session.get_info() for session in self._sessions.values()]
    
    def close_all_sessions(self):
        """Close all sessions (for shutdown)"""
        with self._sessions_lock:
            session_ids = list(self._sessions.keys())
            
        for session_id in session_ids:
            try:
                self.close_workbook(session_id, save=False)
            except Exception as e:
                logger.error(f"Error closing session {session_id} during shutdown: {e}")
        
        logger.info("All sessions closed")
    
    def _evict_lru_session(self):
        """Evict least recently used session (must be called with lock held)"""
        if not self._sessions:
            return
        
        # Find LRU session
        lru_session = min(self._sessions.values(), key=lambda s: s.last_accessed)
        logger.info(f"Evicting LRU session {lru_session.id} (last access: {datetime.fromtimestamp(lru_session.last_accessed).isoformat()})")
        
        # Close it
        self.close_workbook(lru_session.id, save=True)
    
    def _cleanup_worker(self):
        """Background thread to clean up expired sessions"""
        while True:
            try:
                time.sleep(30)  # Check every 30 seconds
                
                current_time = time.time()
                expired_sessions = []
                
                with self._sessions_lock:
                    for session_id, session in self._sessions.items():
                        if current_time - session.last_accessed > self._ttl:
                            expired_sessions.append(session_id)
                
                # Close expired sessions
                for session_id in expired_sessions:
                    logger.info(f"Closing expired session {session_id} (TTL={self._ttl}s)")
                    try:
                        self.close_workbook(session_id, save=True)
                    except Exception as e:
                        logger.error(f"Error closing expired session {session_id}: {e}")
                        
            except Exception as e:
                logger.error(f"Error in cleanup worker: {e}")


# Global singleton instance
SESSION_MANAGER = ExcelSessionManager()