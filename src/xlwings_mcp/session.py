"""
Excel Session Manager for xlwings MCP Server
Manages Excel application instances and workbook sessions with TTL and LRU policies.
"""

import os
import uuid
import time
import threading
import logging
from typing import Dict, Optional, Any, Tuple
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
            
            # Auto-recovery support: Store expired session info for recovery
            self._expired_sessions: Dict[str, Dict[str, Any]] = {}
            self._session_redirects: Dict[str, str] = {}
            self._max_expired_history = int(os.getenv('EXCEL_MCP_MAX_EXPIRED_HISTORY', '100'))
            
            # Configuration from environment
            self._ttl = int(os.getenv('EXCEL_MCP_SESSION_TTL', '600'))  # 10 minutes default
            self._max_sessions = int(os.getenv('EXCEL_MCP_MAX_OPEN', '8'))  # 8 sessions max
            
            # Start cleanup thread
            self._cleanup_thread = threading.Thread(target=self._cleanup_worker, daemon=True)
            self._cleanup_thread.start()
            
            logger.info(f"ExcelSessionManager initialized: TTL={self._ttl}s, MAX={self._max_sessions}, Auto-Recovery=ON")

    def _extract_session_info(self, session: ExcelSession) -> Dict[str, Any]:
        """Extract essential info from session for recovery purposes"""
        try:
            file_mtime = os.path.getmtime(session.filepath) if os.path.exists(session.filepath) else None
        except (OSError, IOError):
            file_mtime = None
        
        return {
            'filepath': session.filepath,
            'visible': session.visible,
            'read_only': session.read_only,
            'created_at': session.created_at,
            'last_accessed': session.last_accessed,
            'file_mtime': file_mtime,
            'expired_at': time.time()
        }
    
    def _validate_file_state(self, session_info: Dict[str, Any]) -> Tuple[bool, Optional[str]]:
        """Validate if file is still in recoverable state"""
        filepath = session_info['filepath']
        
        # Check if file exists
        if not os.path.exists(filepath):
            return False, f"FILE_NOT_FOUND: '{filepath}' no longer exists"
        
        # Check if file is accessible
        try:
            if not session_info['read_only'] and is_file_locked(filepath):
                return False, f"FILE_LOCKED: '{filepath}' is currently locked by another process"
        except Exception as e:
            return False, f"FILE_ACCESS_ERROR: Cannot access '{filepath}': {e}"
        
        # Check if file was modified since session expired (warning, not error)
        stored_mtime = session_info.get('file_mtime')
        if stored_mtime:
            try:
                current_mtime = os.path.getmtime(filepath)
                if current_mtime > stored_mtime:
                    logger.warning(f"FILE_MODIFIED: '{filepath}' was modified since session expired. "
                                 f"Data may be inconsistent (last known: {datetime.fromtimestamp(stored_mtime)}, "
                                 f"current: {datetime.fromtimestamp(current_mtime)})")
            except (OSError, IOError):
                pass  # Ignore mtime check errors
        
        return True, None
    
    def _auto_recover_session(self, session_id: str) -> Optional[ExcelSession]:
        """Attempt to recover an expired session"""
        session_info = self._expired_sessions.get(session_id)
        if not session_info:
            return None
        
        logger.info(f"AUTO_RECOVERY: Attempting to recover session '{session_id}' for '{session_info['filepath']}'")
        
        # Validate file state before recovery
        is_valid, error_msg = self._validate_file_state(session_info)
        if not is_valid:
            logger.warning(f"AUTO_RECOVERY_FAILED: {error_msg}")
            return None
        
        try:
            # Create new session with same parameters
            new_session_id = self.open_workbook(
                filepath=session_info['filepath'],
                visible=session_info['visible'],
                read_only=session_info['read_only']
            )
            
            # Create redirect mapping from old to new session
            self._session_redirects[session_id] = new_session_id
            
            # Get the new session
            new_session = self._sessions.get(new_session_id)
            if new_session:
                logger.info(f"AUTO_RECOVERY_SUCCESS: Session '{session_id}' recovered as '{new_session_id}' "
                           f"for '{session_info['filepath']}'")
                return new_session
            
        except Exception as e:
            logger.error(f"AUTO_RECOVERY_ERROR: Failed to recover session '{session_id}': {e}")
        
        return None

    def _manage_expired_history(self):
        """Manage expired session history to prevent memory bloat"""
        if len(self._expired_sessions) > self._max_expired_history:
            # Remove oldest expired sessions (FIFO)
            expired_items = list(self._expired_sessions.items())
            expired_items.sort(key=lambda x: x[1]['expired_at'])
            
            # Remove excess items
            excess_count = len(self._expired_sessions) - self._max_expired_history
            for i in range(excess_count):
                session_id, session_info = expired_items[i]
                del self._expired_sessions[session_id]
                
                # Also remove any redirect mappings
                redirect_keys_to_remove = [k for k, v in self._session_redirects.items() if k == session_id]
                for key in redirect_keys_to_remove:
                    del self._session_redirects[key]
            
            if excess_count > 0:
                logger.debug(f"MEMORY_CLEANUP: Removed {excess_count} old expired sessions from history")
    
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
        """Get a session by ID with automatic recovery support"""
        with self._sessions_lock:
            # Check for redirect first (if session was recovered)
            actual_session_id = self._session_redirects.get(session_id, session_id)
            
            session = self._sessions.get(actual_session_id)
            if session:
                # Check if session is expired
                if hasattr(session, 'last_accessed'):
                    time_since_access = time.time() - session.last_accessed
                    if time_since_access > self._ttl:
                        logger.warning(f"SESSION_TIMEOUT: Session '{actual_session_id}' expired (last accessed {time_since_access:.0f}s ago, TTL={self._ttl}s)")
                        
                        # Store session info for potential recovery before cleanup
                        session_info = self._extract_session_info(session)
                        
                        # Clean up expired session
                        try:
                            if session.workbook:
                                session.workbook.close()
                            if session.app:
                                session.app.quit()
                        except:
                            pass
                        
                        # Move to expired sessions for potential recovery
                        self._expired_sessions[session_id] = session_info
                        self._manage_expired_history()
                        
                        # Remove from active sessions
                        del self._sessions[actual_session_id]
                        
                        # Remove redirect if it exists
                        if session_id in self._session_redirects:
                            del self._session_redirects[session_id]
                        
                        # Attempt automatic recovery
                        logger.info(f"AUTO_RECOVERY: Session '{session_id}' expired, attempting automatic recovery...")
                        recovered_session = self._auto_recover_session(session_id)
                        if recovered_session:
                            recovered_session.touch()
                            return recovered_session
                        
                        return None
                
                session.touch()
                logger.debug(f"Session {session_id} accessed")
                return session
            else:
                # Session not found in active sessions, try auto-recovery
                if session_id in self._expired_sessions:
                    logger.info(f"AUTO_RECOVERY: Session '{session_id}' not active, attempting recovery...")
                    recovered_session = self._auto_recover_session(session_id)
                    if recovered_session:
                        recovered_session.touch()
                        return recovered_session
                
                logger.warning(f"SESSION_NOT_FOUND: Session '{session_id}' not found and cannot be recovered. It may have been permanently closed.")
            
            return None
    
    def close_workbook(self, session_id: str, save: bool = True) -> bool:
        """Close a workbook and remove session"""
        with self._sessions_lock:
            # Handle redirect mapping if exists
            actual_session_id = self._session_redirects.get(session_id, session_id)
            
            session = self._sessions.get(actual_session_id)
            if not session:
                logger.warning(f"Cannot close: session {session_id} not found")
                return False
            
            try:
                with session.lock:
                    logger.debug(f"Closing session {session_id} (actual: {actual_session_id})")
                    
                    # Save and close workbook
                    if session.workbook:
                        if save and not session.read_only:
                            session.workbook.save()
                        session.workbook.close()
                    
                    # Quit Excel app
                    if session.app:
                        session.app.quit()
                    
                    # Remove from sessions
                    del self._sessions[actual_session_id]
                    
                    # Clean up auto-recovery related data
                    if session_id in self._expired_sessions:
                        del self._expired_sessions[session_id]
                    
                    # Clean up redirect mappings
                    redirect_keys_to_remove = []
                    for k, v in self._session_redirects.items():
                        if k == session_id or v == actual_session_id:
                            redirect_keys_to_remove.append(k)
                    
                    for key in redirect_keys_to_remove:
                        del self._session_redirects[key]
                    
                    logger.info(f"Session {session_id} closed permanently (remaining sessions: {len(self._sessions)})")
                    return True
                    
            except Exception as e:
                logger.error(f"Error closing session {session_id}: {e}")
                # Force remove from sessions even on error
                if actual_session_id in self._sessions:
                    del self._sessions[actual_session_id]
                    
                # Clean up recovery data on error too
                if session_id in self._expired_sessions:
                    del self._expired_sessions[session_id]
                if session_id in self._session_redirects:
                    del self._session_redirects[session_id]
                    
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
        """Background thread to clean up expired sessions while preserving recovery info"""
        while True:
            try:
                time.sleep(30)  # Check every 30 seconds
                
                current_time = time.time()
                expired_sessions = []
                
                with self._sessions_lock:
                    for session_id, session in self._sessions.items():
                        if current_time - session.last_accessed > self._ttl:
                            expired_sessions.append((session_id, session))
                
                # Process expired sessions - move to history instead of permanent deletion
                for session_id, session in expired_sessions:
                    logger.info(f"TTL_CLEANUP: Moving expired session '{session_id}' to recovery history (TTL={self._ttl}s)")
                    try:
                        with self._sessions_lock:
                            # Extract session info for recovery before cleanup
                            session_info = self._extract_session_info(session)
                            
                            # Clean up Excel resources
                            try:
                                if session.workbook:
                                    session.workbook.close()
                                if session.app:
                                    session.app.quit()
                            except Exception as cleanup_error:
                                logger.warning(f"Error cleaning up Excel resources for session {session_id}: {cleanup_error}")
                            
                            # Move to expired sessions for potential recovery
                            self._expired_sessions[session_id] = session_info
                            self._manage_expired_history()
                            
                            # Remove from active sessions
                            if session_id in self._sessions:
                                del self._sessions[session_id]
                            
                            logger.debug(f"TTL_CLEANUP: Session '{session_id}' moved to recovery history (active: {len(self._sessions)}, history: {len(self._expired_sessions)})")
                            
                    except Exception as e:
                        logger.error(f"Error processing expired session {session_id}: {e}")
                        # Force cleanup if regular cleanup fails
                        try:
                            with self._sessions_lock:
                                if session_id in self._sessions:
                                    del self._sessions[session_id]
                        except:
                            pass
                        
            except Exception as e:
                logger.error(f"Error in cleanup worker: {e}")


# Global singleton instance
SESSION_MANAGER = ExcelSessionManager()