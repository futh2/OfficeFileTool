"""Temporary Office startup trial. Python 3.9+, Windows; pip install psutil pywin32."""
import argparse
import contextlib
import ctypes
import hashlib
import json
import os
import queue
import shutil
import struct
import subprocess
import sys
import threading
import tempfile
import time
import winreg
from ctypes import wintypes as w
from pathlib import Path
import psutil

EXTENSIONS = {
    'Word': {'.doc', '.docx', '.docm', '.rtf', '.odt', '.txt'},
    'PowerPoint': {'.ppt', '.pptx', '.pptm', '.pps', '.ppsx', '.ppsm', '.odp'},
    'Excel': {'.xls', '.xlsx', '.xlsm', '.xlsb', '.csv', '.ods'},
    'Access': {'.accdb', '.accde', '.mdb', '.mde'},
    'Publisher': {'.pub'},
    'OneNote': {'.one', '.onepkg', '.onetoc2'},
    'Outlook': {'.msg', '.oft', '.ics'},
}
EXECUTABLES = {'Word': 'WINWORD.EXE', 'PowerPoint': 'POWERPNT.EXE', 'Excel': 'EXCEL.EXE',
               'Access': 'MSACCESS.EXE', 'Publisher': 'MSPUB.EXE', 'OneNote': 'ONENOTE.EXE', 'Outlook': 'OUTLOOK.EXE'}
NORMAL_APPS = {'OneNote', 'Outlook'}
WINDOW_CLASSES = {'OpusApp', 'PPTFrameClass', 'XLMAIN', 'OMain', 'MSWinPub'}

def self_command(*arguments):
    if getattr(sys, 'frozen', False):
        return [sys.executable, *arguments]
    return [sys.executable, str(Path(__file__).resolve()), *arguments]

HASHES = {
    'wwlib.dll': '25d64db84dde4e7bb9de057477f4fa57b9143f40df46edb40a8e60508e31744f',
    'mso.dll': 'b4ee4a7f38a651409fcb69e6541b1def0d6ef9c368f5dac26bd8e72fbb52fb3d',
    'excel.exe': '91a38a93114c7e3459952c8c3bcd46671c0d540220bff7b1194092b4fd82dd2c',
    'msaccess.exe': 'dcb6c29c3f10f4fe4c5dea9148658b807ab17a1a86f6abe9e0d5b8693f3bb27f',
    'mspub.exe': '74dd5b7d9e578c987f348a4a450819d309826726dfb6c40dd1df1a29e4e5bd29',
    'onenote.exe': '303841a7bb513491d8f692fce1f5d337de5755b3d385131524aa11fec21b333b',
    'outlook.exe': 'a72b466a6ddb82a765508dacad8a8820e4a410268dab09054399371ba69f0507',
    'winword.exe': 'c99dc4696c702e8e8eb11f8c8f9b6930b6d9c0565565eabc463eff5fbe122ff1',
    'powerpnt.exe': '07cc3d67d17f14315d92e81f7289c03b3e2bb59a5831ddfa6e1700c540d0f012',
    'mso20win32client.dll': '34ec026f73f63ef068d3e6179cc611a9e04ccc5c38af0df9bdb70af43b97a6a8',
    'mso30win32client.dll': 'a3ddca3cc6de1692700339e32c8b0e59b77566953cc5010e57806e3fd4b34acd',
    'mso98win32client.dll': '69ca9889afc4cf86779143772d90c5811e0d9572c9f0ef4daba6b901a5e41489',
}
APP_VERSION = '1.3.4'
SUPPORTED_OFFICE_VERSION = '16.0.19628.20132'
POLICY_KEY = r'SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'
C2R_KEY = r'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
RESULT_FILE = None
STATE_DIR = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'OfficeStartupTrial'

def choose_app(filename):
    suffix = Path(filename).suffix.lower()
    for app, extensions in EXTENSIONS.items():
        if suffix in extensions:
            return app
    raise ValueError('Unsupported file type / نوع الملف غير مدعوم')

def validate_file(filename, app):
    path = Path(filename).expanduser().resolve()
    if not path.is_file():
        raise ValueError('File does not exist / الملف غير موجود')
    if choose_app(path) != app:
        raise ValueError('File type does not match application / اختر البرنامج المناسب للملف')
    return path

def state_path(app):
    return STATE_DIR / (app.lower() + '.json')

def save_state(state):
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    target = state_path(state['app'])
    temp = target.with_suffix('.tmp')
    temp.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding='utf-8')
    temp.replace(target)

def process_alive(state):
    try:
        p = psutil.Process(state['pid'])
        return abs(p.create_time() - state['create_time']) < .01 and p.name().lower() == EXECUTABLES[state['app']].lower()
    except psutil.NoSuchProcess:
        return False

def verify_binary(path):
    expected = HASHES.get(Path(path).name.lower())
    if not expected or hashlib.sha256(Path(path).read_bytes()).hexdigest() != expected:
        raise RuntimeError('Unsupported Office build; no changes applied: ' + str(path))

def find_executable(app):
    name = EXECUTABLES[app]
    candidates = []
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
            try:
                with winreg.OpenKey(hive, 'Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\' + name, 0, winreg.KEY_READ | view) as key:
                    candidates.append(Path(winreg.QueryValueEx(key, '')[0]))
            except OSError:
                pass
    for env in ('ProgramFiles(x86)', 'ProgramFiles'):
        if os.environ.get(env):
            candidates.append(Path(os.environ[env]) / 'Microsoft Office/root/Office16' / name)
    for path in candidates:
        if path.is_file():
            verify_binary(path)
            return path
    raise RuntimeError('Office application not found: ' + app)

def restore_regions(regions, read, write):
    for region in regions:
        original = bytes.fromhex(region['original'])
        current = read(region['address'], len(original))
        if current not in (original, bytes.fromhex(region['replacement'])):
            raise RuntimeError('Memory changed outside this tool; restoration stopped')
    for region in reversed(regions):
        write(region['address'], bytes.fromhex(region['original']))

class FloatSave(ctypes.Structure):
    _fields_=[(n,w.DWORD) for n in ['ControlWord','StatusWord','TagWord','ErrorOffset','ErrorSelector','DataOffset','DataSelector']]+[('RegisterArea',ctypes.c_ubyte*80),('Cr0NpxState',w.DWORD)]
class Context(ctypes.Structure):
    _fields_=[('ContextFlags',w.DWORD)]+[(n,w.DWORD) for n in ['Dr0','Dr1','Dr2','Dr3','Dr6','Dr7']]+[('FloatSave',FloatSave)]+[(n,w.DWORD) for n in ['SegGs','SegFs','SegEs','SegDs','Edi','Esi','Ebx','Edx','Ecx','Eax','Ebp','Eip','SegCs','EFlags','Esp','SegSs']]+[('ExtendedRegisters',ctypes.c_ubyte*512)]
class ExceptionRecord(ctypes.Structure):
    _fields_=[('code',w.DWORD),('flags',w.DWORD),('record',ctypes.c_void_p),('address',ctypes.c_void_p),('count',w.DWORD),('information',ctypes.c_size_t*15)]
class ExceptionInfo(ctypes.Structure):
    _fields_=[('record',ExceptionRecord),('first_chance',w.DWORD)]
class CreateProcess(ctypes.Structure):
    _fields_=[('file',w.HANDLE),('process',w.HANDLE),('thread',w.HANDLE),('base',ctypes.c_void_p),('debug_offset',w.DWORD),('debug_size',w.DWORD),('tls',ctypes.c_void_p),('start',ctypes.c_void_p),('image_name',ctypes.c_void_p),('unicode',w.WORD)]
class CreateThread(ctypes.Structure):
    _fields_=[('thread',w.HANDLE),('tls',ctypes.c_void_p),('start',ctypes.c_void_p)]
class LoadDll(ctypes.Structure):
    _fields_=[('file',w.HANDLE),('base',ctypes.c_void_p),('offset',w.DWORD),('size',w.DWORD),('name',ctypes.c_void_p),('unicode',w.WORD)]
class Payload(ctypes.Union):
    _fields_=[('exception',ExceptionInfo),('process',CreateProcess),('thread',CreateThread),('dll',LoadDll),('raw',ctypes.c_ubyte*160)]
class Event(ctypes.Structure):
    _fields_=[('code',w.DWORD),('pid',w.DWORD),('tid',w.DWORD),('data',Payload)]
k=ctypes.WinDLL('kernel32',use_last_error=True)
for name,params,result in [
    ('OpenThread',[w.DWORD,w.BOOL,w.DWORD],w.HANDLE),('OpenProcess',[w.DWORD,w.BOOL,w.DWORD],w.HANDLE),
    ('CloseHandle',[w.HANDLE],w.BOOL),('DebugActiveProcess',[w.DWORD],w.BOOL),('DebugActiveProcessStop',[w.DWORD],w.BOOL),
    ('DebugSetProcessKillOnExit',[w.BOOL],w.BOOL),('CheckRemoteDebuggerPresent',[w.HANDLE,ctypes.POINTER(w.BOOL)],w.BOOL),
    ('WaitForDebugEvent',[ctypes.POINTER(Event),w.DWORD],w.BOOL),('ContinueDebugEvent',[w.DWORD,w.DWORD,w.DWORD],w.BOOL),
    ('Wow64GetThreadContext',[w.HANDLE,ctypes.POINTER(Context)],w.BOOL),('Wow64SetThreadContext',[w.HANDLE,ctypes.POINTER(Context)],w.BOOL),
    ('SuspendThread',[w.HANDLE],w.DWORD),('ResumeThread',[w.HANDLE],w.DWORD),
    ('ReadProcessMemory',[w.HANDLE,ctypes.c_void_p,ctypes.c_void_p,ctypes.c_size_t,ctypes.c_void_p],w.BOOL),
]:
    fn=getattr(k,name);fn.argtypes=params;fn.restype=result
def check(result):
    if not result:raise ctypes.WinError(ctypes.get_last_error())
    return result

for name, params, result in [
    ('VirtualAllocEx', [w.HANDLE, ctypes.c_void_p, ctypes.c_size_t, w.DWORD, w.DWORD], ctypes.c_void_p),
    ('VirtualProtectEx', [w.HANDLE, ctypes.c_void_p, ctypes.c_size_t, w.DWORD, ctypes.POINTER(w.DWORD)], w.BOOL),
    ('WriteProcessMemory', [w.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.c_void_p], w.BOOL),
    ('FlushInstructionCache', [w.HANDLE, ctypes.c_void_p, ctypes.c_size_t], w.BOOL),
    ('GetFinalPathNameByHandleW', [w.HANDLE, w.LPWSTR, w.DWORD, w.DWORD], w.DWORD),
]:
    fn = getattr(k, name)
    fn.argtypes, fn.restype = params, result
nt = ctypes.WinDLL('ntdll')
nt.NtSuspendProcess.argtypes = nt.NtResumeProcess.argtypes = [w.HANDLE]
nt.NtSuspendProcess.restype = nt.NtResumeProcess.restype = w.LONG

def is_setup_dialog(window_class, title):
    return window_class == 'NUIDialog' and title.strip().casefold() in {
        'تسجيل الدخول لإعداد office', 'sign in to set up office',
    }

def watch_setup_dialogs(app, pid, created):
    import pythoncom
    import win32gui
    import win32process
    state = dict(app=app, pid=pid, create_time=created)
    report = STATE_DIR / (app.lower() + '_dialogs_' + str(pid) + '.json')
    stop = report.with_suffix('.stop')
    user = ctypes.WinDLL('user32', use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(None, w.HANDLE, w.DWORD, w.HWND, w.LONG, w.LONG, w.DWORD, w.DWORD)
    user.SetWinEventHook.argtypes = [w.DWORD, w.DWORD, w.HMODULE, callback_type, w.DWORD, w.DWORD, w.DWORD]
    user.SetWinEventHook.restype = w.HANDLE
    user.UnhookWinEvent.argtypes = [w.HANDLE]
    user.UnhookWinEvent.restype = w.BOOL
    dismissed, errors, pending, enable_after = [], [], set(), {}
    def close_setup(hwnd):
        try:
            if not hwnd or not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return
            if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                return
            title = win32gui.GetWindowText(hwnd)
            if not is_setup_dialog(win32gui.GetClassName(hwnd), title):
                return
            first_seen = hwnd not in pending
            pending.add(hwnd)
            win32gui.ShowWindow(hwnd, 0)
            enable_after[hwnd] = time.monotonic() + .75
            if first_seen:
                dismissed.append(dict(hwnd=hwnd, title=title, time=time.time(), action='hidden'))
        except Exception as exc:
            errors.append(str(exc))
    @callback_type
    def on_event(hook, event, hwnd, object_id, child_id, thread, milliseconds):
        if event in (0x8000, 0x8002) and object_id == 0 and child_id == 0:
            close_setup(hwnd)
    hook = check(user.SetWinEventHook(0x8000, 0x8002, None, on_event, pid, 0, 0))
    def persist(running):
        report.write_text(json.dumps(dict(pid=pid, ready=True, running=running, dismissed=dismissed, errors=errors), ensure_ascii=False, indent=2), encoding='utf-8')
    try:
        persist(True)
        last_count = (0,0)
        next_scan = 0
        while not stop.exists() and process_alive(state):
            pythoncom.PumpWaitingMessages()
            for hwnd, deadline in list(enable_after.items()):
                if time.monotonic() >= deadline:
                    if win32gui.IsWindow(hwnd):
                        document_windows = []
                        def find_documents(window, _):
                            if win32process.GetWindowThreadProcessId(window)[1] == pid and win32gui.GetClassName(window) in WINDOW_CLASSES:
                                document_windows.append(window)
                        win32gui.EnumWindows(find_documents, None)
                        if not document_windows:
                            enable_after[hwnd] = time.monotonic() + .2
                            continue
                        for window in document_windows:
                            win32gui.EnableWindow(window, True)
                    del enable_after[hwnd]
            if time.monotonic() >= next_scan:
                win32gui.EnumWindows(lambda hwnd, _: close_setup(hwnd), None)
                pending.intersection_update(hwnd for hwnd in list(pending) if win32gui.IsWindow(hwnd))
                next_scan = time.monotonic() + .2
            if last_count != (len(dismissed), len(errors)):
                persist(True)
                last_count = (len(dismissed), len(errors))
            time.sleep(.025)
    finally:
        user.UnhookWinEvent(hook)
        persist(False)

def launch_dialog_watcher(state):
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    report = STATE_DIR / (state['app'].lower() + '_dialogs_' + str(state['pid']) + '.json')
    report.unlink(missing_ok=True)
    report.with_suffix('.stop').unlink(missing_ok=True)
    command = self_command('watch-dialogs', '--app', state['app'], '--pid', str(state['pid']), '--created', str(state['create_time']))
    watcher = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=0x08000000, env=dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT='1'))
    state['dialog_watcher_pid'] = watcher.pid
    state['dialog_report'] = str(report)
    deadline = time.monotonic() + (25 if getattr(sys, 'frozen', False) else 6)
    while time.monotonic() < deadline:
        if report.exists():
            try:
                if json.loads(report.read_text(encoding='utf-8')).get('ready'):
                    return
            except (ValueError, OSError):
                pass
        if watcher.poll() is not None:
            raise RuntimeError('Office setup dialog watcher failed to start')
        time.sleep(.04)
    report.with_suffix('.stop').touch()
    raise RuntimeError('Office setup dialog watcher did not become ready')

class Trial:
    def __init__(self, state, handle):
        self.state, self.handle = state, handle

    def read(self, address, size):
        buffer = ctypes.create_string_buffer(size)
        check(k.ReadProcessMemory(self.handle, address, buffer, size, None))
        return buffer.raw

    def write(self, address, data):
        old = w.DWORD()
        check(k.VirtualProtectEx(self.handle, address, len(data), 0x40, ctypes.byref(old)))
        try:
            buffer = ctypes.create_string_buffer(data)
            check(k.WriteProcessMemory(self.handle, address, buffer, len(data), None))
            check(k.FlushInstructionCache(self.handle, address, len(data)))
        finally:
            scratch = w.DWORD()
            check(k.VirtualProtectEx(self.handle, address, len(data), old.value, ctypes.byref(scratch)))
        if self.read(address, len(data)) != data:
            raise RuntimeError('Memory verification failed')

    def patch(self, label, address, original, replacement):
        if self.read(address, len(original)) != original:
            raise RuntimeError('Unexpected instructions: ' + label)
        self.state['patches'].append(dict(name=label, address=address, original=original.hex(), replacement=replacement.hex()))
        save_state(self.state)
        self.write(address, replacement)

    def hook(self, label, address, enable_translation=False):
        original = bytes.fromhex('558bec6aff')
        if self.read(address, 5) != original:
            raise RuntimeError('Unexpected feature getter: ' + label)
        code = check(k.VirtualAllocEx(self.handle, None, 4096, 0x3000, 4))
        self.state['allocations'].append(code)
        save_state(self.state)
        if enable_translation:
            stub = bytes.fromhex('817c2404') + struct.pack('<I', 0x3cb19a9) + bytes.fromhex('740a817c2404') + struct.pack('<I', 0x164ec3) + bytes.fromhex('7508b801000000c20400') + original
        else:
            stub = bytes.fromhex('817c2404') + struct.pack('<I', 0x3cb19a9) + bytes.fromhex('7508b801000000c20400') + original
        stub += b'\xe9' + struct.pack('<I', ((address + 5) - (code + len(stub) + 5)) & 0xffffffff)
        buffer = ctypes.create_string_buffer(stub)
        check(k.WriteProcessMemory(self.handle, code, buffer, len(stub), None))
        old = w.DWORD()
        check(k.VirtualProtectEx(self.handle, code, 4096, 0x20, ctypes.byref(old)))
        check(k.FlushInstructionCache(self.handle, code, len(stub)))
        self.patch(label, address, original, b'\xe9' + struct.pack('<I', (code - (address + 5)) & 0xffffffff))

    def rollback(self):
        if nt.NtSuspendProcess(self.handle) < 0:
            raise RuntimeError('Cannot suspend process for restoration')
        try:
            restore_regions(self.state['patches'], self.read, self.write)
            # Retain trampoline pages until normal process exit: a thread may still be inside one.
            self.state['status'] = 'memory_restored_restart_required'
            save_state(self.state)
        finally:
            if nt.NtResumeProcess(self.handle) < 0:
                raise RuntimeError('Cannot resume process')

@contextlib.contextmanager
def operation_lock(app):
    create = k.CreateMutexW
    create.argtypes, create.restype = [ctypes.c_void_p, w.BOOL, w.LPCWSTR], w.HANDLE
    ctypes.set_last_error(0)
    handle = check(create(None, False, 'Local\\OfficeStartupTrial_' + app))
    existed = ctypes.get_last_error() == 183
    try:
        if existed:
            raise RuntimeError('Another operation is running for ' + app)
        yield
    finally:
        k.CloseHandle(handle)

def start_trial(app, filename=''):
    with operation_lock(app):
        if app in NORMAL_APPS:
            path = validate_file(filename, app) if filename else None
            command = [str(find_executable(app))]
            if path and app == 'Outlook':
                command += [{'.msg':'/f', '.oft':'/t', '.ics':'/ical'}[path.suffix.lower()]]
            if path:
                command.append(str(path))
            subprocess.Popen(command)
            return {'app':app, 'status':'normal_open'}
        if filename and state_path(app).exists():
            state = json.loads(state_path(app).read_text(encoding='utf-8'))
            if state.get('status') == 'applied' and process_alive(state):
                return open_in_trial(app, filename, state)
        return _start_trial(app, filename)

def open_in_trial(app, filename, state):
    import pythoncom
    import win32com.client
    path = validate_file(filename, app)
    pythoncom.CoInitialize()
    try:
        application = win32com.client.GetActiveObject(app + '.Application')
        if application_pid(app, application) != state['pid']:
            raise RuntimeError('Cannot identify trial process')
        if app == 'Word':
            application.Documents.Open(str(path)).Activate()
        elif app == 'PowerPoint':
            application.Presentations.Open(str(path))
        elif app == 'Excel':
            application.Workbooks.Open(str(path)).Activate()
        elif app == 'Publisher':
            application.Open(str(path))
        else:
            raise RuntimeError('احفظ وأغلق Access قبل فتح قاعدة بيانات أخرى من الأداة.')
        return dict(state, status='opened_in_existing_session')
    finally:
        pythoncom.CoUninitialize()

def _start_trial(app, filename=''):
    document = validate_file(filename, app) if filename else None
    if any((p.info['name'] or '').lower() == EXECUTABLES[app].lower() for p in psutil.process_iter(['name'])):
        raise RuntimeError('احفظ ملفاتك وأغلق ' + app + ' أولاً، ثم أعد التشغيل من الأداة.')
    exe = find_executable(app)
    shared = exe.parent.parent / 'vfs/ProgramFilesCommonX86/Microsoft Shared/Office16'
    required_modules = ['Mso20win32client.dll', 'Mso30win32client.dll', 'Mso98win32client.dll']
    if app == 'Word':
        required_modules.append('MSO.DLL')
    for name in required_modules:
        verify_binary(shared / name)
    command = [str(exe)] + (['/n'] if app == 'Word' else ['/x'] if app == 'Excel' else []) + ([str(document)] if document else [])
    child = subprocess.Popen(command, creationflags=2, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    state = dict(app=app, executable=str(exe), document=str(document or ''), pid=child.pid,
                 create_time=psutil.Process(child.pid).create_time(), status='starting', patches=[], allocations=[])
    process = None
    attached, exited, failure = True, False, None
    handles = set()
    try:
        process = check(k.OpenProcess(0xc38, False, child.pid))
        trial = Trial(state, process)
        check(k.DebugSetProcessKillOnExit(False))
        launch_dialog_watcher(state)
        save_state(state)
        deadline = time.monotonic() + 45
        while time.monotonic() < deadline:
            event = Event()
            if not k.WaitForDebugEvent(ctypes.byref(event), 100):
                if ctypes.get_last_error() == 121:
                    continue
                raise ctypes.WinError(ctypes.get_last_error())
            disposition = 0x10002
            try:
                if event.code == 3:
                    handles.update([event.data.process.file, event.data.process.process, event.data.process.thread])
                elif event.code == 2:
                    handles.add(event.data.thread.thread)
                elif event.code == 6:
                    info = event.data.dll
                    handles.add(info.file)
                    buffer = ctypes.create_unicode_buffer(4096)
                    if info.file and k.GetFinalPathNameByHandleW(info.file, buffer, 4096, 0):
                        path = Path(buffer.value)
                        name, base = path.name.lower(), int(info.base or 0)
                        if name in ('mso20win32client.dll', 'mso30win32client.dll', 'mso98win32client.dll'):
                            verify_binary(path)
                        if name == 'mso20win32client.dll':
                            trial.hook('fluent2025_fallback_provider', base + 0x488fc3, enable_translation=(app == 'Word'))
                        elif name == 'wwlib.dll' and app == 'Word':
                            verify_binary(path)
                            trial.patch('inline_translation_availability', base + 0x119c3d1, bytes.fromhex('558bec51'), bytes.fromhex('b001c390'))
                        elif name == 'mso.dll' and app == 'Word':
                            verify_binary(path)
                            # Native Copilot visibility enum: 0 exposes the existing control.
                            # This does not enable command execution or grant cloud access.
                            trial.patch('copilot_ribbon_visibility_ui_only', base + 0xf4a81, bytes.fromhex('558bec83e4f883ec0c'), bytes.fromhex('31c0c20c0090909090'))
                        elif name == 'mso30win32client.dll':
                            trial.patch('session_permission', base + 0xa95b3, bytes.fromhex('e8ca86feff'), bytes.fromhex('6a0158c390'))
                        elif name == 'mso98win32client.dll':
                            trial.hook('fluent2025_registered_provider', base + 0xa5471, enable_translation=(app == 'Word'))
                            trial.patch('skip_disable_app', base + 0x3b5ecd, b'\x55', b'\xc3')
                elif event.code == 1:
                    if event.data.exception.record.code not in (0x80000003, 0x4000001f):
                        disposition = 0x80010001
                elif event.code == 5:
                    exited = True
                    raise RuntimeError('Office exited during startup')
            finally:
                check(k.ContinueDebugEvent(event.pid, event.tid, disposition))
            if len(state['patches']) == (6 if app == 'Word' else 4):
                state['status'] = 'applied'
                break
        if state['status'] != 'applied':
            raise RuntimeError('Startup timeout: required Office libraries not loaded')
    except BaseException as exc:
        failure = exc
        state['error'] = str(exc)
        if state.get('dialog_report'):
            Path(state['dialog_report']).with_suffix('.stop').touch()
        if process and not exited:
            try:
                trial.rollback()
            except Exception as rollback_error:
                state['rollback_error'] = str(rollback_error)
    finally:
        if attached and not exited:
            attached = not bool(k.DebugActiveProcessStop(child.pid))
        state['debugger_detached'] = not attached or exited
        if attached and not exited:
            state['detach_error'] = str(ctypes.WinError(ctypes.get_last_error()))
        for handle in handles:
            if handle:
                k.CloseHandle(handle)
        if process:
            k.CloseHandle(process)
        save_state(state)
    if failure:
        raise failure
    if not state['debugger_detached']:
        raise RuntimeError('Debugger detach failed; keep this tool open')
    return state

def application_pid(app, application):
    processes = [p.pid for p in psutil.process_iter(['name']) if (p.info['name'] or '').lower() == EXECUTABLES[app].lower()]
    if len(processes) != 1:
        raise RuntimeError('Cannot identify a unique process for ' + app)
    return processes[0]

def restore_trial(app):
    with operation_lock(app):
        return _restore_trial(app)

def _restore_trial(app):
    import pythoncom
    import win32com.client
    import win32gui
    import win32process
    if app in NORMAL_APPS:
        return {'app':app, 'status':'normal_open_no_trial'}
    path = state_path(app)
    if not path.exists():
        return {'status': 'no_trial', 'app': app}
    state = json.loads(path.read_text(encoding='utf-8'))
    if not process_alive(state):
        state['status'] = 'restored_by_process_exit'
        save_state(state)
        return state
    pythoncom.CoInitialize()
    try:
        application = win32com.client.GetActiveObject(app + '.Application')
        if application_pid(app, application) != state['pid']:
            raise RuntimeError('Another Office instance is active. Close the trial normally to restore it.')
        if app == 'Access':
            raise RuntimeError('احفظ العمل وأغلق Access يدويًا، ثم اضغط استعادة. لا يمكن التحقق آليًا من حفظ جميع النماذج.')
        collection = getattr(application, {'Word':'Documents', 'PowerPoint':'Presentations', 'Excel':'Workbooks', 'Publisher':'Documents'}[app])
        documents = []
        for document in collection:
            if not document.Saved or not document.Path:
                raise RuntimeError('احفظ كل الملفات المفتوحة في ' + app + ' أولاً ثم اضغط استعادة. لن تُغلق الملفات غير المحفوظة.')
            documents.append(str(document.FullName))
        windows = []
        def collect(handle, _):
            if win32process.GetWindowThreadProcessId(handle)[1] == state['pid'] and win32gui.GetClassName(handle) in WINDOW_CLASSES:
                windows.append(handle)
        win32gui.EnumWindows(collect, None)
        if not windows:
            raise RuntimeError('No Office document window found; close the trial manually')
        setup_windows = []
        def collect_setup(handle, _):
            if win32process.GetWindowThreadProcessId(handle)[1] == state['pid'] and is_setup_dialog(win32gui.GetClassName(handle), win32gui.GetWindowText(handle)):
                setup_windows.append(handle)
        win32gui.EnumWindows(collect_setup, None)
        for handle in setup_windows + windows:
            if win32gui.IsWindow(handle):
                win32gui.PostMessage(handle, 0x10, 0, 0)
        try:
            psutil.Process(state['pid']).wait(timeout=12)
        except psutil.NoSuchProcess:
            pass
        except psutil.TimeoutExpired:
            raise RuntimeError('Office is waiting for a dialog. Complete it; no process was terminated.')
        if process_alive(state):
            raise RuntimeError('Trial process is still running')
        state['status'] = 'restored_by_normal_exit'
        state['reopened_files'] = documents
        save_state(state)
        for document in documents:
            subprocess.Popen([state['executable'], document], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return state
    finally:
        pythoncom.CoUninitialize()

def office_details():
    result = {}
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, C2R_KEY, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            for name in ['VersionToReport', 'Platform', 'UpdatesEnabled']:
                try:result[name] = winreg.QueryValueEx(key, name)[0]
                except FileNotFoundError:pass
    except FileNotFoundError:
        pass
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, POLICY_KEY, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            result['automatic_policy'] = winreg.QueryValueEx(key, 'enableautomaticupdates')[0]
    except FileNotFoundError:
        pass
    return result


def write_update_policy(enabled, hive=None, policy_key=POLICY_KEY, config_key=C2R_KEY):
    if hive is None:
        hive = winreg.HKEY_LOCAL_MACHINE
    access = winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
    with winreg.CreateKeyEx(hive, policy_key, 0, access) as key:
        winreg.SetValueEx(key, 'enableautomaticupdates', 0, winreg.REG_DWORD, int(enabled))
    with winreg.CreateKeyEx(hive, config_key, 0, access) as key:
        winreg.SetValueEx(key, 'UpdatesEnabled', 0, winreg.REG_SZ, 'True' if enabled else 'False')
    return {'status':'updates_allowed' if enabled else 'updates_disabled'}


def update_permission(enabled):
    if ctypes.windll.shell32.IsUserAnAdmin():
        return write_update_policy(enabled)
    from win32com.shell import shell
    import win32event
    import win32api
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    fd, result_path = tempfile.mkstemp(prefix='policy_', suffix='.json', dir=str(STATE_DIR))
    os.close(fd)
    try:
        command = self_command('policy-enable' if enabled else 'policy-disable', '--result', result_path)
        info = shell.ShellExecuteEx(fMask=0x40, lpVerb='runas', lpFile=command[0], lpParameters=subprocess.list2cmdline(command[1:]), nShow=0)
        try:
            while win32event.WaitForSingleObject(info['hProcess'], 250) == 258:
                pass
            code = win32process_exit_code(info['hProcess'])
        finally:
            win32api.CloseHandle(info['hProcess'])
        content = Path(result_path).read_text(encoding='utf-8')
        data = json.loads(content) if content else {}
        if code or data.get('error'):
            raise RuntimeError(data.get('error') or 'لم تُطبق السياسة. اسمح بطلب صلاحية المسؤول ثم أعد المحاولة.')
        return data
    finally:
        Path(result_path).unlink(missing_ok=True)


def win32process_exit_code(handle):
    import win32process
    return win32process.GetExitCodeProcess(handle)


def installation_xml():
    return ('<Configuration>\n'
            '  <Add OfficeClientEdition="32" Channel="Current" Version="' + SUPPORTED_OFFICE_VERSION + '">\n'
            '    <Product ID="O365HomePremRetail">\n'
            '      <Language ID="ar-sa" />\n'
            '      <Language ID="en-us" />\n'
            '    </Product>\n'
            '  </Add>\n'
            '  <Updates Enabled="FALSE" />\n'
            '  <Display Level="Full" AcceptEULA="FALSE" />\n'
            '</Configuration>\n')


def register_open_with():
    if not getattr(sys, 'frozen', False):
        raise RuntimeError('Registration requires the packaged OfficeFileTool.exe')
    source = Path(sys.executable).resolve()
    install_dir = Path(os.environ['LOCALAPPDATA']) / 'Programs/OfficeFileTool' / APP_VERSION
    install_dir.mkdir(parents=True, exist_ok=True)
    destination = install_dir / 'OfficeFileTool.exe'
    if source != destination and (not destination.exists() or hashlib.sha256(source.read_bytes()).digest() != hashlib.sha256(destination.read_bytes()).digest()):
        staging = destination.with_suffix('.new')
        shutil.copy2(source, staging)
        staging.replace(destination)
    exe = str(destination)
    command = '"' + exe + '" "%1"'
    def value(key, name, text, kind=winreg.REG_SZ):
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as handle:
            winreg.SetValueEx(handle, name, 0, kind, text)
    root = 'Software\\Classes\\'
    progid = 'OfficeFileTool.Document'
    value(root + progid, '', 'Office File Tool document')
    value(root + progid + '\\shell\\open\\command', '', command)
    value(root + progid + '\\DefaultIcon', '', exe + ',0')
    application = root + 'Applications\\OfficeFileTool.exe'
    value(application, 'FriendlyAppName', 'Office File Tool')
    value(application + '\\shell\\open\\command', '', command)
    capabilities = 'Software\\OfficeFileTool\\Capabilities'
    value(capabilities, 'ApplicationName', 'Office File Tool')
    value(capabilities, 'ApplicationDescription', 'Open Office files with automatic application detection')
    for extensions in EXTENSIONS.values():
        for ext in extensions:
            value(application + '\\SupportedTypes', ext, '')
            value(root + ext + '\\OpenWithProgids', progid, b'', winreg.REG_NONE)
            value(capabilities + '\\FileAssociations', ext, progid)
            verb = root + 'SystemFileAssociations\\' + ext + '\\shell\\OfficeFileTool'
            value(verb, '', 'فتح باستخدام أداة Office')
            value(verb, 'Icon', exe + ',0')
            value(verb + '\\command', '', command)
    value('Software\\RegisteredApplications', 'OfficeFileTool', capabilities)
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        shortcut = win32com.client.Dispatch('WScript.Shell').CreateShortcut(str(Path(os.environ['APPDATA']) / 'Microsoft/Windows/Start Menu/Programs/Office File Tool.lnk'))
        shortcut.TargetPath = exe
        shortcut.WorkingDirectory = str(destination.parent)
        shortcut.Description = 'Office File Tool'
        shortcut.Save()
    finally:
        pythoncom.CoUninitialize()
    ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
    return {'status':'registered', 'installed_executable':exe}


def unregister_open_with():
    def delete_tree(key):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ | winreg.KEY_WRITE) as handle:
                children = [winreg.EnumKey(handle, i) for i in range(winreg.QueryInfoKey(handle)[0])]
            for child in children:
                delete_tree(key + '\\' + child)
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key)
        except FileNotFoundError:
            pass
    for group in EXTENSIONS.values():
        for ext in group:
            delete_tree('Software\\Classes\\SystemFileAssociations\\' + ext + '\\shell\\OfficeFileTool')
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\Classes\\' + ext + '\\OpenWithProgids', 0, winreg.KEY_SET_VALUE) as key:
                    winreg.DeleteValue(key, 'OfficeFileTool.Document')
            except FileNotFoundError:
                pass
    for key in ['Software\\Classes\\OfficeFileTool.Document', 'Software\\Classes\\Applications\\OfficeFileTool.exe', 'Software\\OfficeFileTool\\Capabilities']:
        delete_tree(key)
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\RegisteredApplications', 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, 'OfficeFileTool')
    except FileNotFoundError:
        pass
    (Path(os.environ['APPDATA']) / 'Microsoft/Windows/Start Menu/Programs/Office File Tool.lnk').unlink(missing_ok=True)
    ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
    return {'status':'unregistered'}


def emit_result(result):
    if RESULT_FILE:
        Path(RESULT_FILE).write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    elif sys.stdout is not None:
        print(json.dumps(result, ensure_ascii=True, indent=2))

def main():
    global RESULT_FILE
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', nargs='?')
    parser.add_argument('--app', choices=list(EXECUTABLES), default='Word')
    parser.add_argument('--file', default='')
    parser.add_argument('--pid', type=int)
    parser.add_argument('--created', type=float)
    parser.add_argument('--result')
    args = parser.parse_args()
    RESULT_FILE = args.result
    if args.action in ('policy-enable', 'policy-disable'):
        if not ctypes.windll.shell32.IsUserAnAdmin():
            raise PermissionError('يلزم تشغيل هذه العملية بصلاحية المسؤول.')
        return emit_result(write_update_policy(args.action == 'policy-enable'))
    if args.action in ('updates-enable', 'updates-disable'):
        return emit_result(update_permission(args.action == 'updates-enable'))
    if args.action == 'watch-dialogs':
        return watch_setup_dialogs(args.app, args.pid, args.created)
    if args.action == 'register':
        return emit_result(register_open_with())
    if args.action == 'unregister':
        return emit_result(unregister_open_with())
    if not args.action:
        return gui(args.app, args.file)
    if args.action not in ('start', 'restore', 'status'):
        path = Path(args.action).resolve()
        return start_trial(choose_app(path), str(path))
    if args.action == 'start':
        result = start_trial(args.app, args.file)
    elif args.action == 'restore':
        result = restore_trial(args.app)
    else:
        path = state_path(args.app)
        result = json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'status': 'no_trial'}
        if 'pid' in result:
            result['process_alive'] = process_alive(result)
    emit_result(result)


def gui(initial_app='Word', initial_file='', auto_start=False):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    root = tk.Tk()
    root.title('Office File Tool')
    root.geometry('860x770')
    root.minsize(800, 730)
    root.configure(bg='#f3f5f7')
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TFrame', background='#ffffff')
    style.configure('TLabel', background='#ffffff', foreground='#172b3a', font=('Segoe UI', 11))
    style.configure('Muted.TLabel', foreground='#526373', font=('Segoe UI', 10))
    style.configure('Heading.TLabel', font=('Segoe UI', 22, 'bold'))
    style.configure('Section.TLabel', font=('Segoe UI', 12, 'bold'))
    style.configure('TButton', padding=(18,10), background='#e9eef2', foreground='#172b3a', borderwidth=0, font=('Segoe UI', 11))
    style.map('TButton', background=[('active','#dce5eb'), ('disabled','#edf0f2')], foreground=[('disabled','#7d8792')])
    style.configure('Primary.TButton', background='#0f766e', foreground='#ffffff', font=('Segoe UI', 11, 'bold'))
    style.map('Primary.TButton', background=[('active','#115e59'), ('disabled','#cad8d6')], foreground=[('disabled','#526373')])
    style.configure('TEntry', fieldbackground='#f8fafc', padding=10, bordercolor='#aebbc5')
    style.configure('TCombobox', fieldbackground='#f8fafc', padding=7)
    style.configure('Status.TLabel', background='#eef5f4', foreground='#19564f', padding=(14,12))
    style.configure('Horizontal.TProgressbar', background='#0f766e', troughcolor='#edf1f3', borderwidth=0)
    app = tk.StringVar(value=initial_app)
    filename = tk.StringVar(value=initial_file)
    status = tk.StringVar(value='جاهز. اختر ملفًا؛ ستتعرّف الأداة على البرنامج المناسب تلقائيًا.')
    hint = tk.StringVar()
    shell = ttk.Frame(root, padding=(28,24))
    shell.pack(fill='both', expand=True, padx=16, pady=16)
    ttk.Label(shell, text='أداة ملفات أوفيس', style='Heading.TLabel').pack(anchor='e')
    ttk.Label(shell, text='افتح ملفك هنا أو من قائمة «فتح باستخدام» في ويندوز.', style='Muted.TLabel').pack(anchor='e', pady=(4,22))
    notebook = ttk.Notebook(shell)
    notebook.pack(fill='both', expand=True)
    files_page = ttk.Frame(notebook, padding=(12,18))
    settings_container = ttk.Frame(notebook)
    notebook.add(files_page, text='  الملفات  ')
    notebook.add(settings_container, text='  الإصدار والتحديثات  ')
    settings_canvas = tk.Canvas(settings_container, background='#ffffff', highlightthickness=0)
    scrollbar = ttk.Scrollbar(settings_container, orient='vertical', command=settings_canvas.yview)
    scrollbar.pack(side='left', fill='y')
    settings_canvas.pack(side='left', fill='both', expand=True)
    settings_canvas.configure(yscrollcommand=scrollbar.set)
    settings_page = ttk.Frame(settings_canvas, padding=(18,18))
    settings_window = settings_canvas.create_window(0, 0, window=settings_page, anchor='nw')
    settings_page.bind('<Configure>', lambda _:settings_canvas.configure(scrollregion=settings_canvas.bbox('all')))
    settings_canvas.bind('<Configure>', lambda event:settings_canvas.itemconfigure(settings_window, width=event.width))
    root.bind_all('<MouseWheel>', lambda event:settings_canvas.yview_scroll(-int(event.delta/120), 'units') if notebook.select()==str(settings_container) else None)
    selection = ttk.Frame(files_page)
    selection.pack(fill='x')
    ttk.Label(selection, text='البرنامج', style='Section.TLabel').pack(side='right', padx=(14,0))
    select = ttk.Combobox(selection, textvariable=app, values=list(EXECUTABLES), state='readonly', width=23, justify='right')
    select.pack(side='right')
    def refresh_hint(*_):
        if app.get() in NORMAL_APPS:
            hint.set('Outlook وOneNote: فتح عادي للملفات، دون وضع التجربة.')
        elif app.get() == 'Access':
            hint.set('Access: احفظ العمل وأغلق البرنامج يدويًا قبل إنهاء التجربة.')
        else:
            hint.set('يمكنك فتح ملفات إضافية في جلسة الأداة نفسها.')
    app.trace_add('write', refresh_hint)
    refresh_hint()
    ttk.Label(files_page, textvariable=hint, style='Muted.TLabel').pack(anchor='e', pady=(8,16))
    ttk.Label(files_page, text='الملف', style='Section.TLabel').pack(anchor='e')
    row = ttk.Frame(files_page)
    row.pack(fill='x', pady=(7,12))
    entry = ttk.Entry(row, textvariable=filename, justify='right')
    entry.pack(side='right', fill='x', expand=True)
    def browse():
        path = filedialog.askopenfilename(title='اختر ملف Office', filetypes=[('Office', ' '.join('*'+ext for group in EXTENSIONS.values() for ext in sorted(group)))])
        if path:
            filename.set(path)
            app.set(choose_app(path))
    browse_button = ttk.Button(row, text='اختيار ملف…', command=browse)
    browse_button.pack(side='left', padx=(0,10))
    actions = ttk.Frame(files_page)
    actions.pack(fill='x', pady=(0,15))
    messages, busy = queue.Queue(), [False]
    def run(action):
        if busy[0]:
            return
        selected, chosen = app.get(), filename.get().strip().strip('"')
        if action == 'start' and chosen:
            try:
                selected = choose_app(chosen)
                app.set(selected)
            except ValueError as error:
                messagebox.showerror('نوع الملف', str(error), parent=root)
                return
        busy[0] = True
        root.configure(cursor='watch')
        for button in (start, restore, browse_button, register, stop_updates, allow_updates):
            button.configure(state='disabled')
        select.configure(state='disabled')
        entry.configure(state='disabled')
        progress.start(14)
        status.set({'start':'جارٍ فتح الملف…', 'restore':'جارٍ استعادة التشغيل الطبيعي…', 'register':'جارٍ الإضافة إلى «فتح باستخدام»…', 'updates-disable':'جارٍ إيقاف تحديثات Office…', 'updates-enable':'جارٍ السماح بالتحديثات…'}[action])
        def worker():
            try:
                command = self_command(action, '--app', selected)
                if action == 'start' and chosen:
                    command += ['--file', chosen]
                STATE_DIR.mkdir(parents=True, exist_ok=True)
                descriptor, output = tempfile.mkstemp(prefix='result_', suffix='.json', dir=str(STATE_DIR))
                os.close(descriptor)
                try:
                    result = subprocess.run(command + ['--result', output], capture_output=True, text=True, encoding='utf-8', errors='replace', creationflags=0x08000000, env=dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUTF8='1'))
                    content = Path(output).read_text(encoding='utf-8')
                    data = json.loads(content) if content else {}
                    if result.returncode or data.get('error'):
                        raise RuntimeError(data.get('error') or result.stderr.strip() or 'لم تكتمل العملية؛ أعد المحاولة.')
                finally:
                    Path(output).unlink(missing_ok=True)
                messages.put(('ok', data))
            except Exception as error:
                messages.put(('error', str(error)))
        threading.Thread(target=worker, daemon=True).start()
    start = ttk.Button(actions, text='فتح الملف / تشغيل', style='Primary.TButton', command=lambda:run('start'))
    start.pack(side='right')
    restore = ttk.Button(actions, text='استعادة الوضع الأصلي', command=lambda:run('restore'))
    restore.pack(side='right', padx=10)
    ttk.Label(files_page, textvariable=status, style='Status.TLabel', wraplength=700, justify='right', anchor='e').pack(fill='x')
    progress = ttk.Progressbar(files_page, mode='indeterminate', style='Horizontal.TProgressbar')
    progress.pack(fill='x', pady=(6,18))
    ttk.Separator(files_page).pack(fill='x')
    integration = ttk.Frame(files_page)
    integration.pack(fill='x', pady=(18,8))
    ttk.Label(integration, text='افتح ملفاتك من الزر الأيمن', style='Section.TLabel').pack(side='right')
    register = ttk.Button(integration, text='إضافة إلى «فتح باستخدام»', command=lambda:run('register'))
    register.pack(side='left')
    ttk.Label(files_page, text='تُحفظ نسخة ثابتة في برامج حسابك، دون تغيير برنامجك الافتراضي.', style='Muted.TLabel').pack(anchor='e')
    footer = ttk.Frame(files_page)
    footer.pack(fill='x', side='bottom', pady=(15,0))
    ttk.Label(footer, text='احفظ ملفاتك قبل الاستعادة. أغلق أي جلسة عادية قبل بدء التجربة.', style='Muted.TLabel').pack(side='right')
    ttk.Label(footer, text='v' + APP_VERSION, style='Muted.TLabel').pack(side='left')
    ttk.Label(settings_page, text='الإصدار المدعوم لوضع التجربة', style='Section.TLabel').pack(anchor='e')
    ttk.Label(settings_page, text='Microsoft 365 • ' + SUPPORTED_OFFICE_VERSION + ' • 32-bit', font=('Segoe UI', 15, 'bold')).pack(anchor='e', pady=(8,6))
    details = office_details()
    installed = details.get('VersionToReport', 'غير مثبت')
    platform = details.get('Platform', 'غير معروف')
    ttk.Label(settings_page, text='المثبت على هذا الجهاز: ' + installed + ' / ' + platform, style='Muted.TLabel').pack(anchor='e')
    ttk.Label(settings_page, text='تتحقق الأداة من بصمات الملفات قبل التشغيل. قد يغيّر تحديث Office الإزاحات؛\nالإصدارات الأخرى ونسخة 64-bit غير مدعومة لوضع التجربة.', style='Muted.TLabel', justify='right').pack(anchor='e', pady=(10,18))
    ttk.Separator(settings_page).pack(fill='x')
    ttk.Label(settings_page, text='تحديثات Office', style='Section.TLabel').pack(anchor='e', pady=(18,6))
    update_status = tk.StringVar()
    def refresh_updates():
        details = office_details()
        disabled = details.get('automatic_policy') == 0 or str(details.get('UpdatesEnabled', '')).lower() == 'false'
        update_status.set('الحالة الحالية: التحديثات التلقائية متوقفة' if disabled else 'الحالة الحالية: التحديثات التلقائية مسموحة أو غير محددة')
    refresh_updates()
    ttk.Label(settings_page, textvariable=update_status, style='Muted.TLabel').pack(anchor='e')
    update_row = ttk.Frame(settings_page)
    update_row.pack(fill='x', pady=(14,10))
    stop_updates = ttk.Button(update_row, text='إيقاف التحديثات التلقائية', command=lambda:run('updates-disable'))
    stop_updates.pack(side='right')
    allow_updates = ttk.Button(update_row, text='السماح بالتحديثات مجددًا', command=lambda:run('updates-enable'))
    allow_updates.pack(side='right', padx=10)
    ttk.Label(settings_page, text='سياسة محلية تستمر بعد إعادة التشغيل وتتطلب موافقة المسؤول.\nزر السماح يغيّر الإعداد فقط؛ لا يستدعي أداة التحديث أو مهمة فحص.', style='Muted.TLabel', justify='right').pack(anchor='e')
    ttk.Separator(settings_page).pack(fill='x', pady=18)
    ttk.Label(settings_page, text='تثبيت الإصدار نفسه على جهاز آخر', style='Section.TLabel').pack(anchor='e')
    ttk.Label(settings_page, text='استخدم Office Deployment Tool مع ملف الإعداد المحدد بهذا الإصدار.\nقد لا يبقى الإصدار القديم متاحًا للتنزيل؛ احتفظ بحزمة التثبيت بعد تنزيلها.', style='Muted.TLabel', justify='right').pack(anchor='e', pady=(8,12))
    def export_xml():
        target = filedialog.asksaveasfilename(title='حفظ إعداد تثبيت نفس الإصدار', initialfile='Office-' + SUPPORTED_OFFICE_VERSION + '-32bit.xml', defaultextension='.xml', filetypes=[('Office configuration','*.xml')])
        if target:
            Path(target).write_text(installation_xml(), encoding='utf-8')
            messagebox.showinfo('إعداد التثبيت', 'حُفظ ملف الإعداد.\nضعه بجانب setup.exe الخاص بـ Office Deployment Tool.\nنفّذ أولًا: setup.exe /download "' + Path(target).name + '"\nثم: setup.exe /configure "' + Path(target).name + '"', parent=root)
    ttk.Button(settings_page, text='حفظ إعداد تثبيت الإصدار نفسه', command=export_xml).pack(anchor='e')
    def poll():
        try:
            kind, data = messages.get_nowait()
            busy[0] = False
            root.configure(cursor='')
            progress.stop()
            for button in (start, restore, browse_button, register, stop_updates, allow_updates):
                button.configure(state='normal')
            select.configure(state='readonly')
            entry.configure(state='normal')
            if kind == 'error':
                status.set('تعذّرت العملية. راجع الرسالة ثم أعد المحاولة.')
                messagebox.showerror('Office File Tool', data, parent=root)
            else:
                refresh_updates()
                status.set({'applied':'اكتمل التشغيل. ملفك جاهز للاستخدام.',
                    'opened_in_existing_session':'فُتح الملف في جلسة الأداة الحالية.',
                    'normal_open':'فُتح الملف بالتشغيل العادي.',
                    'normal_open_no_trial':'هذا البرنامج يعمل بالتشغيل العادي؛ لا توجد تجربة لاستعادتها.',
                    'no_trial':'لا توجد جلسة تجربة لاستعادتها.',
                    'updates_disabled':'أُوقفت التحديثات التلقائية حتى تعيد السماح بها.',
                    'updates_allowed':'أُعيد السماح بالتحديثات. لم يُشغَّل فحص أو تنزيل تحديث.',
                    'registered':'تمت الإضافة. اضغط على أي ملف Office بالزر الأيمن واختر Office File Tool.'}.get(data['status'], 'اكتملت الاستعادة إلى التشغيل الطبيعي.'))
        except queue.Empty:
            pass
        root.after(150, poll)
    def close():
        if busy[0]:
            messagebox.showinfo('Office File Tool', 'انتظر اكتمال العملية قبل إغلاق الأداة.', parent=root)
        else:
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    root.bind('<Control-o>', lambda _:browse())
    root.after(150, poll)
    if auto_start:
        root.after(250, lambda:run('start'))
    root.mainloop()

if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        if RESULT_FILE:
            emit_result({'error':str(error)})
        elif getattr(sys, 'frozen', False) and (len(sys.argv) == 1 or sys.argv[1] not in ('start', 'restore', 'status', 'watch-dialogs')):
            ctypes.windll.user32.MessageBoxW(None, str(error), 'Office File Tool', 0x10)
        elif sys.stderr is not None:
            print(str(error), file=sys.stderr)
        sys.exit(1)
