# -*- coding: utf-8 -*-
"""Lưu khóa API Anthropic của người dùng (đề xuất hoàn thiện 15/11/2026, T5) – trong kho khóa của hệ điều hành:
  • Windows: Credential Manager (advapi32 CredWriteW/CredReadW/CredDeleteW qua ctypes);
  • macOS: Keychain (lệnh /usr/bin/security);
  • hệ khác (hoặc khi kho khóa lỗi): tệp ~/.battu_thieu_vi_hoa/khoa_api, quyền chỉ chủ tài khoản đọc (0600).
Biến môi trường ANTHROPIC_API_KEY (nếu có) được ưu tiên khi đọc. Khóa không bao giờ ghi vào kho lá số, nhật ký hay bản xuất."""
import os, sys, subprocess

DICH_VU = 'BatTuThieuViHoa-AnthropicAPI'
TAI_KHOAN = 'nguoi_dung'


def _tep():
    d = os.path.join(os.path.expanduser('~'), '.battu_thieu_vi_hoa'); os.makedirs(d, exist_ok=True)
    return os.path.join(d, 'khoa_api')


# ---------------------------------------------------------------- Windows
def _win():
    import ctypes
    from ctypes import wintypes

    class CREDENTIALW(ctypes.Structure):
        _fields_ = [('Flags', wintypes.DWORD), ('Type', wintypes.DWORD), ('TargetName', wintypes.LPWSTR),
                    ('Comment', wintypes.LPWSTR), ('LastWritten', wintypes.FILETIME), ('CredentialBlobSize', wintypes.DWORD),
                    ('CredentialBlob', ctypes.POINTER(ctypes.c_char)), ('Persist', wintypes.DWORD), ('AttributeCount', wintypes.DWORD),
                    ('Attributes', ctypes.c_void_p), ('TargetAlias', wintypes.LPWSTR), ('UserName', wintypes.LPWSTR)]
    adv = ctypes.WinDLL('advapi32', use_last_error=True)
    return ctypes, CREDENTIALW, adv


def _win_ghi(khoa):
    ctypes, C, adv = _win()
    b = khoa.encode('utf-16-le'); buf = ctypes.create_string_buffer(b, len(b))
    c = C(); c.Type = 1; c.TargetName = DICH_VU; c.UserName = TAI_KHOAN; c.Persist = 2
    c.CredentialBlobSize = len(b); c.CredentialBlob = ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))
    if not adv.CredWriteW(ctypes.byref(c), 0): raise OSError(ctypes.get_last_error())


def _win_doc():
    ctypes, C, adv = _win()
    p = ctypes.POINTER(C)()
    if not adv.CredReadW(DICH_VU, 1, 0, ctypes.byref(p)): return None
    try:
        n = p.contents.CredentialBlobSize
        return ctypes.string_at(p.contents.CredentialBlob, n).decode('utf-16-le')
    finally:
        adv.CredFree(p)


def _win_xoa():
    ctypes, C, adv = _win(); adv.CredDeleteW(DICH_VU, 1, 0)


# ---------------------------------------------------------------- macOS
def _mac(*a, nhap=None):
    return subprocess.run(['/usr/bin/security'] + list(a), capture_output=True, text=True, input=nhap, timeout=20)


# ---------------------------------------------------------------- chung
def ghi(khoa):
    """Lưu khóa → nơi đã lưu ('Credential Manager' | 'Keychain' | 'tệp riêng tư')."""
    khoa = (khoa or '').strip()
    if not khoa: xoa(); return ''
    try:
        if sys.platform == 'win32': _win_ghi(khoa); return 'Credential Manager của Windows'
        if sys.platform == 'darwin':
            r = _mac('add-generic-password', '-U', '-s', DICH_VU, '-a', TAI_KHOAN, '-w', khoa)
            if r.returncode == 0: return 'Keychain của macOS'
    except Exception:
        pass
    p = _tep()
    fd = os.open(p, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, 'w', encoding='utf-8') as f: f.write(khoa)
    try: os.chmod(p, 0o600)
    except OSError: pass
    return 'tệp riêng tư (%s)' % p


def doc():
    """Khóa đã lưu, hoặc biến môi trường ANTHROPIC_API_KEY; không có → ''."""
    k = os.environ.get('ANTHROPIC_API_KEY', '').strip()
    if k: return k
    try:
        if sys.platform == 'win32':
            k = _win_doc()
            if k: return k
        elif sys.platform == 'darwin':
            r = _mac('find-generic-password', '-s', DICH_VU, '-a', TAI_KHOAN, '-w')
            if r.returncode == 0 and r.stdout.strip(): return r.stdout.strip()
    except Exception:
        pass
    p = _tep()
    return open(p, encoding='utf-8').read().strip() if os.path.exists(p) else ''


def xoa():
    try:
        if sys.platform == 'win32': _win_xoa()
        elif sys.platform == 'darwin': _mac('delete-generic-password', '-s', DICH_VU, '-a', TAI_KHOAN)
    except Exception:
        pass
    p = _tep()
    if os.path.exists(p): os.remove(p)


def che(khoa):
    """Hiện khóa đã che: 'sk-ant-…a1b2'."""
    return (khoa[:7] + '…' + khoa[-4:]) if khoa and len(khoa) > 12 else ('đã lưu' if khoa else '')
