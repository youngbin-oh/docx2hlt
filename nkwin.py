"""nkwin.py — drive NKEditor to turn an imported doc into a saved .hlt.

NKEditor builds the structured patent document when it *opens* a KIPO
PatentCAFDOC .xml passed on the command line (verified: full element tree,
claims, drawings, paragraph identifiers). The only reliable, side-effect-free
way to get that document back out as .hlt is its own "다른 이름으로 저장"
(Save As) path, which for an imported doc always produces the *.HLT type.

This module automates that: launch NKEditor on the .xml, wait for the document
to build its element tree, then trigger "다른 이름으로 저장" by posting
WM_COMMAND ID_FILE_SAVE_AS (0xE104) to the frame and BM_CLICK the dialog's
저장 button (which writes <base>.hlt into Data\\Hlt). We then move the result
wherever we want.

Everything here is window-message based (EnumWindows / PostMessage /
SendMessage) rather than synthetic mouse/keyboard input, because the build
host is reached over a remote desktop where injected clicks are unreliable.
Only ASCII identifiers are matched — class names, control IDs and the doc
filename — never Korean window captions, since NKEditor's frame is an ANSI
window and the host's non-Unicode codepage is not Korean.
"""
import ctypes
import os
import re
import subprocess
import time
from ctypes import wintypes

_user = ctypes.windll.user32

_CB = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

DEBUG = os.environ.get('NKWIN_DEBUG', '') != ''


def _log(msg):
    if DEBUG:
        import sys as _s
        print('  [nkwin] ' + msg, file=_s.stderr, flush=True)

NK_EXE = r'C:\KipoNet\NKEditor\NKEditor.exe'
NK_HLT_DIR = r'C:\KipoNet\NKEditor\Data\Hlt'

# Why the last nk_export_rtf() returned None (None after a success), and
# whether NKEditor died under it — the caller retries only a crash
# (heap-corruption exits were seen right after an NKEditor auto-update).
LAST_EXPORT_ERROR = None
LAST_EXPORT_CRASHED = False


def _wpid(h):
    """PID that owns window `h`."""
    pid = wintypes.DWORD()
    _user.GetWindowThreadProcessId(h, ctypes.byref(pid))
    return pid.value


def _launch(path):
    """Start a private NKEditor on `path`, minimized and without activation.
    NKEditor 6.3 is not single-instance - every launch is a new process with
    its own frame - so a normal launch pops a second editor window over the
    one the user already has open. Everything here is message-driven, so the
    minimized frame (and its unfocused Save-As dialog) works the same."""
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 7                       # SW_SHOWMINNOACTIVE
    return subprocess.Popen([NK_EXE, path], startupinfo=si)


def _exe_name(pid):
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(0x1000, False, pid)   # QUERY_LIMITED_INFORMATION
    if not h:
        return ''
    try:
        buf = ctypes.create_unicode_buffer(520)
        n = wintypes.DWORD(520)
        if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)):
            return os.path.basename(buf.value)
        return ''
    finally:
        k32.CloseHandle(h)


def _cmdlines(pids):
    """{pid: command line} via CIM (one PowerShell call)."""
    if not pids:
        return {}
    flt = ' or '.join('ProcessId=%d' % p for p in pids)
    r = subprocess.run(
        ['powershell', '-NoProfile', '-Command',
         '[Console]::OutputEncoding=[Text.Encoding]::UTF8; '
         'Get-CimInstance Win32_Process -Filter "%s" | '
         '%% { "$($_.ProcessId)|$($_.CommandLine)" }' % flt],
        capture_output=True, timeout=30)
    out = {}
    for line in r.stdout.decode('utf-8', 'replace').splitlines():
        p, _, cmd = line.partition('|')
        if p.strip().isdigit():
            out[int(p)] = cmd
    return out


def editors_showing(path):
    """PIDs of running NKEditor frames whose document is `path`. The frame
    title ends in '[<doc path>]' but NKEditor cuts long paths short, so a
    cut title only counts when the process was also launched on `path`.
    (On a host whose ANSI codepage is not Korean, Korean title text reads
    back as '?' and nothing matches.)"""
    want = os.path.abspath(path).lower()
    full, cut = [], []
    t = ctypes.create_unicode_buffer(1024)
    for h, cn, _t, p in _top_windows():
        if not cn.startswith('Afx:') or p in full or p in cut:
            continue
        _user.GetWindowTextW(h, t, 1024)
        m = re.search(r'\[([^\]]+)\]?\s*$', t.value)
        if not m or _exe_name(p).lower() != 'nkeditor.exe':
            continue
        shown = m.group(1).strip().lower()
        if shown == want:
            full.append(p)
        elif len(shown) >= 12 and want.startswith(shown):
            cut.append(p)
    cmds = _cmdlines(cut)
    return full + [p for p in cut if want in cmds.get(p, '').lower()]


def _top_windows():
    out = []

    def cb(h, l):
        if _user.IsWindowVisible(h):
            cn = ctypes.create_unicode_buffer(64)
            _user.GetClassNameW(h, cn, 64)
            t = ctypes.create_unicode_buffer(200)
            _user.GetWindowTextW(h, t, 200)
            out.append((h, cn.value, t.value, _wpid(h)))
        return True

    _user.EnumWindows(_CB(cb), 0)
    return out


def _find_nk_window(pid=None, match=None):
    """Return the NKEditor main frame hwnd. If `pid` is given, only consider
    windows owned by that process. If `match` is given, prefer the window
    whose title contains it (e.g. the document filename).

    NOTE: the frame is an ANSI window and the system non-Unicode codepage is
    not Korean, so Korean title text comes back as '?'. We therefore match
    only on the ASCII bits — the 'Afx:' class and an ASCII `match` (the doc
    filename)."""
    best = None
    for h, cn, t, p in _top_windows():
        if pid is not None and p != pid:
            continue
        if cn.startswith('Afx:'):
            if match:
                if match in t:
                    return h
            elif best is None:
                best = h
    return best


def _has_child_class(hwnd, cls):
    """True if `hwnd` has a (recursive) child window of class `cls`."""
    found = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == cls:
            found.append(c)
        return True

    _user.EnumChildWindows(hwnd, _CB(cb), 0)
    return bool(found)


def _find_saveas_dialogs(pid=None):
    """All '다른 이름으로 저장' (Save As) common dialogs of `pid`, if open.

    A real Save-As dialog needs BOTH the DirectUIHWND shell view AND the
    filename ComboBoxEx32 — a DirectUI overwrite-confirm lacks the combo,
    so requiring both keeps it from being mistaken for the file dialog.
    (Several dialogs can stack when re-triggered, so return them all.)"""
    out = []
    for h, cn, t, p in _top_windows():
        if pid is not None and p != pid:
            continue
        if cn == '#32770' and _has_child_class(h, 'DirectUIHWND') \
                and _has_child_class(h, 'ComboBoxEx32'):
            out.append(h)
    return out


def _find_saveas_dialog(pid=None):
    """First open Save-As dialog of `pid`, or None."""
    ds = _find_saveas_dialogs(pid)
    return ds[0] if ds else None


def _modal_up(pid):
    """True while pid shows a modal prompt — a #32770 that is NOT a save-as
    dialog (overwrite confirm, error box, recovery prompt). While one is up,
    SendMessage to any of that thread's windows blocks forever — the modal
    message pump swallows it — so callers must answer it, not click through."""
    for h, cn, t, p in _top_windows():
        if p == pid and cn == '#32770' \
                and not _has_child_class(h, 'ComboBoxEx32'):
            return h
    return None


def _find_button_id(dlg, ctrl_id):
    found = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'Button' and _user.GetDlgCtrlID(c) == ctrl_id \
                and _user.IsWindowEnabled(c):
            found.append(c)
        return True

    _user.EnumChildWindows(dlg, _CB(cb), 0)
    return found[0] if found else None


def _find_button_text(dlg, keywords):
    """Return the first Button child whose caption (mnemonic stripped)
    contains any of `keywords`. Dialog buttons are Unicode so Korean captions
    read correctly."""
    found = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'Button' and _user.IsWindowEnabled(c):
            t = ctypes.create_unicode_buffer(80)
            _user.GetWindowTextW(c, t, 80)
            cap = t.value.replace('&', '').split('(')[0].strip()
            for kw in keywords:
                if kw in cap:
                    found.append(c)
                    return True
        return True

    _user.EnumChildWindows(dlg, _CB(cb), 0)
    return found[0] if found else None


def _dismiss_nk_dialogs(pid, prefer='yes', ok_paths=None):
    """Click a button on the nk process's modal message-box dialogs to clear
    recovery / file-not-found / overwrite prompts. `ok_paths` (save loops
    only): basenames/abspaths an overwrite-confirm may legitimately target —
    a confirm naming anything else gets 아니요, not 예.

    Only plain message boxes are touched — a Save-As common dialog (which has
    a SHELLDLL_DefView file list) is never dismissed here. `prefer` picks which
    button by caption: 'yes' = 예/확인 (overwrite confirms, error acks);
    'no' = 아니요/취소 (to skip crash-recovery on launch).

    Note: Vista+ overwrite-confirm and XML-conversion prompts DO contain a
    DirectUIHWND child — they're still plain button dialogs, so the file-list
    class is the only safe discriminator."""
    yes = ('예', '확인', 'Yes', 'OK')
    no = ('아니요', '아니오', '취소', 'No', 'Cancel')
    want = yes if prefer == 'yes' else no
    other = no if prefer == 'yes' else yes
    hit = []
    for h, cn, t, p in _top_windows():
        if p != pid or cn != '#32770' \
                or _has_child_class(h, 'SHELLDLL_DefView'):
            continue
        # Overwrite-confirm safety. When ok_paths is provided, a text-less
        # DirectUI confirm (예/아니요 buttons, no readable Static) can't be
        # proven to target our own output — deny it. Dialogs carrying text
        # are ordinary prompts and follow `prefer`.
        eff = want
        if prefer == 'yes' and ok_paths is not None:
            if not _dlg_texts(h) and _find_button_text(h, yes) \
                    and _find_button_text(h, no):
                eff = no
                _log('unverifiable confirm (DirectUI) -> 아니요')
        btn = (_find_button_text(h, eff) or _find_button_text(h, other)
               or _find_button_id(h, 1) or _find_button_id(h, 2)
               or _find_button_id(h, 6) or _find_button_id(h, 7))
        if btn:
            _log('dismiss dlg %s (%r) prefer=%s btn %s'
                 % (hex(h), t, prefer, hex(btn)))
            _user.SendMessageW(btn, 0x00F5, 0, 0)   # BM_CLICK
            hit.append(h)
    return hit


def _dlg_texts(dlg):
    """All Static/SysLink caption texts inside a dialog — used to read which
    file an overwrite-confirm prompt is asking about."""
    texts = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value in ('Static', 'SysLink'):
            t = ctypes.create_unicode_buffer(600)
            _user.GetWindowTextW(c, t, 600)
            if t.value:
                texts.append(t.value)
        return True

    _user.EnumChildWindows(dlg, _CB(cb), 0)
    return texts


def _find_filename_edit(dlg):
    """The filename Edit inside a common Save-As dialog: it lives at the end
    of the ComboBoxEx32 > ComboBox > Edit chain (the dialog also hosts other
    Edit controls, e.g. the search box). Returns the Edit hwnd."""

    def child_class(hwnd, cls):
        hit = []
        def cb(c, l):
            cn = ctypes.create_unicode_buffer(64)
            _user.GetClassNameW(c, cn, 64)
            if cn.value == cls:
                hit.append(c)
            return True
        _user.EnumChildWindows(hwnd, _CB(cb), 0)
        return hit[0] if hit else None

    # prefer the filename combo's Edit; fall back to the last Edit in the dlg
    for combo in (child_class(dlg, 'ComboBoxEx32'),):
        if not combo:
            continue
        cb2 = child_class(combo, 'ComboBox')
        if cb2:
            edit = child_class(cb2, 'Edit')
            if edit:
                return edit
    edits = []
    def cb2(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'Edit':
            edits.append(c)
        return True
    _user.EnumChildWindows(dlg, _CB(cb2), 0)
    return edits[-1] if edits else None


def _dialog_dir(dlg):
    """Current folder shown by a common Save-As dialog. Primary source is
    CDM_GETFOLDERPATH (WM_USER+102) — the API intended for this, which still
    resolves a real filesystem path when the address band only shows a
    library label like '문서'. Falls back to scraping the address band."""
    if dlg:
        buf = ctypes.create_unicode_buffer(520)
        n = _user.SendMessageW(dlg, 0x466, 520, buf)   # CDM_GETFOLDERPATH
        if n and n > 0 and buf.value:
            return buf.value
    dirs = []
    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'ToolbarWindow32':
            t = ctypes.create_unicode_buffer(300)
            _user.GetWindowTextW(c, t, 300)
            m = re.search(r'([A-Za-z]:[\\/].+)$', t.value)
            if m:
                dirs.append(m.group(1).strip())
        return True
    _user.EnumChildWindows(dlg, _CB(cb), 0)
    # the longest drive path wins (address band shows the deepest location)
    return max(dirs, key=len) if dirs else None


def _find_save_button(dlg):
    """The 저장/예 default button in a dialog = the enabled Button with
    ctrl-id 1. Returns its hwnd."""
    found = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'Button' and _user.GetDlgCtrlID(c) == 1 \
                and _user.IsWindowEnabled(c):
            found.append(c)
        return True

    _user.EnumChildWindows(dlg, _CB(cb), 0)
    return found[0] if found else None


def _send_input(seq):
    """seq = [(wVk, wScan, flags)] — flags 4=UNICODE, 2=KEYUP."""
    class KI(ctypes.Structure):
        _fields_ = [('wVk', wintypes.WORD), ('wScan', wintypes.DWORD),
                    ('dwFlags', wintypes.DWORD), ('time', wintypes.DWORD),
                    ('ex', ctypes.POINTER(ctypes.c_ulong))]

    class INP(ctypes.Structure):
        _fields_ = [('type', wintypes.DWORD), ('ki', KI),
                    ('pad', ctypes.c_byte * 16)]

    arr = []
    for vk, scan, fl in seq:
        i = INP()
        i.type = 1
        i.ki.wVk = vk
        i.ki.wScan = scan
        i.ki.dwFlags = fl
        arr.append(i)
    _user.SendInput(len(arr), (INP * len(arr))(*arr), ctypes.sizeof(INP))


def _keypress(vk, mods=()):
    """Press vk (optionally with held modifiers, e.g. (0x11,) for Ctrl)."""
    seq = [(m, 0, 0) for m in mods]
    seq += [(vk, 0, 0), (vk, 0, 2)]
    seq += [(m, 0, 2) for m in reversed(mods)]
    _send_input(seq)


def _wait_for(pred, timeout, step=0.4):
    t0 = time.time()
    while time.time() - t0 < timeout:
        v = pred()
        if v:
            return v
        time.sleep(step)
    return None


def _tree_item_count(hwnd):
    """Item count of the document's structure tree (SysTreeView32).
    >0 means the imported doc finished building its element tree."""
    counts = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'SysTreeView32':
            counts.append(_user.SendMessageW(c, 0x1105, 0, 0))  # TVM_GETCOUNT
        return True

    _user.EnumChildWindows(hwnd, _CB(cb), 0)
    return max(counts) if counts else 0


def nk_import_and_save_hlt(xml_path, base_name, launch_timeout=45,
                           save_timeout=30, keep_open=False, out_dir=None):
    """Open `xml_path` in NKEditor, drive Save-As to produce a .hlt, and return
    the saved .hlt path. Returns None on failure.

    base_name: desired filename stem — the generated <base>.hlt.
    out_dir: directory to save into (default NK Data\\Hlt). The full path is
             typed into the filename field so the dialog's last-used folder
             can't redirect the output somewhere else.
    keep_open: leave NKEditor running on the saved doc (nice for inspection);
             otherwise it is closed after the save.
    """
    xml_path = os.path.abspath(xml_path)
    out_dir = out_dir or NK_HLT_DIR
    t0_launch = time.time()
    target = os.path.join(out_dir, base_name + '.hlt')
    # Clear every copy we might overwrite — target AND the dialog's default
    # dir (Data\Hlt). With ours gone, an overwrite-confirm can only ever be
    # about a foreign file, which the save loop answers 아니요 to.
    for cand in (target, target[:-4] + '.HLT',
                 os.path.join(NK_HLT_DIR, base_name + '.hlt'),
                 os.path.join(NK_HLT_DIR, base_name + '.HLT')):
        try:
            if os.path.exists(cand):
                os.remove(cand)
        except OSError:
            pass
        except OSError:
            pass

    _log('launching NKEditor on %s' % xml_path)
    proc = _launch(xml_path)
    pid = proc.pid
    _log('nk pid=%d' % pid)

    # wait for the main document window to appear, dismissing any of THIS
    # process's modal prompts (recovery / file-not-found) that pop up first.
    xbase = os.path.basename(xml_path)
    hwnd = None
    t0 = time.time()
    while time.time() - t0 < launch_timeout and not hwnd:
        _dismiss_nk_dialogs(pid, prefer='no')   # skip crash-recovery prompts
        hwnd = _find_nk_window(pid, xbase) or _find_nk_window(pid)
        if not hwnd and time.time() - t0 > 6:
            # Older NKEditor builds were single-instance (6.3 starts a new
            # process per launch): if it was already running, our Popen'd
            # launcher handed the file to the existing process, so the new doc
            # window belongs to a *different* pid. Fall back to matching the
            # filename in the title across every NKEditor process — and
            # dismiss that pid's prompts too (a crash-recovery dialog on the
            # user's instance blocks our handoff just the same).
            hwnd = _find_nk_window(None, xbase)
            if hwnd:
                _dismiss_nk_dialogs(_wpid(hwnd), prefer='no')
        else:
            # our own pid has no window yet — but an existing instance may
            # already have grabbed the file. A stale crash-recovery prompt
            # on that instance blocks the handoff; dismiss ONLY prompts
            # whose text looks like auto-recovery (never touch generic
            # save/discard prompts that might belong to the user's own
            # unsaved work).
            for h, cn, _t, p in _top_windows():
                if cn == '#32770' and p != pid \
                        and any('복구' in s for s in _dlg_texts(h)):
                    _dismiss_nk_dialogs(p, prefer='no')
        time.sleep(0.4)
    _log('nk window=%s' % (hex(hwnd) if hwnd else None))
    # dialogs and the save path are all keyed to whichever pid owns the doc
    # window (it may differ from the launched pid in the single-instance case).
    pid = _wpid(hwnd) if hwnd else pid
    if not hwnd:
        return None
    # Wait until the doc finished building its element tree (structure pane
    # populated) — a large spec can take several seconds. Recovery / error
    # prompts can also appear here and block the import, so keep dismissing
    # them (prefer '아니요' — we want the file we opened, not a recovered one).
    loaded = None
    t0 = time.time()
    tbuf = ctypes.create_unicode_buffer(512)
    while time.time() - t0 < launch_timeout and not loaded:
        _dismiss_nk_dialogs(pid, prefer='no')
        if _tree_item_count(hwnd) > 2:
            loaded = True
            break
        # v6.3 keeps the structure pane collapsed and the tree empty even
        # with the document fully imported — the frame title gaining the
        # [path] suffix is the only load marker there.
        _user.GetWindowTextW(hwnd, tbuf, 512)
        if '[' in tbuf.value:
            loaded = True
            break
        time.sleep(0.5)
    _log('tree loaded=%s count=%d' % (bool(loaded), _tree_item_count(hwnd)))
    time.sleep(0.5)

    # Trigger "다른 이름으로 저장" via WM_COMMAND ID_FILE_SAVE_AS (0xE104)
    # posted straight to the frame. This is deterministic — it needs no input
    # focus and is unaffected by overlapping (remote-desktop) windows, unlike
    # screen-coordinate clicks on the quick-access ribbon icon. Retry a few
    # times in case the doc is still settling. The shell dialog can take
    # ~7 s to materialize on a slow host, so each attempt waits a while.
    dlg = None
    for _attempt in range(6):
        _user.PostMessageW(hwnd, 0x0111, 0xE104, 0)   # WM_COMMAND SAVE_AS
        dlg = _wait_for(lambda: _find_saveas_dialog(pid), 12)
        _log('save-as attempt %d -> dlg %s'
             % (_attempt, hex(dlg) if dlg else None))
        if dlg:
            break
        time.sleep(1)
    if not dlg:
        return None
    time.sleep(0.8)   # let the DirectUI dialog finish initializing
    # The dialog's starting folder follows the instance's last-used directory
    # (e.g. the user's Desktop when our launch was handed to an already-running
    # NKEditor). The filename Edit is DirectUI-shimmed: WM_SETTEXT is refused
    # and WM_PASTE can leave a mangled name, so instead of forcing a path we
    # read the dialog's folder and watch that exact file, then move it.
    dlg_dir = _dialog_dir(dlg)
    dlg_file = os.path.join(dlg_dir, base_name + '.hlt') if dlg_dir else None
    _log('dialog dir=%r -> file=%r' % (dlg_dir, dlg_file))
    # The filename box can't be driven reliably (DirectUI), so the save lands
    # in the dialog's own folder. If a same-named file already sits there and
    # predates this run, it's foreign — refuse rather than overwrite it.
    foreign = bool(dlg_file) and os.path.exists(dlg_file) \
        and os.path.getmtime(dlg_file) < t0_launch
    if foreign:
        _log('save-as cancelled: %r predates this run' % dlg_file)
        print('WARNING: save-as aborted — %s exists outside the target '
              'folder' % dlg_file)
        cb = _find_button_text(dlg, ('취소', 'Cancel'))
        if cb:
            _user.SendMessageW(cb, 0x00F5, 0, 0)
        return None
    # trigger the 저장 button: prefer ctrl-id 1, else find it by caption.
    # PostMessage only — SendMessage here deadlocks whenever a nested modal
    # (overwrite confirm) is pumping the same thread.
    btn = _find_save_button(dlg) or _find_button_text(dlg, ('저장', 'Save'))
    _log('save button=%s' % (hex(btn) if btn else None))
    if btn:
        _user.PostMessageW(btn, 0x00F5, 0, 0)    # BM_CLICK
    else:
        _user.PostMessageW(dlg, 0x0111, 1, 0)    # WM_COMMAND IDOK
    def _produced():
        # When the dialog shows a library ('문서') rather than a real path,
        # _dialog_dir can't resolve it — so sweep every plausible folder by
        # name+recency instead of trusting one source.
        # dlg_dir too: once the dialog has closed it can't be re-read, and
        # its last-used folder (another case's Sent, say) is where the file
        # actually lands
        dirs = {out_dir, NK_HLT_DIR, os.path.dirname(xml_path),
                os.path.expanduser('~\\Documents'),
                os.path.expanduser('~\\Desktop'),
                os.path.expanduser('~\\Downloads')}
        if dlg_dir:
            dirs.add(dlg_dir)
        dd = _dialog_dir(_find_saveas_dialog(pid) or dlg)
        if dd:
            dirs.add(dd)
        for d in dirs:
            for cand in (os.path.join(d, base_name + '.hlt'),
                         os.path.join(d, base_name + '.HLT')):
                if os.path.exists(cand) and os.path.getsize(cand) > 0 \
                        and os.path.getmtime(cand) >= t0_launch - 5:
                    return cand
        return None

    def _dlg_alive(h):
        return h and _user.IsWindow(h)

    # the native write can take several seconds on a big spec; keep nudging
    # the save-as 저장 button and confirm overwrite prompts until the file
    # lands or we time out.
    saved = None
    t0 = time.time()
    while time.time() - t0 < save_timeout and not saved:
        saved = _produced()
        if saved:
            break
        # answer any pending modal prompt FIRST (overwrite confirm -> 예:
        # foreign same-named files were already rejected before 저장 was
        # clicked, so a confirm now can only concern this run's own output).
        _dismiss_nk_dialogs(pid, prefer='yes')
        # re-click 저장 on every still-open save-as dialog if the first
        # click didn't take — but NEVER while a modal is up, and only via
        # PostMessage: SendMessage to the modal-blocked thread would hang
        # us dead.
        if not _modal_up(pid):
            for d in _find_saveas_dialogs(pid):
                b = (_find_save_button(d)
                     or _find_button_text(d, ('저장', 'Save')))
                if b:
                    _user.PostMessageW(b, 0x00F5, 0, 0)   # BM_CLICK
                else:
                    _user.PostMessageW(d, 0x0111, 1, 0)   # WM_COMMAND IDOK
        time.sleep(0.8)
    if saved:
        _settled(saved)

    if not keep_open and pid == proc.pid:
        # Only close the frame when it belongs to the process WE launched.
        # If the import was handed to an already-running (user) instance the
        # hwnd is that user's frame — closing it would discard their documents.
        try:
            _user.PostMessageW(hwnd, 0x0010, 0, 0)   # WM_CLOSE
            _close_with_prompts(pid, hwnd)
        except Exception:
            pass
    return saved


def _settled(path, quiet=1.5, limit=60):
    """Wait until `path` stops growing: the save loops see the file as soon
    as NKEditor creates it, and an RTF export of a big document is still
    being written then (a 영문 PCT spec) - a WM_CLOSE sent that early is
    answered by the 'XML 변환' prompt only after the close loop gave up."""
    t0 = time.time()
    last = -1
    while time.time() - t0 < limit:
        try:
            n = os.path.getsize(path)
        except OSError:
            return
        if n == last:
            return
        last = n
        time.sleep(quiet)


def _close_with_prompts(pid, hwnd, timeout=30):
    """After WM_CLOSE, keep answering NKEditor's delayed close prompts
    ('XML 변환하시겠습니까?' etc.) with 아니요 until the frame is gone —
    a single dismiss misses prompts that appear a beat after the close."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        _dismiss_nk_dialogs(pid, prefer='no')
        if not _user.IsWindow(hwnd):
            break
        time.sleep(0.6)


def _type_combos(dlg):
    """File-type ComboBoxes in a Save-As dialog: the bottom combos that are
    plain ComboBox (not wrapped in ComboBoxEx32 — that one is the filename).
    Returns list of hwnd."""
    combos = []

    def cb(c, l):
        cn = ctypes.create_unicode_buffer(64)
        _user.GetClassNameW(c, cn, 64)
        if cn.value == 'ComboBox':
            par = ctypes.create_unicode_buffer(64)
            _user.GetClassNameW(_user.GetParent(c), par, 64)
            if par.value != 'ComboBoxEx32':
                combos.append(c)
        return True

    _user.EnumChildWindows(dlg, _CB(cb), 0)
    return combos


def _combo_items(combo):
    items = []
    n = _user.SendMessageW(combo, 0x0146, 0, 0)          # CB_GETCOUNT
    for i in range(min(n, 40)):
        ln = _user.SendMessageW(combo, 0x0149, i, 0)     # CB_GETLBTEXTLEN
        buf = ctypes.create_unicode_buffer(ln + 1)
        _user.SendMessageW(combo, 0x0148, i, ctypes.byref(buf))  # CB_GETLBTEXT
        items.append(buf.value)
    return items


def _select_combo_item(dlg, combo, index):
    """Select item `index` and tell the dialog the selection was committed —
    common file dialogs only apply a new filter/extension on CBN_SELENDOK
    (9), not plain CBN_SELCHANGE (1)."""
    _user.SendMessageW(combo, 0x014E, index, 0)          # CB_SETCURSEL
    cid = _user.GetDlgCtrlID(combo)
    _user.PostMessageW(dlg, 0x0111, (1 << 16) | cid, combo)   # SELCHANGE
    _user.PostMessageW(dlg, 0x0111, (9 << 16) | cid, combo)   # SELENDOK
    time.sleep(0.4)


def nk_export_rtf(hlt_path, launch_timeout=45, save_timeout=30):
    """Open `hlt_path` in NKEditor and save it back as RTF via Save-As
    (file-type combo -> 'RTF(*.RTF)'). Returns the produced .rtf path or None.
    Used so a docx2hlt run can offer a Word-comparable view of what actually
    landed in the .hlt."""
    global LAST_EXPORT_ERROR, LAST_EXPORT_CRASHED
    LAST_EXPORT_ERROR = None
    LAST_EXPORT_CRASHED = False
    hlt_path = os.path.abspath(hlt_path)
    base_name = os.path.splitext(os.path.basename(hlt_path))[0]
    t0_launch = time.time()

    _log('export-rtf: launching NKEditor on %s' % hlt_path)
    proc = _launch(hlt_path)
    pid = proc.pid
    hwnd = None

    def _dead():
        # the frame vanished without our WM_CLOSE, or the process we launched
        # exited non-zero (a single-instance handoff exits 0 — not a crash)
        if hwnd and not _user.IsWindow(hwnd):
            return True
        return proc.poll() not in (None, 0)

    def _fail(msg, crashed=False):
        global LAST_EXPORT_ERROR, LAST_EXPORT_CRASHED
        LAST_EXPORT_CRASHED = crashed
        LAST_EXPORT_ERROR = ('NKEditor 비정상 종료 (%s)' % msg) if crashed \
            else msg
        _log('export-rtf: %s' % LAST_EXPORT_ERROR)
        # an early bail-out must not leave our own editor window behind
        if hwnd and pid == proc.pid and _user.IsWindow(hwnd):
            try:
                _user.PostMessageW(hwnd, 0x0010, 0, 0)   # WM_CLOSE
                _close_with_prompts(pid, hwnd)
            except Exception:
                pass
        return None
    t0 = time.time()
    while time.time() - t0 < launch_timeout and not hwnd:
        _dismiss_nk_dialogs(pid, prefer='no')
        hwnd = _find_nk_window(pid) or _find_nk_window(
            None, os.path.basename(hlt_path))
        if hwnd:
            _dismiss_nk_dialogs(_wpid(hwnd), prefer='no')
        else:
            # single-instance handoff: our file may land on an existing
            # instance — sweep only its crash-recovery prompts
            for h, cn, _t, p in _top_windows():
                if cn == '#32770' and p != pid \
                        and any('복구' in s for s in _dlg_texts(h)):
                    _dismiss_nk_dialogs(p, prefer='no')
        time.sleep(0.4)
    if not hwnd:
        return _fail('NKEditor 창이 뜨지 않음', _dead())
    pid = _wpid(hwnd)
    # wait for the doc to finish loading — the frame title gains the
    # [path] suffix once the file is fully open. A fresh launch can take
    # 25-30 s on a slow machine; SAVE_AS posted before that is swallowed.
    t0 = time.time()
    while time.time() - t0 < launch_timeout:
        tt = ctypes.create_unicode_buffer(512)
        _user.GetWindowTextW(hwnd, tt, 512)
        if '[' in tt.value:
            break
        _dismiss_nk_dialogs(pid, prefer='no')
        time.sleep(0.4)

    dlg = None
    for _attempt in range(6):
        _user.PostMessageW(hwnd, 0x0111, 0xE104, 0)   # WM_COMMAND SAVE_AS
        dlg = _wait_for(lambda: _find_saveas_dialog(pid), 12)
        if dlg:
            break
        time.sleep(1)
    if not dlg:
        return _fail('다른 이름으로 저장 대화상자가 뜨지 않음', _dead())
    time.sleep(0.8)

    # switch the file-type combo to RTF — whichever combo exposes an 'RTF'
    # item. That also rewrites the proposed filename extension.
    picked = None
    for combo in _type_combos(dlg):
        items = _combo_items(combo)
        for i, it in enumerate(items):
            if 'rtf' in it.lower():
                _select_combo_item(dlg, combo, i)
                picked = it
                break
        if picked:
            break
    _log('export-rtf: type=%r' % picked)
    if not picked:
        return _fail('저장 대화상자에 RTF 형식이 없음', _dead())

    # type the full output path into the filename field — otherwise the
    # dialog saves into the instance's last-used folder (e.g. the user's
    # Desktop) and litters it.
    want_rtf = os.path.join(os.path.dirname(hlt_path), base_name + '.rtf')
    for cand in (want_rtf, want_rtf[:-4] + '.RTF'):
        try:
            if os.path.exists(cand):
                os.remove(cand)     # ours — pre-clear so no confirm fires
        except OSError:
            pass
    dlg_dir = _dialog_dir(dlg)
    _log('export-rtf: dir=%r' % dlg_dir)
    # same foreign-file guard as the .hlt save: a same-named file that
    # predates this run in the dialog's folder must never be overwritten
    if dlg_dir:
        f = os.path.join(dlg_dir, base_name + '.rtf')
        foreign = os.path.exists(f) and os.path.getmtime(f) < t0_launch
        if not foreign:
            f = f[:-4] + '.RTF'
            foreign = os.path.exists(f) and os.path.getmtime(f) < t0_launch
        if foreign:
            cb = _find_button_text(dlg, ('취소', 'Cancel'))
            if cb:
                _user.SendMessageW(cb, 0x00F5, 0, 0)
            return _fail('저장 폴더에 이전부터 있던 같은 이름의 파일이 있어 '
                         '덮어쓰지 않음: %s' % f)

    btn = _find_save_button(dlg) or _find_button_text(dlg, ('저장', 'Save'))
    if btn:
        _user.PostMessageW(btn, 0x00F5, 0, 0)    # BM_CLICK — never SendMessage
    else:                                       # (modal deadlock risk)
        _user.PostMessageW(dlg, 0x0111, 1, 0)

    def _want():
        w = [want_rtf, want_rtf[:-4] + '.RTF']
        # sweep likely save dirs by name+recency — the dialog may show a
        # library label ('문서') that _dialog_dir can't resolve
        dirs = {os.path.dirname(hlt_path),
                os.path.expanduser('~\\Documents'),
                os.path.expanduser('~\\Desktop'),
                os.path.expanduser('~\\Downloads')}
        dd = _dialog_dir(_find_saveas_dialog(pid) or dlg) or dlg_dir
        if dd:
            dirs.add(dd)
        for d in dirs:
            w += [os.path.join(d, base_name + ext)
                  for ext in ('.rtf', '.RTF')]
        return w

    saved = None
    t0 = time.time()
    while time.time() - t0 < save_timeout and not saved:
        for cand in _want():
            # only count files created during THIS run — a stale same-named
            # file is foreign, not our output
            if os.path.exists(cand) and os.path.getsize(cand) > 0 \
                    and os.path.getmtime(cand) >= t0_launch - 5:
                saved = cand
                break
        if saved:
            break
        if _dead():
            break
        _dismiss_nk_dialogs(pid, prefer='yes')
        if not _modal_up(pid):
            for d in _find_saveas_dialogs(pid):
                b = (_find_save_button(d)
                     or _find_button_text(d, ('저장', 'Save')))
                if b:
                    _user.PostMessageW(b, 0x00F5, 0, 0)
                else:
                    _user.PostMessageW(d, 0x0111, 1, 0)
        time.sleep(0.8)

    if not saved:
        if _dead():
            _fail('RTF 저장 중', True)
        else:
            _fail('%d초 안에 rtf 파일이 생기지 않음' % save_timeout)
    else:
        _settled(saved)
    if pid == proc.pid and _user.IsWindow(hwnd):
        try:
            _user.PostMessageW(hwnd, 0x0010, 0, 0)   # WM_CLOSE our own frame
            _close_with_prompts(pid, hwnd)
        except Exception:
            pass
    return saved
