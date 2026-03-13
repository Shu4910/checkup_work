"""
ブラウザURL・タイトル取得モジュール
Chrome、Edge、Firefoxに対応
"""
import win32gui
import win32process
import psutil


def get_browser_info(process_name, window_title, hwnd=0):
    """
    ブラウザの場合、URLとページタイトルを取得

    Args:
        process_name (str): プロセス名
        window_title (str): ウィンドウタイトル
        hwnd (int): ウィンドウハンドル

    Returns:
        dict: {
            'url': str,
            'page_title': str
        }
    """
    browser_list = ['chrome.exe', 'msedge.exe', 'firefox.exe']

    if process_name.lower() not in browser_list:
        return {'url': '', 'page_title': ''}

    try:
        if process_name.lower() in ['chrome.exe', 'msedge.exe']:
            return _get_chromium_info(window_title, hwnd)
        elif process_name.lower() == 'firefox.exe':
            return _get_firefox_info(window_title, hwnd)
    except Exception as e:
        print(f"Error getting browser info: {e}")

    return {'url': '', 'page_title': ''}


def _get_chromium_info(window_title, hwnd=0):
    """
    Chrome/Edgeの情報を取得
    UIAutomationを使用してアドレスバーからURLを取得

    Args:
        window_title (str): ウィンドウタイトル
        hwnd (int): ウィンドウハンドル

    Returns:
        dict: {'url': str, 'page_title': str}
    """
    page_title = window_title
    if ' - ' in window_title:
        parts = window_title.rsplit(' - ', 1)
        if len(parts) == 2:
            page_title = parts[0]

    try:
        import uiautomation as auto

        window = None
        if hwnd:
            window = auto.ControlFromHandle(hwnd)

        if window is None or not window.Exists(0, 0):
            window = auto.PaneControl(ClassName="Chrome_WidgetWin_1", searchDepth=2)

        if window is None or not window.Exists(0, 0):
            return {'url': '', 'page_title': page_title}

        url = ''
        edit_control = None

        # 戦略1: ToolBar 内の EditControl（最も確実）
        try:
            toolbar = window.ToolBarControl(searchDepth=8)
            if toolbar.Exists(0, 0):
                ctrl = toolbar.EditControl(searchDepth=3)
                if ctrl.Exists(0, 0):
                    edit_control = ctrl
        except Exception:
            pass

        # 戦略2: Name で検索（日本語・英語）
        if edit_control is None:
            for name in ["アドレスと検索バー", "Address and search bar",
                         "Search Google or type a URL"]:
                try:
                    ctrl = window.EditControl(searchDepth=12, Name=name)
                    if ctrl.Exists(0, 0):
                        edit_control = ctrl
                        break
                except Exception:
                    pass

        # 戦略3: AutomationId で検索
        if edit_control is None:
            try:
                ctrl = window.EditControl(searchDepth=12, AutomationId="OmniboxViewViews")
                if ctrl.Exists(0, 0):
                    edit_control = ctrl
            except Exception:
                pass

        if edit_control:
            try:
                url = edit_control.GetValuePattern().Value
            except Exception:
                pass
            if not url:
                try:
                    url = edit_control.GetLegacyIAccessiblePattern().Value
                except Exception:
                    pass

        return {'url': url if url else '', 'page_title': page_title}

    except Exception as e:
        print(f"Error in _get_chromium_info: {e}")

    return {'url': '', 'page_title': page_title}


def _get_firefox_info(window_title, hwnd=0):
    """
    Firefoxの情報を取得
    UIAutomationを使用してアドレスバーからURLを取得

    Args:
        window_title (str): ウィンドウタイトル
        hwnd (int): ウィンドウハンドル

    Returns:
        dict: {'url': str, 'page_title': str}
    """
    try:
        import uiautomation as auto

        window = None
        if hwnd:
            window = auto.ControlFromHandle(hwnd)

        if window is None or not window.Exists(0, 0):
            window = auto.PaneControl(searchDepth=1, Name=window_title)

        if window and window.Exists(0, 0):
            edit_control = None
            try:
                ctrl = window.EditControl(searchDepth=10, AutomationId="urlbar-input")
                if ctrl.Exists(0, 0):
                    edit_control = ctrl
            except:
                pass
            if edit_control is None:
                try:
                    ctrl = window.ComboBoxControl(searchDepth=10, Name="検索または URL を入力してください")
                    if ctrl.Exists(0, 0):
                        edit_control = ctrl
                except:
                    pass

            url = ''
            if edit_control:
                try:
                    url = edit_control.GetValuePattern().Value
                except:
                    pass
                if not url:
                    try:
                        url = edit_control.GetLegacyIAccessiblePattern().Value
                    except:
                        pass

            page_title = window_title
            if ' - ' in window_title:
                parts = window_title.rsplit(' - ', 1)
                if len(parts) == 2:
                    page_title = parts[0]

            return {
                'url': url if url else '',
                'page_title': page_title if page_title else ''
            }
    except Exception as e:
        print(f"Error in _get_firefox_info: {e}")

    page_title = window_title
    if ' - ' in window_title:
        parts = window_title.rsplit(' - ', 1)
        if len(parts) == 2:
            page_title = parts[0]

    return {'url': '', 'page_title': page_title}
