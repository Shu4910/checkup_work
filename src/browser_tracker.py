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
    try:
        import uiautomation as auto

        window = None
        if hwnd:
            window = auto.ControlFromHandle(hwnd)

        if window is None or not window.Exists(0, 0):
            window = auto.PaneControl(searchDepth=1, Name=window_title)

        if window and window.Exists(0, 0):
            edit_control = None
            for name in ["アドレスと検索バー", "Address and search bar"]:
                try:
                    ctrl = window.EditControl(searchDepth=10, Name=name)
                    if ctrl.Exists(0, 0):
                        edit_control = ctrl
                        break
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
        print(f"Error in _get_chromium_info: {e}")

    page_title = window_title
    if ' - ' in window_title:
        parts = window_title.rsplit(' - ', 1)
        if len(parts) == 2:
            page_title = parts[0]

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
