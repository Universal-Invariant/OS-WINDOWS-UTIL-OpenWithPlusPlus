#include "stdafx.h"

// Static vector used in dialog proc
static std::vector<PlaceholderInfo> s_placeholders;

extern HMODULE g_hInst;


// Forward declaration
INT_PTR CALLBACK PlaceholderDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);


POINT GetDialogBaseUnitsInPixels(HWND hDlg = NULL) {
    // 1 horizontal base unit = 1/4 of average char width
    // 1 vertical base unit   = 1/8 of average char height
    LONG baseUnits = GetDialogBaseUnits();
    POINT pt;
    pt.x = LOWORD(baseUnits) / 4;   // pixels per horizontal base unit
    pt.y = HIWORD(baseUnits) / 8;   // pixels per vertical base unit
    return pt;
}

// Get system metrics for consistent sizing
int GetSystemMetricDPIAware(HWND hDlg, int nIndex) {
    HDC hdc = GetDC(hDlg);
    int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(hDlg, hdc);

    int base = GetSystemMetrics(nIndex);
    if (dpi != 96) { // scale if not 100% DPI
        base = MulDiv(base, dpi, 96);
    }
    return base;
}

// Or simpler: use dialog base units
POINT GetDialogUnits(HWND hDlg = NULL) {
    LONG base = GetDialogBaseUnits();
    return { LOWORD(base) / 4, HIWORD(base) / 8 }; // pixels per DLU
}

std::optional<std::map<std::wstring, std::wstring>> promptUserForPlaceholders(HWND hParent, const std::vector<PlaceholderInfo>& placeholderInfos)
{
    std::map<std::wstring, std::wstring> results;

    if (placeholderInfos.empty()) { return results; }

    DialogContext ctx{ &placeholderInfos, &results };

    INT_PTR ret = DialogBoxParam(g_hInst, MAKEINTRESOURCE(IDD_PLACEHOLDER_DIALOG), hParent, PlaceholderDialogProc, (LPARAM)&ctx);

    if (ret != IDOK) { results.clear(); return std::nullopt; }

    return results;
}




// Dialog procedure
INT_PTR CALLBACK PlaceholderDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    
    int controlID = LOWORD(wParam);
    int notifyCode = HIWORD(wParam);


    switch (message)
    {
        case WM_INITDIALOG:
        {
            DialogContext* pCtx = reinterpret_cast<DialogContext*>(lParam);
            SetWindowLongPtr(hDlg, DWLP_USER, (LONG_PTR)pCtx);

            const auto& placeholders = *pCtx->placeholderInfos;
            s_placeholders.clear();

            // 🎨 Layout constants
            const int LEFT_MARGIN = 10;
            const int LABEL_WIDTH = 80;
            const int EDIT_X = LEFT_MARGIN + LABEL_WIDTH + 10;
            const int EDIT_WIDTH = 250;
            const int ROW_PADDING = 8;
            const int TOP_PADDING = 12;
            const int BUTTON_PADDING_TOP = 20;
            const int BUTTON_HEIGHT = 24;
            const int BOTTOM_MARGIN = 20;
            const int DIALOG_WIDTH = EDIT_X + EDIT_WIDTH + 20;
            
            int currentY = TOP_PADDING;                        
            for (const auto& ph : placeholders) {
                // Create label
                std::wstring label = ph.name + L":";
                CreateWindowW(L"STATIC", label.c_str(),
                    WS_CHILD | WS_VISIBLE | SS_LEFT,
                    LEFT_MARGIN, currentY, LABEL_WIDTH, 20,
                    hDlg, NULL, g_hInst, NULL);

                // Create edit box
                HWND hEdit = CreateWindowW(L"EDIT", ph.defaultValue.c_str(),
                    WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                    EDIT_X, currentY, EDIT_WIDTH, 22,
                    hDlg, (HMENU)(INT_PTR)s_placeholders.size(), g_hInst, NULL);

                if (hEdit) {
                    s_placeholders.push_back({ ph.name, ph.defaultValue, ph.fullPlaceholder, ph.shouldFocus, hEdit });
                    if (ph.shouldFocus)
                    {                     
                        PostMessageW(hDlg, WM_SET_INITIAL_FOCUS, 0, (LPARAM)hEdit); // necessary because we cannot set focus inside init
                    }
                    PostMessageW(hDlg, WM_SET_TB_END, 0, (LPARAM)hEdit); // necessary because we cannot set focus inside init
                }

                currentY += (std::max)(20, 22) + ROW_PADDING;
            }

            currentY += BUTTON_PADDING_TOP;

            // Create buttons
            int btnWidth = 80, btnSpacing = 10, btnY = currentY;
            CreateWindowW(L"BUTTON", L"OK",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                DIALOG_WIDTH - btnWidth * 2 - btnSpacing - 10, btnY,
                btnWidth, BUTTON_HEIGHT,
                hDlg, (HMENU)IDOK, g_hInst, NULL);

            CreateWindowW(L"BUTTON", L"Cancel",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                DIALOG_WIDTH - btnWidth - 10, btnY,
                btnWidth, BUTTON_HEIGHT,
                hDlg, (HMENU)IDCANCEL, g_hInst, NULL);

            currentY += BUTTON_HEIGHT + BOTTOM_MARGIN;

            // Adjust for window frame
            RECT rcClient = { 0, 0, DIALOG_WIDTH, currentY };
            DWORD style = GetWindowLongW(hDlg, GWL_STYLE);
            DWORD exStyle = GetWindowLongW(hDlg, GWL_EXSTYLE);
            AdjustWindowRectEx(&rcClient, style, FALSE, exStyle);
            SetWindowPos(hDlg, NULL, 0, 0,
                rcClient.right - rcClient.left,
                rcClient.bottom - rcClient.top,
                SWP_NOMOVE | SWP_NOZORDER);

            // Center dialog
            RECT rc;
            GetWindowRect(hDlg, &rc);
            int screenW = GetSystemMetrics(SM_CXSCREEN);
            int screenH = GetSystemMetrics(SM_CYSCREEN);
            SetWindowPos(hDlg, NULL,
                (screenW - (rc.right - rc.left)) / 2,
                (screenH - (rc.bottom - rc.top)) / 2,
                0, 0, SWP_NOSIZE | SWP_NOZORDER);

                                  

        }
        return TRUE;
        case WM_SET_TB_END: // Put cursor at end of edit box
        {
            HWND hTarget = (HWND)lParam;
            if (hTarget && IsWindow(hTarget)) {                
                int index = GetWindowTextLength(hTarget);
                //index = -1;
                SendMessageW(hTarget, EM_SETSEL, index, index);
            }
        }
        return TRUE;
        case WM_SET_INITIAL_FOCUS: // Set initial focus
        {
            HWND hTarget = (HWND)lParam;
            if (hTarget && IsWindow(hTarget)) {
                SetFocus(hTarget);
            }
        }
        return TRUE;
        case WM_COMMAND:
            
            // 👇 Only handle BN_CLICKED for buttons
            if (notifyCode == BN_CLICKED)
            {
                if (controlID == IDOK) {

                    // 👇 Retrieve context
                    DialogContext* pCtx = reinterpret_cast<DialogContext*>(GetWindowLongPtr(hDlg, DWLP_USER));
                    if (pCtx && pCtx->results) {
                        pCtx->results->clear();
                        for (const auto& pi : s_placeholders) {
                            wchar_t buffer[1024] = { 0 };
                            GetWindowTextW(pi.hEdit, buffer, 1024);
                            (*pCtx->results)[pi.name] = buffer;
                        }
                    }
                    EndDialog(hDlg, IDOK);
                    return TRUE;
                }
                else if (controlID == IDCANCEL) {
                    EndDialog(hDlg, IDCANCEL);
                    return TRUE;
                }
            }
            break;

        case WM_CLOSE:
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
    }

    return FALSE;
}