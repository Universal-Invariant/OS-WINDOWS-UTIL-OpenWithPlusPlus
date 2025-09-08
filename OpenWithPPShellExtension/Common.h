
#pragma once
#include <string>
#include <vector>
#include <map>
#include <windows.h>

#define PRODUCT_NAME L"OpenWithPP"

#define WM_SET_INITIAL_FOCUS            (WM_USER + 123)
#define WM_SET_TB_END                   (WM_USER + 124)





struct PlaceholderInfo {
    std::wstring name;
    std::wstring defaultValue;
    std::wstring fullPlaceholder;
    bool shouldFocus;
    HWND hEdit; // Only used in dialog proc
};

struct DialogContext {
    const std::vector<PlaceholderInfo>* placeholderInfos; 
    std::map<std::wstring, std::wstring>* results;
};

std::optional<std::map<std::wstring, std::wstring>> promptUserForPlaceholders(HWND hParent, const std::vector<PlaceholderInfo>& placeholderInfos);



