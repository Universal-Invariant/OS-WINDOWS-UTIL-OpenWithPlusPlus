
#include "stdafx.h"
#include "Main.h"

// TODO: better deal with multiple file selection/application and quotations

std::wstring trim(const std::wstring& s) {
    if (s.empty()) return s;

    size_t start = 0;
    size_t end = s.length();

    while (start < end && iswspace(s[start])) ++start;
    while (end > start && iswspace(s[end - 1])) --end;

    return s.substr(start, end - start);
}

Item::~Item()
{
	if (Icon)
	{
		DeleteObject(Icon);
		Icon = NULL;
	}
}


std::wstring JoinList(std::list<std::wstring>* list, const std::wstring& sep)
{
	std::wstring ret;

	if ((*list).size() > 0)
		ret = *(*list).begin();

	if ((*list).size() > 1)
	{
		std::list<std::wstring>::iterator it = (*list).begin();
		it++;

		for (it; it != (*list).end(); it++)
			ret += sep + (*it);
	}

	return ret;
}


std::wstring ToLower(std::wstring val)
{
	std::transform(val.begin(), val.end(), val.begin(), tolower);
	return val;
}


std::wstring GetExtNoDot(std::wstring pathName)
{
	size_t period = pathName.find_last_of(L".");
	return ToLower(pathName.substr(period + 1));
}


BOOL FileExists(std::wstring file)
{
	if (file.length() == 0)
		return false;

	DWORD dwAttrib = GetFileAttributes(file.c_str());
	return dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY);
}


BOOL DirectoryExist(std::wstring path)
{
	DWORD dwAttrib = GetFileAttributes(path.c_str());
	return (dwAttrib != INVALID_FILE_ATTRIBUTES && (dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}


HBITMAP Create32BitHBITMAP(UINT cx, UINT cy, PBYTE* ppbBits)
{
	BITMAPINFO bmi;
	ZeroMemory(&bmi, sizeof(bmi));
	bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
	bmi.bmiHeader.biWidth = cx;
	bmi.bmiHeader.biHeight = -(LONG)cy;
	bmi.bmiHeader.biPlanes = 1;
	bmi.bmiHeader.biBitCount = 32;
	bmi.bmiHeader.biCompression = BI_RGB;
	HDC hDC = GetDC(NULL);
	HBITMAP hbmp = CreateDIBSection(hDC, &bmi, DIB_RGB_COLORS, (void**)ppbBits, NULL, 0);
	ReleaseDC(NULL, hDC);
	return(hbmp);
}


HBITMAP ConvertIconToBitmap(HICON hicon)
{
    IWICImagingFactory* pFactory;
    IWICBitmap* pBitmap;
    IWICFormatConverter* pConverter;
    HBITMAP ret = NULL;

    if (SUCCEEDED(CoCreateInstance(CLSID_WICImagingFactory, NULL,
        CLSCTX_INPROC_SERVER, IID_IWICImagingFactory, (void**)&pFactory)))
    {
        if (SUCCEEDED(pFactory->CreateBitmapFromHICON(hicon, &pBitmap)))
        {
            if (SUCCEEDED(pFactory->CreateFormatConverter(&pConverter)))
            {
                UINT cx, cy;
                PBYTE pbBits;
                HBITMAP hbmp;

                if (SUCCEEDED(pConverter->Initialize(pBitmap, GUID_WICPixelFormat32bppPBGRA,
                    WICBitmapDitherTypeNone, NULL, 0.0f, WICBitmapPaletteTypeCustom)) && SUCCEEDED(
                        pConverter->GetSize(&cx, &cy)) && NULL != (hbmp = Create32BitHBITMAP(cx, cy, &pbBits)))
                {
                    UINT cbStride = cx * sizeof(UINT32);
                    UINT cbBitmap = cy * cbStride;

                    if (SUCCEEDED(pConverter->CopyPixels(NULL, cbStride, cbBitmap, pbBits)))
                        ret = hbmp;
                    else
                        DeleteObject(hbmp);
                }

                pConverter->Release();
            }

            pBitmap->Release();
        }

        pFactory->Release();
    }

    return ret;
}


HRESULT SetIcon(HMENU menu, UINT position, UINT flags, Item* item)
{
	std::wstring path(item->IconFile);

	if (path.find(L"%") != std::string::npos)
	{
		WCHAR szPath[500];
		DWORD result = ExpandEnvironmentStrings(path.c_str(), szPath, std::size(szPath));

		if (result)
			path = szPath;
	}

	// starts with
	if (path.rfind(L"..\\", 0) == 0)
	{
		WCHAR szExeDir[500];
		SHRegGetPath(HKEY_CURRENT_USER, L"Software\\" PRODUCT_NAME, L"ExeDir", szExeDir, NULL);
		std::wstring exeDir(szExeDir);
		path = exeDir + path;
	}

	if (!FileExists(path))
		return S_OK;

	if (!item->Icon)
	{
		HICON iconSmall;
		int res = ExtractIconEx(path.c_str(), item->IconIndex, NULL, &iconSmall, 1);

		if (!res)
			return E_FAIL;

		HBITMAP bmp = ConvertIconToBitmap(iconSmall);
		
		if (iconSmall)
			DestroyIcon(iconSmall);

		if (!bmp)
			return E_FAIL;

		item->Icon = bmp;
	}

	int res = SetMenuItemBitmaps(menu, position, flags, item->Icon, item->Icon);

	if (!res)
		return E_FAIL;

	return S_OK;
}


HRESULT CMain::LoadXML()
{
	for (Item* item : g_Items)
		delete item;

	g_Items.clear();

	WCHAR path[500];
	SHRegGetPath(HKEY_CURRENT_USER, L"Software\\" PRODUCT_NAME, L"SettingsLocation", path, NULL);

	if (!FileExists(path))
		return E_FAIL;

	ATL::CComPtr<IXMLDOMDocument> doc;
	HRESULT hr = doc.CoCreateInstance(__uuidof(DOMDocument));

	if (FAILED(hr))
		return hr;

	VARIANT_BOOL success;
	hr = doc->load(ATL::CComVariant(path), &success);

	if (FAILED(hr))
		return hr;
	
	ATL::CComPtr<IXMLDOMNodeList> items;
	hr = doc->selectNodes(ATL::CComBSTR(L"AppSettings/Items/Item"), &items);

	if (FAILED(hr))
		return hr;

	IXMLDOMNode* itemNode;

	while (S_OK == items->nextNode(&itemNode))
	{
		Item* item = new Item();
		
		ATL::CComPtr<IXMLDOMNodeList> itemFields;
		hr = itemNode->get_childNodes(&itemFields);

		if FAILED(hr)
			hr;

		IXMLDOMNode* itemField;

		while (S_OK == itemFields->nextNode(&itemField))
		{
			BSTR nodeName;
			hr = itemField->get_nodeName(&nodeName);

			if FAILED(hr)
				return hr;

			ATL::CComBSTR cNodeName(nodeName);
			
			BSTR nodeText;
			hr = itemField->get_text(&nodeText);

			if FAILED(hr)
				return hr;

			ATL::CComBSTR cNodeText(nodeText);

			if (cNodeName == L"Name")
				item->Name = cNodeText;
			else if (cNodeName == L"Path")
				item->Path = cNodeText;
			else if (cNodeName == L"Arguments")
				item->Arguments = cNodeText;
			else if (cNodeName == L"WorkingDirectory")
				item->WorkingDirectory = cNodeText;
			else if (cNodeName == L"IconFile")
				item->IconFile = cNodeText;
			else if (cNodeName == L"IconIndex")
				item->IconIndex = std::stoi(std::wstring(cNodeText));
			else if (cNodeName == L"FileTypes")
				item->FileTypes = cNodeText;
			else if (cNodeName == L"Filter")
				item->Filter = cNodeText;
			else if (cNodeName == L"SubMenu")
				item->SubMenu = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"Directories")
				item->Directories = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"RunAsAdmin")
				item->RunAsAdmin = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"HideWindow")
				item->HideWindow = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"UseVariableQuotes")
				item->UseVariableQuotes = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"UseVariables")
				item->UseVariables = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"Hidden")
				item->Hidden = (cNodeText == L"true") ? true : false;
			else if (cNodeName == L"Sort")
				item->Sort = (cNodeText == L"true") ? true : false;
		}

		g_Items.push_back(item);
	}

	return S_OK;
}


STDMETHODIMP CMain::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hProgID)
{
	FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stg;
	g_ShellItems.clear();

	if (pDataObj)
	{
		if (FAILED(pDataObj->GetData(&fmt, &stg)))
			return E_INVALIDARG;

		int uNumFiles = DragQueryFile((HDROP)stg.hGlobal, 0xFFFFFFFF, NULL, 0);

		if (!uNumFiles)
		{
			ReleaseStgMedium(&stg);
			return E_INVALIDARG;
		}

		for (int i = 0; i < uNumFiles; i++)
		{
			WCHAR buffer[500];
			DragQueryFile((HDROP)stg.hGlobal, i, buffer, std::size(buffer));
			g_ShellItems.push_back(buffer);
		}

		ReleaseStgMedium(&stg);
	}
	else if (pidlFolder)
	{
		WCHAR buffer[500];

		if (!SHGetPathFromIDList(pidlFolder, buffer))
			return E_FAIL;

		g_ShellItems.push_back(buffer);
	}
	else
		return E_FAIL;

	return S_OK;
}


STDMETHODIMP CMain::QueryContextMenu(HMENU hmenu, UINT uMenuIndex, UINT uidFirstCmd, UINT uidLastCmd, UINT uFlags)
{
	if (!g_Items.size() || SHRegGetBoolUSValue(
		L"Software\\" PRODUCT_NAME, L"Reload", FALSE, TRUE))
	{
		REGISTRY_ENTRY re = GetRegEntry(HKEY_CURRENT_USER,
			L"Software\\" PRODUCT_NAME, L"Reload", NULL, REG_DWORD, 0);

		if (FAILED(CreateRegKeyAndSetValue(&re)))
			return E_FAIL;

		HRESULT hr = LoadXML();

		if (FAILED(hr))
			return hr;

		WCHAR szExeDir[500];
		SHRegGetPath(HKEY_CURRENT_USER, L"Software\\" PRODUCT_NAME, L"ExeDir", szExeDir, NULL);
		std::wstring exeDir(szExeDir);
		std::wstring iconPath = exeDir + L"Icons\\Main.ico";

		if (FileExists(iconPath))
			IconItem.IconFile = iconPath;
	}

	UINT command = uidFirstCmd;
	HMENU subMenu = CreatePopupMenu();

	bool addSubSep = false;
	bool isCtrlPressed = GetKeyState(VK_CONTROL) < 0;
	bool isFile = FileExists(*g_ShellItems.begin());
	bool isDirectory = !isFile && DirectoryExist(*g_ShellItems.begin());

	std::wstring ext = GetExtNoDot(*g_ShellItems.begin());
	std::wstring path = *g_ShellItems.begin();

	g_EditCommandIndex = -1;

	int res = InsertMenu(hmenu, uMenuIndex, MF_BYPOSITION | MF_POPUP, (UINT_PTR)subMenu, L"Open with++");
	int subMenuIndex = uMenuIndex;

	if (!res)
		return E_FAIL;

	if (FileExists(IconItem.IconFile))
		SetIcon(hmenu, uMenuIndex, MF_BYPOSITION, &IconItem);

	uMenuIndex += 1;

	for (UINT i = 0; i < g_Items.size(); i++)
	{
		g_Items[i]->CommandIndex = -1;

		if (isFile && g_Items[i]->FileTypes != L"" && ext != L""
			&& (L" " + g_Items[i]->FileTypes + L" ").find(L" " + ext + L" ") != std::wstring::npos
			&& (g_Items[i]->Filter == L"" || std::regex_search(path, std::wregex(g_Items[i]->Filter)))
			&& (!g_Items[i]->Hidden || (g_Items[i]->Hidden && isCtrlPressed)))
		{
			g_Items[i]->CommandIndex = command - uidFirstCmd;

			if (g_Items[i]->SubMenu)
			{
				res = InsertMenu(subMenu, -1, MF_BYPOSITION, command, g_Items[i]->Name.c_str());

				if (!res)
					return E_FAIL;

				SetIcon(hmenu, command, MF_BYCOMMAND, g_Items[i]);
				addSubSep = true;
			}
			else
			{
				res = InsertMenu(hmenu, uMenuIndex, MF_BYPOSITION, command, g_Items[i]->Name.c_str());

				if (!res)
					return E_FAIL;

				SetIcon(hmenu, uMenuIndex, MF_BYPOSITION, g_Items[i]);
				uMenuIndex += 1;
			}

			command += 1;
		}           
	}

	if (addSubSep)
	{
		res = InsertMenu(subMenu, -1, MF_BYPOSITION | MF_SEPARATOR, command, NULL);

		if (!res)
			return E_FAIL;

		command += 1;
	}

	bool addSubSep2 = false;

	for (UINT i = 0; i < g_Items.size(); i++)
	{
		if ((g_Items[i]->FileTypes == L"*.*" && isFile || g_Items[i]->Directories && isDirectory)
			&& (g_Items[i]->Filter == L"" || std::regex_search(path, std::wregex(g_Items[i]->Filter)))
			&& (!g_Items[i]->Hidden || (g_Items[i]->Hidden && isCtrlPressed)))
		{
			g_Items[i]->CommandIndex = command - uidFirstCmd;

			if (g_Items[i]->SubMenu)
			{
				res = InsertMenu(subMenu, -1, MF_BYPOSITION, command, g_Items[i]->Name.c_str());

				if (!res)
					return E_FAIL;

				SetIcon(hmenu, command, MF_BYCOMMAND, g_Items[i]);
				addSubSep2 = true;
			}
			else
			{
				res = InsertMenu(hmenu, uMenuIndex, MF_BYPOSITION, command, g_Items[i]->Name.c_str());

				if (!res)
					return E_FAIL;

				SetIcon(hmenu, uMenuIndex, MF_BYPOSITION, g_Items[i]);
				uMenuIndex += 1;
			}

			command += 1;
		}
	}

	if (addSubSep2)
	{
		res = InsertMenu(subMenu, -1, MF_BYPOSITION | MF_SEPARATOR, command, NULL);

		if (!res)
			return E_FAIL;

		command += 1;
	}

	if (addSubSep || addSubSep2)
	{
		res = InsertMenu(subMenu, -1, MF_BYPOSITION, command, L"Customize Open with++");

		if (!res)
			return E_FAIL;

		if (FileExists(IconItem.IconFile))
			SetIcon(subMenu, command, MF_BYCOMMAND, &IconItem);

		g_EditCommandIndex = command - uidFirstCmd;
	}
	else
		DeleteMenu(hmenu, subMenuIndex, MF_BYPOSITION);

	return MAKE_HRESULT(SEVERITY_SUCCESS, 0, command - uidFirstCmd + 1);
}

std::wstring GetFileCreation(std::wstring pFileName)
{
	std::wstringstream ss;
	WIN32_FILE_ATTRIBUTE_DATA wfad;
	SYSTEMTIME st;
	GetFileAttributesExW(pFileName.c_str(), GetFileExInfoStandard, &wfad);
	FileTimeToSystemTime(&wfad.ftCreationTime, &st);
	ss << st.wYear << '/' << st.wMonth << '/' << st.wDay << " " << (st.wHour % 24) << ":" << (st.wMinute % 60) << ":" << (st.wSecond % 60);
	return ss.str();
}

std::wstring GetLastFileAccess(std::wstring pFileName)
{
	std::wstringstream ss;
	WIN32_FILE_ATTRIBUTE_DATA wfad;
	SYSTEMTIME st;
	GetFileAttributesExW(pFileName.c_str(), GetFileExInfoStandard, &wfad);
	FileTimeToSystemTime(&wfad.ftLastAccessTime, &st);
	ss << st.wYear << '/' << st.wMonth << '/' << st.wDay << " " << (st.wHour % 24) << ":" << (st.wMinute % 60) << ":" << (st.wSecond % 60);
	return ss.str();
}

std::wstring GetLastFileWrite(std::wstring pFileName)
{
	std::wstringstream ss;
	WIN32_FILE_ATTRIBUTE_DATA wfad;
	SYSTEMTIME st;
	GetFileAttributesExW(pFileName.c_str(), GetFileExInfoStandard, &wfad);
	FileTimeToSystemTime(&wfad.ftLastWriteTime, &st);
	ss << st.wYear << '/' << st.wMonth << '/' << st.wDay << " " << (st.wHour % 24) << ":" << (st.wMinute % 60) << ":" << (st.wSecond % 60);
	return ss.str();
}

std::wstring GetFileAttributes(std::wstring pFileName)
{
	WIN32_FILE_ATTRIBUTE_DATA wfad;
	SYSTEMTIME st;
	GetFileAttributesExW(pFileName.c_str(), GetFileExInfoStandard, &wfad);
	return std::to_wstring(wfad.dwFileAttributes);
}


STDAPI_(std::wstring) FormatCommand(std::wstring args, std::list<std::wstring> g_ShellItems, bool useVariableQuotes, int index = 0)
{

	// TODO: allow configurable token identifing chars
	auto wrappingTok = L"%";
	
	// Handle empty item list case
	if (g_ShellItems.empty()) return args;
	// Index should be within range
	if (index < 0 || index >= g_ShellItems.size()) return args; 

	// Build current file path object
	auto it = g_ShellItems.begin();
	std::advance(it, index);
	std::filesystem::path currentFile(*it);	
	
	auto Q = (useVariableQuotes) ? L"\"" : L"";

	// Build all required lists
	std::vector<std::wstring> allFiles;
	std::vector<std::wstring> allPaths;
	std::vector<std::wstring> allFilenames;
	std::vector<std::wstring> allFilesNoExt;
	std::vector<std::wstring> allFilenamesNoExt;
	std::vector<std::wstring> allExts;
	std::vector<std::wstring> allDotExts;
	std::vector<std::wstring> allFileSizes;
	std::vector<std::wstring> allCreatedDates;
	std::vector<std::wstring> allLastAccess;
	std::vector<std::wstring> allLastWrite;
	std::vector<std::wstring> allAttributes;
	std::vector<std::wstring> allExists;
	
	auto removeExt = [](std::filesystem::path currentFile) -> std::wstring
		{
			return trim(currentFile.wstring()).replace(currentFile.wstring().length() - currentFile.extension().wstring().length(), currentFile.extension().wstring().length(), L"");
		};

	auto removeExtDot = [](std::filesystem::path currentFile) -> std::wstring
		{
			auto e = trim(currentFile.extension().wstring());
			if (e.length() > 0 && e[0] == L'.')
				return e.erase(0, 1);
			return e;
		};


	for (const auto& item2 : g_ShellItems) {
		auto item = trim(item2);
		if (item.length() == 0) continue;

		std::filesystem::path fp(item);
		allFiles.push_back(item);
		allPaths.push_back(trim(fp.parent_path().wstring()));
		allFilenames.push_back(trim(fp.filename().wstring()));
		allFilesNoExt.push_back(removeExt(fp));
		allFilenamesNoExt.push_back(trim(fp.filename().stem().wstring()));
		allExts.push_back(removeExtDot(fp));
		allDotExts.push_back(trim(fp.extension().wstring()));
		if (std::filesystem::exists(fp))
		{
			allExists.push_back(L"1");
			allFileSizes.push_back(trim(std::to_wstring(std::filesystem::file_size(fp))));
			allCreatedDates.push_back(GetFileCreation(item));
			allLastAccess.push_back(GetLastFileAccess(item));
			allLastWrite.push_back(GetLastFileWrite(item));
			allAttributes.push_back(GetFileAttributes(item));
		}
		else {
			allExists.push_back(L"0");
			allFileSizes.push_back(L"0");
			allCreatedDates.push_back(L" / / ");
			allLastAccess.push_back(L" / / ");
			allLastWrite.push_back(L" / / ");
			allAttributes.push_back(L"-1");
		}
		
	}

	// Helper function to join lists with semicolon and optional quotes
	auto joinList = [](const std::vector<std::wstring>& items, bool useQuotes = false, std::wstring delim = L";") -> std::wstring {
		if (items.empty()) return L"";
		std::wstringstream ss;
		
		for (int i = 0; i < items.size(); i++) {
			if (useQuotes) { 
				ss << L"\"" << trim(items[i]) << L"\"";
			}
			else {
				ss << trim(items[i]);
			}

			if (i < items.size() - 1) ss << delim;
		}
		
		return trim(ss.str());
	};

	// Helper function for safe token replacement
	// TODO: Make more robust for arbitrary token wrapping identifier
	auto replaceToken = [wrappingTok](std::wstring& str, const std::wstring& token, const std::wstring& value) {		
		size_t pos = str.find(wrappingTok + token + wrappingTok);
		if (pos != std::wstring::npos) {
			str.replace(pos, token.length() + 2, value);
		}
	};




	replaceToken(args, L"numItems", std::to_wstring(allFiles.size()));

	// Perform replacements for all tokens. We use a single loop to deal with the 4 cases _  q_, _[], and q_[].
	for (int i = 0; i < 4; i++)
	{		
		std::wstring QQ = Q;
		if (i % 2 == 1) QQ = L"\""; // forces quotes for q tokens
		#define replaceToken2(a, b)	replaceToken(args, std::wstring((i % 2 == 0) ? L"" : L"q") + a + ((i < 2) ? L"" : L"[]"), QQ + ((i < 2) ? b[index] : joinList(b)) + QQ)

		
		replaceToken2(L"file", allFiles);
		replaceToken2(L"path", allPaths);
		replaceToken2(L"filename", allFilenames);
		replaceToken2(L"file-no-ext", allFilesNoExt);
		replaceToken2(L"filename-no-ext", allFilenamesNoExt);
		replaceToken2(L"ext", allExts);
		replaceToken2(L".ext", allDotExts);
		replaceToken2(L"exists", allExists);
		replaceToken2(L"filesize", allFileSizes);
		replaceToken2(L"created", allCreatedDates);
		replaceToken2(L"accessed", allLastAccess);
		replaceToken2(L"written", allLastWrite);
		replaceToken2(L"attributes", allAttributes);
	}
	

	return args;
}



extern "C" __declspec(dllexport) wchar_t* __stdcall FormatCommandVB(BSTR args, BSTR shellItemsString, bool useVariableQuotes, int index = 0)
{
	// Convert BSTR to std::wstring
	std::wstring argsW(args);
	std::wstring shellItemsW(shellItemsString);
	
	// Split shell items string into list
	std::list<std::wstring> g_ShellItems;
	size_t pos = 0;
	while ((pos = shellItemsW.find(L';', pos)) != std::wstring::npos) {
		g_ShellItems.push_back(shellItemsW.substr(0, pos));
		shellItemsW = shellItemsW.substr(pos + 1);
		pos = 0;
	}
	if (!shellItemsW.empty()) {
		g_ShellItems.push_back(shellItemsW);
	}

	std::wstring value = FormatCommand(argsW, g_ShellItems, useVariableQuotes, index);
	// Allocate and return string safely
	size_t bufferSize = (value.size() + 1) * sizeof(wchar_t);
	wchar_t* buffer = (wchar_t*)CoTaskMemAlloc(bufferSize);
	if (!buffer) return nullptr;
	wcscpy_s(buffer, value.size() + 1, value.c_str());
	return buffer;
}

std::vector<PlaceholderInfo> extractPlaceholders(const std::wstring& input) {
	std::vector<PlaceholderInfo> placeholders;
	
	// TODO: allow configurable prefixes and suffixes
	const std::wstring prefix = L"%$";
	const std::wstring suffix = L"$%";

	size_t pos = 0;

	while (pos < input.length()) {
		size_t start = input.find(prefix, pos);
		if (start == std::wstring::npos) break;
		start += prefix.length();

		size_t end = input.find(suffix, start);
		if (end == std::wstring::npos) break;

		std::wstring fullContent = input.substr(start, end - start);

		PlaceholderInfo info;
		info.fullPlaceholder = prefix + fullContent + suffix;
		info.shouldFocus = false; // default

		size_t eqPos = fullContent.find(L'=');

		std::wstring namePart;
		if (eqPos != std::wstring::npos) {
			namePart = fullContent.substr(0, eqPos);
			info.defaultValue = trim(fullContent.substr(eqPos + 1));
		}
		else {
			namePart = fullContent;
			info.defaultValue.clear();
		}

		// 👇 Check for * prefix
		if (!namePart.empty() && namePart[0] == L'*') {
			info.shouldFocus = true;
			namePart = namePart.substr(1); // remove *			
		}

		info.name = trim(namePart);

		placeholders.push_back(info);
		pos = end + suffix.length();
	}

	return placeholders;
}

std::wstring replacePlaceholders(const std::wstring& input, const std::map<std::wstring, std::wstring>& replacements)
{
	std::wstring result = input;

	// Get list of placeholders (we need fullPlaceholder and name)
	auto placeholders = extractPlaceholders(input); 

	// Sort by position or reverse order? We'll do reverse by start position
	std::vector<std::pair<size_t, PlaceholderInfo>> found;

	for (const auto& ph : placeholders) {
		size_t pos = 0;
		while ((pos = result.find(ph.fullPlaceholder, pos)) != std::wstring::npos) {
			found.emplace_back(pos, ph);
			pos += ph.fullPlaceholder.length();
		}
	}

	// Sort descending by position
	std::sort(found.begin(), found.end(), [](const auto& a, const auto& b) {
		return a.first > b.first;
		});

	for (const auto& [pos, ph] : found) {
		auto it = replacements.find(ph.name);
		if (it != replacements.end()) {
			result.replace(pos, ph.fullPlaceholder.length(), it->second);
		}
	}

	return result;
}


// Used to create the command to display in VB but uses dummy variables
extern "C" __declspec(dllexport) const wchar_t* FormatCommandVBD(const wchar_t* args, const wchar_t* files, bool useVariableQuotes, int index = 0) 
{
	// Creating an input string stream from the input string
	std::wstringstream ss(files);

	// Temporary string to store each token
	std::wstring token;
	std::list<std::wstring> itemList;

	while (std::getline(ss, token, L';'))
		itemList.push_back(token);


	

	auto wargs = std::wstring(args);
	// wargs = L"C:\\test %$name$% %$file$% -r -d"; // for debugging

	// Call the original function
	std::wstring result = FormatCommand(wargs, itemList, useVariableQuotes, index);

	

	// Fill in the user variables with default names(since it is only for displaying example string)
	auto placeholderNames = extractPlaceholders(result); // from Part 1 above
	std::map<std::wstring, std::wstring> replacements;
	size_t pos = 0;
	while (pos < placeholderNames.size())
	{
		auto s = std::wstring(L"<Var");
		//replacements[placeholderNames[pos].name] = s.append(L"(").append(std::to_wstring(pos + 1)).append(L") `").append(placeholderNames[pos].fullPlaceholder).append(L"`>");
		replacements[placeholderNames[pos].name] = L"<Var " + std::to_wstring(pos + 1) + L">";
		if (placeholderNames[pos].defaultValue != L"")
			replacements[placeholderNames[pos].name] = placeholderNames[pos].defaultValue;

		pos++;
	}
	result = replacePlaceholders(result, replacements);



	// Allocate memory for the result and return it
	size_t bufferSize = (result.size() + 1) * sizeof(wchar_t);
	wchar_t* buffer = (wchar_t*)CoTaskMemAlloc(bufferSize);
	wcscpy_s(buffer, result.size() + 1, result.c_str());
	return buffer;
}



std::optional<std::wstring> fixupUserVariables(std::wstring cmdline)
{
	std::wstring result = cmdline;

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
	icex.dwICC = ICC_STANDARD_CLASSES;
	InitCommonControlsEx(&icex);

	// 👇 Extract full placeholder info (names + defaults)
	auto placeholders = extractPlaceholders(cmdline);

	if (placeholders.empty()) {
		//MessageBoxW(NULL, result.c_str(), L"No user variables to replace!", MB_OK | MB_ICONINFORMATION);
		return result;
	}
	// Query user for parameters	
	auto replacements = promptUserForPlaceholders(NULL, placeholders);
	if (!replacements.has_value()) return std::nullopt;

	if (!replacements.value().empty()) {
		result = replacePlaceholders(cmdline, replacements.value());
		//MessageBoxW(NULL, result.c_str(), L"Result", MB_OK | MB_ICONINFORMATION);
	}
	else {
		//MessageBoxW(NULL, result.c_str(), L"No user variables replaced!", MB_OK | MB_ICONINFORMATION);
	}

	return result;
}

// VB Wrapper to use fixupUserVariables in VB
extern "C" __declspec(dllexport) wchar_t* __stdcall FixupUserVariablesVB(BSTR cmdline)
{
	// Convert BSTR to std::wstring
	std::wstring input(cmdline);

	// Call your existing function
	auto result = fixupUserVariables(input);

	if (!result.has_value()) {
		// User canceled - return nullptr for VB.NET to interpret as Nothing
		return nullptr;
	}

	// Allocate and return string safely
	size_t bufferSize = (result.value().size() + 1) * sizeof(wchar_t);
	wchar_t* buffer = (wchar_t*)CoTaskMemAlloc(bufferSize);
	if (!buffer) return nullptr;
	wcscpy_s(buffer, result.value().size() + 1, result.value().c_str());
	return buffer;
}



STDMETHODIMP CMain::InvokeCommand(LPCMINVOKECOMMANDINFO pCmdInfo)
{
	if (HIWORD(pCmdInfo->lpVerb) != 0)
		return E_FAIL;

	HWND hwnd = GetActiveWindow();
	WORD id = LOWORD(pCmdInfo->lpVerb);

	for (UINT i = 0; i < g_Items.size(); i++)
	{
		if (id == g_Items[i]->CommandIndex)
		{
			PROCESS_INFORMATION pi = { 0 };
			STARTUPINFO si = { sizeof(si) };



			// transform user variables(%$<..>$%) and replace with user queries if they exist and if useVariables is set
			auto sargs = g_Items[i]->Arguments;
			if (g_Items[i]->UseVariables)
			{
				auto ssargs = fixupUserVariables(g_Items[i]->Arguments);

				if (!ssargs.has_value())
				{
					// DEBUG:
					//MessageBox(NULL, L"Dialog Cancelled", L"OpenWithPP cmdLine", MB_OK);
					return E_FAIL;
				}
				sargs = ssargs.value();
			}

			for (int j = 0; j < g_ShellItems.size(); j++)
			{
				std::wstring args = FormatCommand(sargs, g_ShellItems, g_Items[j]->UseVariableQuotes, j);

				// DEBUG: Show user the command line being used.
				//MessageBox(NULL, args.c_str(), L"OpenWithPP cmdLine", MB_OK);




				std::wstring verb;

				if (g_Items[i]->RunAsAdmin || GetKeyState(VK_SHIFT) < 0)
					verb = L"runas";

				std::wstring path = g_Items[i]->Path;
				std::wstring var(L"%");

				if (path.find(var) != std::string::npos)
				{
					WCHAR szPath[500];
					DWORD result = ExpandEnvironmentStrings(path.c_str(), szPath, std::size(szPath));

					if (result)
						path = szPath;
				}

				WCHAR szExeDir[500];
				SHRegGetPath(HKEY_CURRENT_USER, L"Software\\" PRODUCT_NAME, L"ExeDir", szExeDir, NULL);
				std::wstring exeDir(szExeDir);

				// starts with
				if (path.rfind(L"..\\", 0) == 0)
					path = exeDir + path;

				if (args.find(var) != std::string::npos)
				{
					WCHAR szArgs[900];
					DWORD result = ExpandEnvironmentStrings(args.c_str(), szArgs, std::size(szArgs));

					if (result)
						args = szArgs;
				}

				std::wstring guiExe(L"OpenWithPPGUI.exe");

				if (guiExe == path)
					path = exeDir + guiExe;

				SHELLEXECUTEINFO info;

				info.cbSize = sizeof(SHELLEXECUTEINFO);
				info.fMask = NULL;
				info.hwnd = hwnd;
				info.lpVerb = verb.c_str();
				info.lpFile = path.c_str();
				info.lpParameters = args.c_str();

				if (g_Items[i]->WorkingDirectory.length() == 0)
				{
					WCHAR szDir[500];

					if (g_ShellItems.size() > 0)
					{
						std::wstring firstItem = *g_ShellItems.begin();

						if (FileExists(firstItem))
						{
							wcscpy_s(szDir, firstItem.c_str());
							PathRemoveFileSpec(szDir);
						}
						else if (DirectoryExist(firstItem))
						{
							wcscpy_s(szDir, firstItem.c_str());
						}
					}

					info.lpDirectory = DirectoryExist(szDir) ? szDir : NULL;
				}
				else
				{
					if (g_Items[i]->WorkingDirectory.find(var) != std::string::npos)
					{
						WCHAR szWorkingDirectory[500];
						DWORD result = ExpandEnvironmentStrings(
							g_Items[i]->WorkingDirectory.c_str(), szWorkingDirectory, std::size(szWorkingDirectory));

						if (result)
							g_Items[i]->WorkingDirectory = szWorkingDirectory;
					}

					info.lpDirectory = g_Items[i]->WorkingDirectory.c_str();
				}

				info.nShow = g_Items[i]->HideWindow ? SW_HIDE : SW_NORMAL;
				info.hInstApp = NULL;

				ShellExecuteEx(&info);
				
			}
			return S_OK;
		}
	}

	if (id == g_EditCommandIndex)
	{
		WCHAR path[500];
		SHRegGetPath(HKEY_CURRENT_USER, L"Software\\" PRODUCT_NAME, L"ExeLocation", path, NULL);

		SHELLEXECUTEINFO info;

		info.cbSize = sizeof(SHELLEXECUTEINFO);
		info.fMask = NULL;
		info.hwnd = hwnd;
		info.lpVerb = NULL;
		info.lpFile = path;
		info.lpParameters = NULL;
		info.lpDirectory = NULL;
		info.nShow = SW_NORMAL;
		info.hInstApp = NULL;

		ShellExecuteEx(&info);
		return S_OK;
	}

	return E_FAIL;
}


STDMETHODIMP CMain::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax)
{
	return S_OK;
}
