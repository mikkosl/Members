// Members.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "Members.h"
#include <string>
#include <locale>
#include <vector>
#include <commdlg.h> // Add this at the top of your file
#include <fstream>
#include <sstream>
#include <Windows.h>
#include "sqlite3.h"
#include <iostream>
#include <string>
#include <cwchar>
#include <cstdlib>
#include <stdexcept>
#include <cstdio>
#include <memory>
#include <array>

#define MAX_LOADSTRING 100
#define IDC_FIRSTNAME 201
#define IDC_SURNAME   202
#define IDC_MUNICIPALITY      203
#define IDC_ROW       204
#define IDC_ADD_BUTTON 205
#define IDC_REMOVE_BUTTON 206
#define IDC_MORE_BUTTON 207
#define IDC_BACK_BUTTON 208
#define IDC_SEARCH_BUTTON 209
#define IDC_UPDATE_BUTTON 218

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
std::wstring memberList[500];
int rc;
sqlite3_stmt* stmt = nullptr;
std::string filePath;
int currentPage = 0;
const int MEMBERS_PER_PAGE = 40;
HWND hAddButton = NULL;
HWND hRemoveButton = NULL;
HWND hMoreButton = NULL;
HWND hBackButton = NULL;
HWND hSearchButton = NULL;
HWND hUpdateButton = NULL;
HWND hSurname = NULL;
HWND hFirstName = NULL;
HWND hMunicipality = NULL;
HWND hRow = NULL;
sqlite3* db = nullptr; // Add a global variable for the database connection
int totalMembers = 0;

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    GettingStartedDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);

void ParseMemberEntry(const std::wstring& entry, std::wstring& last, std::wstring& first, std::wstring& town);
void ImportMembersFromCSV(const std::wstring& csvPath);
void ExportMembersToCSV(const std::wstring& csvPath);
void LoadMembers(int page);

std::string toUtf8(const wchar_t* wstr) {
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, nullptr, 0, nullptr, nullptr);
    std::string result(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr, -1, &result[0], size_needed, nullptr, nullptr);
    return result;
}

// Converts UTF-8 encoded const char* to std::wstring without <codecvt>

std::wstring utf8ToWstring(const char* utf8Str) {
    if (!utf8Str) return L"";

    int size_needed = MultiByteToWideChar(CP_UTF8, 0, utf8Str, -1, nullptr, 0);
    if (size_needed <= 0) return L"";

    std::wstring result(size_needed - 1, 0); // exclude null terminator
    MultiByteToWideChar(CP_UTF8, 0, utf8Str, -1, &result[0], size_needed);

    return result;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_MEMBERS, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_MEMBERS));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    
    if (db) {
        sqlite3_close(db);
    }
    return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MEMBERS));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_MEMBERS);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    std::wstring last, first, town;

    switch (message)
    {
    case WM_COMMAND:
    {
            int wmId = LOWORD(wParam);
            wchar_t row[100] = L"";
            int rowNumber = 0;
            // Parse the menu selections:
            switch (wmId)
            {
            case ID_FILE_CREATE:
            {
                OPENFILENAME ofn = { 0 };
                TCHAR szFile[MAX_PATH] = { 0 };
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hWnd;
                ofn.lpstrFile = szFile;
                ofn.nMaxFile = MAX_PATH;
                ofn.lpstrFilter = L"Database Files (*.db)\0*.db\0All Files (*.*)\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
                if (GetSaveFileName(&ofn)) {
                    // Convert TCHAR to std::string (UTF-8) for sqlite3_open
                    filePath = toUtf8(szFile);
                    char* errMsg = nullptr;
                    const char* createTableSQL =
                        "CREATE TABLE IF NOT EXISTS Members ("
                        "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                        "surname TEXT NOT NULL, "
                        "firstName TEXT NOT NULL, "
                        "municipality TEXT);";
                    if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database: ", MB_OK | MB_ICONERROR);
                        return 1;
                    }
                    if (sqlite3_exec(db, createTableSQL, nullptr, nullptr, &errMsg) != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(errMsg).c_str(), L"Table creation failed: ", MB_OK | MB_ICONERROR);
                        sqlite3_free(errMsg);
                        sqlite3_close(db);                    // Close the database
                    }
                    sqlite3_close(db);                    // Close the database
                }
			}
			break;
            case ID_FILE_OPEN:
            {
			    OPENFILENAME ofn = { 0 };
                TCHAR szFile[MAX_PATH] = { 0 };
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hWnd;
                ofn.lpstrFile = szFile;
                ofn.nMaxFile = MAX_PATH;
                ofn.lpstrFilter = L"Database Files (*.db)\0*.db\0All Files (*.*)\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

                if (GetOpenFileName(&ofn)) {
                    // Convert TCHAR to std::string (UTF-8) for sqlite3_open
                    filePath = toUtf8(szFile);

                    char* errMsg = nullptr;
                    const char* createTableSQL =
                        "CREATE TABLE IF NOT EXISTS Members ("
                        "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                        "surname TEXT NOT NULL, "
                        "firstName TEXT NOT NULL, "
                        "municipality TEXT);";

                    if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database: ", MB_OK | MB_ICONERROR);
                        return 1;
                    }

                    if (sqlite3_exec(db, createTableSQL, nullptr, nullptr, &errMsg) != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(errMsg).c_str(), L"Table creation failed: ", MB_OK | MB_ICONERROR);
                        sqlite3_free(errMsg);
                        sqlite3_close(db);                    // Close the database
                    }

                    const char* sql = "SELECT surname, firstName, municipality FROM Members;";
                    rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
                    if (rc != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to prepare statement: ", MB_OK | MB_ICONERROR);
                        sqlite3_close(db);
                        return 0;
                    }
                    int i = 0;
                    for (int j = 0; j < 500; ++j) {
                        memberList[j].clear();
                    }
                    while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
                        const char* surname = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 0));
                        const char* firstName = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1));
                        const char* municipality = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 2));
                        memberList[i] = utf8ToWstring(surname) + L", " + utf8ToWstring(firstName) + L", " + utf8ToWstring(municipality) + L"\n";
                        i++;
                    }
                    if (rc != SQLITE_DONE) MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to retrieve data: ", MB_OK | MB_ICONERROR);
                    sqlite3_finalize(stmt);
                    LoadMembers(currentPage);
                    InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the rows
                }
            }
            break;
            case ID_FILE_CLOSE:
            {
                sqlite3_close(db);
                for (int j = 0; j < 500; ++j) {
                    memberList[j].clear();
                }
                currentPage = 0;
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to clear the rows
            }
            break;
            case ID_FILE_IMPORT:
            {
                OPENFILENAME ofn = { 0 };
                TCHAR szFile[MAX_PATH] = { 0 };
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hWnd;
                ofn.lpstrFile = szFile;
                ofn.nMaxFile = MAX_PATH;
                ofn.lpstrFilter = L"Database Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

                if (GetOpenFileName(&ofn)) {
                    ImportMembersFromCSV(szFile);
                    LoadMembers(currentPage);
                    InvalidateRect(GetActiveWindow(), NULL, TRUE);
                }
            }
            break;
            case ID_FILE_EXPORT:
            {
                OPENFILENAME ofn = { 0 };
                TCHAR szFile[MAX_PATH] = { 0 };
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hWnd;
                ofn.lpstrFile = szFile;
                ofn.nMaxFile = MAX_PATH;
                ofn.lpstrFilter = L"Database Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.Flags = OFN_PATHMUSTEXIST;

                if (GetSaveFileName(&ofn)) {
                    ExportMembersToCSV(szFile);
                }
            }
            break;
            case ID_MEMBER_ADD:
                HWND hOld;
                
                hOld = GetDlgItem(hWnd, IDC_SURNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_FIRSTNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_MUNICIPALITY);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ROW);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_UPDATE_BUTTON);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_REMOVE_BUTTON);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_SEARCH_BUTTON);
                if (hOld) DestroyWindow(hOld);

                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"Municipality:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_MUNICIPALITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);

                hAddButton = CreateWindow(L"BUTTON", L"Add Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_ADD_BUTTON, NULL, NULL);
            break;
            case IDC_ADD_BUTTON:
            {
                wchar_t firstName[100], surname[100], municipality[100];
                GetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), surname, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), firstName, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_MUNICIPALITY), municipality, 100);

                std::string last = toUtf8(surname);
                std::string first = toUtf8(firstName);
                std::string town = toUtf8(municipality);
                const char* sql = "INSERT INTO Members (surname, firstName, municipality) VALUES (?, ?, ?);";

                if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                   MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
                   return 0;
                }

                rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
                if (rc != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to prepare statement: ", MB_OK | MB_ICONERROR);
                    sqlite3_close(db);
                    return 0;
                }

                sqlite3_bind_text(stmt, 1, last.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 2, first.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 3, town.c_str(), -1, SQLITE_TRANSIENT);

                rc = sqlite3_step(stmt);
                if (rc != SQLITE_DONE) {
                    std::wstring errMsgW = utf8ToWstring(sqlite3_errmsg(db));
                    MessageBox(hWnd, errMsgW.c_str(), L"Insert failed: ", MB_OK | MB_ICONINFORMATION);
                } else {
                    // Clear the input fields after successful insert
                    SetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_MUNICIPALITY), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_ROW), L"");
                }

                sqlite3_finalize(stmt);
                LoadMembers(currentPage);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display new rows
                return 0;
            }
            break;
            case ID_MEMBER_REMOVE:
            {
                HWND hOld;
                
                hOld = GetDlgItem(hWnd, IDC_SURNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_FIRSTNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_MUNICIPALITY);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ROW);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_UPDATE_BUTTON);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ADD_BUTTON);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_SEARCH_BUTTON);
                if (hOld) DestroyWindow(hOld);

                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"Municipality:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_MUNICIPALITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);
                
                hRemoveButton = CreateWindow(L"BUTTON", L"Remove Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_REMOVE_BUTTON, NULL, NULL);
            }
            break;
            case IDC_REMOVE_BUTTON:
            {
                wchar_t firstName[100], surname[100], municipality[100];
                std::string last;
                std::string first;
                std::string town; 
                
                GetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), surname, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), firstName, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_MUNICIPALITY), municipality, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_ROW), row, 100);

                if (row[0]  == L'\0') {
                    last = toUtf8(surname);
                    first = toUtf8(firstName);
                    town = toUtf8(municipality);
                }
                else {
                    rowNumber = _wtoi(row);
                    std::wstring wlast, wfirst, wtown;
                    ParseMemberEntry(memberList[rowNumber - 1], wlast, wfirst, wtown);
                    last = toUtf8(wlast.c_str());
                    first = toUtf8(wfirst.c_str());
                    town = toUtf8(wtown.c_str());
                }
                                 
                const char* sql = "DELETE FROM Members WHERE surname = ? AND firstName = ? AND municipality = ?;";

                if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
                    return 0;
                }

                rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
                if (rc != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to prepare statement: ", MB_OK | MB_ICONERROR);
                    sqlite3_close(db);
                    return 0;
                }

                sqlite3_bind_text(stmt, 1, last.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 2, first.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 3, town.c_str(), -1, SQLITE_TRANSIENT);

                rc = sqlite3_step(stmt);
                if (rc != SQLITE_DONE) {
                    std::wstring errMsgW = utf8ToWstring(sqlite3_errmsg(db));
                    MessageBox(hWnd, errMsgW.c_str(), L"Remove failed: ", MB_OK | MB_ICONINFORMATION);
                }
                else {
                    // Clear the input fields after successful delete
                    SetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_MUNICIPALITY), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_ROW), L"");
                }

                sqlite3_finalize(stmt);
                LoadMembers(currentPage);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the rows
            }
            break; 
            case ID_MEMBER_UPDATE:
            {
                HWND hOld;

                hOld = GetDlgItem(hWnd, IDC_SURNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_FIRSTNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_MUNICIPALITY);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ROW);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ADD_BUTTON);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_REMOVE_BUTTON);
                if (hOld) DestroyWindow(hOld);

                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"Municipality:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_MUNICIPALITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);

                hSearchButton = CreateWindow(L"BUTTON", L"Search Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_SEARCH_BUTTON, NULL, NULL);

            }
            break;
            case IDC_SEARCH_BUTTON:
            {
                wchar_t surname[100];
                std::string last;
                std::string first;
                std::string town;
                HWND hOld;
                std::wstring surnameW, firstNameW, municipalityW, rowW;
                const char* sql;

                GetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), surname, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_ROW), row, 100);

                if (row[0] == L'\0') {
                    last = toUtf8(surname);
                    sql = "SELECT surname, firstName, municipality FROM Members WHERE surname = ?;";
					if (surname[0] == L'\0') {
                        MessageBoxW(NULL, L"Please enter a surname or row number.", L"Debug", MB_OK);
                        return 0;
                    }
                }
                else {
                    rowNumber = _wtoi(row);
                    std::wstring wlast, wfirst, wtown;
                    ParseMemberEntry(memberList[rowNumber - 1], wlast, wfirst, wtown);
                    last = toUtf8(wlast.c_str());
                    first = toUtf8(wfirst.c_str());
                    town = toUtf8(wtown.c_str());
                    sql = "SELECT surname, firstName, municipality FROM Members WHERE surname = ? AND firstname = ? AND municipality = ?;";
                }

                if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
                    return 0;
                }

                rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);

                if (rc != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to prepare statement: ", MB_OK | MB_ICONERROR);
                    sqlite3_close(db);
                    return 0;
                }

                if (row[0] == L'\0') {
                    sqlite3_bind_text(stmt, 1, last.c_str(), -1, SQLITE_TRANSIENT);
                }
                else {
                    sqlite3_bind_text(stmt, 1, last.c_str(), -1, SQLITE_TRANSIENT);
                    sqlite3_bind_text(stmt, 2, first.c_str(), -1, SQLITE_TRANSIENT);
                    sqlite3_bind_text(stmt, 3, town.c_str(), -1, SQLITE_TRANSIENT);
                    rowW = std::to_wstring(rowNumber);
                }

                bool found = false;
                while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
                    found = true;
                    const char* surnameUtf8 = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 0));
                    surnameW = utf8ToWstring(surnameUtf8);
                    const char* firstNameUtf8 = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1));
                    firstNameW = utf8ToWstring(firstNameUtf8);
                    const char* municipalityUtf8 = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 2));
                    municipalityW = utf8ToWstring(municipalityUtf8);
                    // process row
                }
                if (!found) {
                    MessageBoxW(NULL, L"No results found.", L"Debug", MB_OK);
					return 0;
                }

                if (row[0] == L'\0') {
                    rowNumber = 0;
                    while (rowNumber < 500 && !memberList[rowNumber].empty()) {
                        std::wstring wlast, wfirst, wtown;
                        ParseMemberEntry(memberList[rowNumber], wlast, wfirst, wtown);
                        if (wlast == surnameW) {
                            firstNameW = wfirst;
                            municipalityW = wtown;
                            break;
                        }
                        rowNumber++;
                    }
                    rowW = std::to_wstring(rowNumber + 1);
                }
                hOld = GetDlgItem(hWnd, IDC_SURNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_FIRSTNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_MUNICIPALITY);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ROW);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_SEARCH_BUTTON);
                if (hOld) DestroyWindow(hOld);

                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                hSurname = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                hFirstName = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"Municipality:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                hMunicipality = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_MUNICIPALITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                hRow = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);
                    
                hUpdateButton = CreateWindow(L"BUTTON", L"Update Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_UPDATE_BUTTON, NULL, NULL); 
             
                SetWindowTextW(hSurname, surnameW.c_str());
                SetWindowTextW(hFirstName, firstNameW.c_str());
                SetWindowTextW(hMunicipality, municipalityW.c_str());
                SetWindowTextW(hRow, rowW.c_str());
            }
            break;
            case IDC_UPDATE_BUTTON:
            {
                wchar_t firstName[100], surname[100], municipality[100], row[100];
                int rowNumber = 0;
                HWND hOld;

                GetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), surname, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), firstName, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_MUNICIPALITY), municipality, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_ROW), row, 100);
                
                rowNumber = _wtoi(row);
                std::wstring wlast, wfirst, wtown;
                ParseMemberEntry(memberList[rowNumber - 1], wlast, wfirst, wtown);
                std::string oldLast = toUtf8(wlast.c_str());
                std::string oldFirst = toUtf8(wfirst.c_str());
                std::string oldTown = toUtf8(wtown.c_str());
                std::string newLast = toUtf8(surname);
                std::string newFirst = toUtf8(firstName);
                std::string newTown = toUtf8(municipality);
                const char* sql = "UPDATE Members SET surname = ?, firstName = ?, municipality = ? WHERE surname = ? AND firstName = ? AND municipality = ?;";
                if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
                    return 0;
                }
                rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
                if (rc != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to prepare statement: ", MB_OK | MB_ICONERROR);
                    sqlite3_close(db);
                    return 0;
                }
                sqlite3_bind_text(stmt, 1, newLast.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 2, newFirst.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 3, newTown.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 4, oldLast.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 5, oldFirst.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 6, oldTown.c_str(), -1, SQLITE_TRANSIENT);
                
                rc = sqlite3_step(stmt);
                if (rc != SQLITE_DONE) {
                    std::wstring errMsgW = utf8ToWstring(sqlite3_errmsg(db));
                    MessageBox(hWnd, errMsgW.c_str(), L"Update failed: ", MB_OK | MB_ICONINFORMATION);
                }
                else {
                    // Optionally check how many rows were changed:
                    int changes = sqlite3_changes(db);
                    if (changes == 0) {
                        MessageBoxW(hWnd, L"No rows were updated (original values not found).", L"Update", MB_OK);
                    }
                }
                memberList[rowNumber-1] = utf8ToWstring(newLast.c_str()) + L", " + utf8ToWstring(newFirst.c_str()) + L", " + utf8ToWstring(newTown.c_str()) + L"\n";
                
                sqlite3_finalize(stmt);

                LoadMembers(currentPage);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the updated rows
                
                hOld = GetDlgItem(hWnd, IDC_SURNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_FIRSTNAME);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_MUNICIPALITY);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_ROW);
                if (hOld) DestroyWindow(hOld);
                hOld = GetDlgItem(hWnd, IDC_UPDATE_BUTTON);
                if (hOld) DestroyWindow(hOld);

                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"Municipality:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_MUNICIPALITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);

                hSearchButton = CreateWindow(L"BUTTON", L"Search Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_SEARCH_BUTTON, NULL, NULL);
            
            }
            break;
            case ID_MEMBER_DELETEALL:
                {
                // Ask user to confirm with OK and Cancel buttons
                int res = MessageBoxW(hWnd,
                    L"All members will be deleted from the database. This action cannot be undone.\n\nDo you want to proceed?",
                    L"Warning",
                    MB_OKCANCEL | MB_ICONWARNING | MB_DEFBUTTON2);

                if (res != IDOK) {
                    // User cancelled -> do nothing
                    break;
                }

                const char* sql = "DELETE FROM Members;";
                if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
                    return 0;
                }
                rc = sqlite3_exec(db, sql, nullptr, nullptr, nullptr);
                if (rc != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to delete all members: ", MB_OK | MB_ICONERROR);
                }
                sqlite3_close(db);
                for (int i = 0; i < 500; ++i) {
                    memberList[i].clear();
                }
                currentPage = 0;
                LoadMembers(currentPage);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the rows
            }
            break;
            case IDC_MORE_BUTTON:
            {
                currentPage++;
                LoadMembers(currentPage);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the next page
            }
			break;
            case IDC_BACK_BUTTON:
            {
                if (currentPage > 0) currentPage--;
                LoadMembers(currentPage);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the previous page
			}
			break;
            case IDM_GETTING_STARTED:
            {
				DialogBox(hInst, MAKEINTRESOURCE(IDD_GETTING_STARTED), hWnd, GettingStartedDlgProc);
            }
			break;
            case IDM_ABOUT:
            {
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            }
            break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
            break;
            default:
				return DefWindowProc(hWnd, message, wParam, lParam);
            }
    }
    break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            int x = 20;
            int y = 10;
			std::wstring rowStr;

            int startIdx = currentPage * MEMBERS_PER_PAGE;
            int endIdx = startIdx + MEMBERS_PER_PAGE;
            if (endIdx > totalMembers) endIdx = totalMembers;

            // Corrected bounds-checked version
            for (int i = startIdx; i < endIdx && i >= 0 && i < 500 && !memberList[i].empty(); ++i) {
                if (i < 9) rowStr = L"[00" + std::to_wstring(i + 1) + L"]: " + memberList[i].c_str();
                else if (i < 99) rowStr = L"[0" + std::to_wstring(i + 1) + L"]: " + memberList[i].c_str();
                else rowStr = L"[" + std::to_wstring(i + 1) + L"]: " + memberList[i].c_str();

                if (i % 20 == 0 && i > 0) {
                    x += 300;
                    y = 10;
                }
				if (i % 40 == 0 && i > 0) {
                    x = 20;
                    y = 10;
                }

                TextOutW(hdc, x, y, rowStr.c_str(), (int)memberList[i].length() + 7);
                y += 20; // Move down for next line
            }
            // After drawing, handle the More button:
            if (endIdx < totalMembers) {
                if (!hMoreButton) {
                    hMoreButton = CreateWindow(L"BUTTON", L"More", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                        700, 350, 100, 30, hWnd, (HMENU)IDC_MORE_BUTTON, NULL, NULL);
                } else {
                    ShowWindow(hMoreButton, SW_SHOW);
                }
            } else {
                if (hMoreButton) ShowWindow(hMoreButton, SW_HIDE);
            }

            // Handle the Back button: hide on first 7 pages (currentPage < 7), show otherwise
            if (!hBackButton) {
                hBackButton = CreateWindow(L"BUTTON", L"Back", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    600, 350, 100, 30, hWnd, (HMENU)IDC_BACK_BUTTON, NULL, NULL);
            }
            if (currentPage == 0) {
                ShowWindow(hBackButton, SW_HIDE);
            } else {
                ShowWindow(hBackButton, SW_SHOW);
            }

            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

INT_PTR CALLBACK GettingStartedDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Helper function to parse last, first, and town from a memberList entry
void ParseMemberEntry(const std::wstring& entry, std::wstring& last, std::wstring& first, std::wstring& town)
{
    // Format: "surname, firstName, municipality\n"
    size_t pos1 = entry.find(L", ");
    if (pos1 == std::wstring::npos) { last = first = town = L""; return; }
    last = entry.substr(0, pos1);

    size_t pos2 = entry.find(L", ", pos1 + 2);
    if (pos2 == std::wstring::npos) { first = town = L""; return; }
    first = entry.substr(pos1 + 2, pos2 - (pos1 + 2));

    // Remove trailing newline if present
    size_t end = entry.find_last_not_of(L"\r\n");
    if (end == std::wstring::npos || pos2 + 2 > end) { town = L""; return; }
    town = entry.substr(pos2 + 2, end - (pos2 + 2) + 1);
}

// Example function to import members from a CSV file
void ImportMembersFromCSV(const std::wstring& csvPath) {

    std::wifstream file(csvPath);
    if (!file.is_open()) {
        MessageBox(nullptr, L"Failed to open CSV file.", L"Import Error", MB_OK | MB_ICONERROR);
        return;
    }
    std::wstring line;
    int i = 0;
    for (; i < 500 && !memberList[i].empty(); ++i);
    while (std::getline(file, line) && i < 500) {
        std::wstringstream ss(line);
        std::wstring surname, firstname, municipality;
        if (std::getline(ss, surname, L';') &&
            std::getline(ss, firstname, L';') &&
            std::getline(ss, municipality, L';')) {
            // Remove possible whitespace
            surname.erase(0, surname.find_first_not_of(L" \t"));
            firstname.erase(0, firstname.find_first_not_of(L" \t"));
            municipality.erase(0, municipality.find_first_not_of(L" \t"));

            // Insert into database (reuse your existing DB insert logic)
            std::string last = toUtf8(surname.c_str());
            std::string first = toUtf8(firstname.c_str());
            std::string town = toUtf8(municipality.c_str());
            const char* sql = "INSERT INTO Members (surname, firstName, municipality) VALUES (?, ?, ?);";
            sqlite3_stmt* stmt = nullptr;
            if (sqlite3_open(filePath.c_str(), &db) == SQLITE_OK &&
                sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
                sqlite3_bind_text(stmt, 1, last.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 2, first.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 3, town.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_step(stmt);
                sqlite3_finalize(stmt);
                memberList[i] = std::wstring(surname) + L", " + std::wstring(firstname) + L", " + std::wstring(municipality) + L"\n";
                i++;
            }
        }
    }
    file.close();
}

void ExportMembersToCSV(const std::wstring& csvPath) {
    if (filePath.empty()) {
        MessageBoxW(nullptr, L"No database is open. Open a .db file first.", L"Export Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Convert wide path to UTF-8 for std::ofstream (consistent with other code)
    std::string csvPathUtf8 = toUtf8(csvPath.c_str());

    std::ofstream ofs(csvPathUtf8, std::ios::binary);
    if (!ofs.is_open()) {
        MessageBoxW(nullptr, L"Failed to open target CSV file for writing.", L"Export Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Optional: write UTF-8 BOM so Excel recognizes UTF-8
    const unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
    ofs.write(reinterpret_cast<const char*>(bom), sizeof(bom));

    sqlite3* localDb = nullptr;
    sqlite3_stmt* localStmt = nullptr;

    if (sqlite3_open(filePath.c_str(), &localDb) != SQLITE_OK) {
        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(localDb)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
        if (localDb) sqlite3_close(localDb);
        ofs.close();
        return;
    }

    const char* sql = "SELECT surname, firstName, municipality FROM Members ORDER BY surname COLLATE NOCASE, firstName COLLATE NOCASE, municipality COLLATE NOCASE;";
    if (sqlite3_prepare_v2(localDb, sql, -1, &localStmt, nullptr) != SQLITE_OK) {
        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(localDb)).c_str(), L"Failed to prepare statement", MB_OK | MB_ICONERROR);
        sqlite3_close(localDb);
        ofs.close();
        return;
    }

    auto escapeField = [](const std::string& in) -> std::string {
        // CSV rules: if field contains semicolon, quote or newline, wrap in double quotes and double internal quotes.
        bool needsQuotes = false;
        for (char c : in) {
            if (c == ';' || c == '"' || c == '\n' || c == '\r') { needsQuotes = true; break; }
        }
        if (!needsQuotes) return in;
        std::string out;
        out.push_back('"');
        for (char c : in) {
            if (c == '"') out.append("\"\""); else out.push_back(c);
        }
        out.push_back('"');
        return out;
    };

    int exported = 0;
    while (sqlite3_step(localStmt) == SQLITE_ROW) {
        const unsigned char* s0 = sqlite3_column_text(localStmt, 0);
        const unsigned char* s1 = sqlite3_column_text(localStmt, 1);
        const unsigned char* s2 = sqlite3_column_text(localStmt, 2);
        std::string surname = s0 ? reinterpret_cast<const char*>(s0) : std::string();
        std::string firstname = s1 ? reinterpret_cast<const char*>(s1) : std::string();
        std::string municipality = s2 ? reinterpret_cast<const char*>(s2) : std::string();

        ofs << escapeField(surname) << ';' << escapeField(firstname) << ';' << escapeField(municipality) << '\n';
        ++exported;
    }

    sqlite3_finalize(localStmt);
    sqlite3_close(localDb);
    ofs.close();

    std::wstring msg = L"Exported " + std::to_wstring(exported) + L" members to:\n" + csvPath;
    MessageBoxW(nullptr, msg.c_str(), L"Export Complete", MB_OK | MB_ICONINFORMATION);
}

void LoadMembers(int page) {
    // clear existing
    for (int j = 0; j < 500; ++j) memberList[j].clear();
    totalMembers = 0;
    if (filePath.empty()) return;

    sqlite3* localDb = nullptr;
    sqlite3_stmt* localStmt = nullptr;
    sqlite3_stmt* countStmt = nullptr;

    if (sqlite3_open(filePath.c_str(), &localDb) != SQLITE_OK) {
        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(localDb)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
        if (localDb) sqlite3_close(localDb);
        return;
    }

    // Get total number of members
    const char* countSql = "SELECT COUNT(*) FROM Members;";
    if (sqlite3_prepare_v2(localDb, countSql, -1, &countStmt, nullptr) == SQLITE_OK) {
        if (sqlite3_step(countStmt) == SQLITE_ROW) {
            totalMembers = sqlite3_column_int(countStmt, 0);
        }
        sqlite3_finalize(countStmt);
    } else {
        sqlite3_finalize(countStmt);
    }

    // Load entire ordered list into memberList (starting at index 0)
    const char* sql =
        "SELECT surname, firstName, municipality "
        "FROM Members "
        "ORDER BY surname COLLATE NOCASE, firstName COLLATE NOCASE, municipality COLLATE NOCASE;";

    if (sqlite3_prepare_v2(localDb, sql, -1, &localStmt, nullptr) != SQLITE_OK) {
        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(localDb)).c_str(), L"Failed to prepare statement", MB_OK | MB_ICONERROR);
        sqlite3_close(localDb);
        return;
    }

    int i = 0;
    while (sqlite3_step(localStmt) == SQLITE_ROW && i < 500) {
        const char* surname = reinterpret_cast<const char*>(sqlite3_column_text(localStmt, 0));
        const char* firstName = reinterpret_cast<const char*>(sqlite3_column_text(localStmt, 1));
        const char* municipality = reinterpret_cast<const char*>(sqlite3_column_text(localStmt, 2));
        memberList[i] = utf8ToWstring(surname) + L", " + utf8ToWstring(firstName) + L", " + utf8ToWstring(municipality) + L"\n";
        ++i;
    }

    sqlite3_finalize(localStmt);
    sqlite3_close(localDb);
}
