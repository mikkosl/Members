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
#include <sqlite3.h>
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
#define IDC_CITY      203
#define IDC_ROW       204
#define IDC_ADD_BUTTON 205
#define IDC_REMOVE_BUTTON 206
#define IDC_MORE_BUTTON 207
#define IDC_BACK_BUTTON 208

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
std::wstring memberList[500];
int rc;
sqlite3_stmt* stmt = nullptr;
sqlite3* db = nullptr;
std::string filePath;
int currentPage = 0;
const int MEMBERS_PER_PAGE = 40;
HWND hAddButton = NULL;
HWND hRemoveButton = NULL;
HWND hMoreButton = NULL;
HWND hBackButton = NULL;

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

void ParseMemberEntry(const std::wstring& entry, std::wstring& last, std::wstring& first, std::wstring& town);
void ImportMembersFromCSV(const std::wstring& csvPath);

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
    
    sqlite3_close(db);
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
    sqlite3* db = nullptr;

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
                        "city TEXT);";
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
                        "city TEXT);";

                    if (sqlite3_open(filePath.c_str(), &db) != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database: ", MB_OK | MB_ICONERROR);
                        return 1;
                    }

                    if (sqlite3_exec(db, createTableSQL, nullptr, nullptr, &errMsg) != SQLITE_OK) {
                        MessageBox(nullptr, utf8ToWstring(errMsg).c_str(), L"Table creation failed: ", MB_OK | MB_ICONERROR);
                        sqlite3_free(errMsg);
                        sqlite3_close(db);                    // Close the database
                    }

                    const char* sql = "SELECT surname, firstName, city FROM Members;";
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
                        const char* city = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 2));
                        memberList[i] = utf8ToWstring(surname) + L", " + utf8ToWstring(firstName) + L", " + utf8ToWstring(city) + L"\n";
                        i++;
                    }
                    if (rc != SQLITE_DONE) MessageBox(nullptr, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to retrieve data: ", MB_OK | MB_ICONERROR);
                    sqlite3_finalize(stmt);
                    InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the rows
                }
            }
            break;
            case ID_FILE_IMPORT:
                ImportMembersFromCSV(L"Members.csv");
                InvalidateRect(GetActiveWindow(), NULL, TRUE);
            break;
            case ID_MEMBER_ADD:
                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"City:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_CITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);

                ShowWindow(hRemoveButton, SW_HIDE);
                hAddButton = CreateWindow(L"BUTTON", L"Add Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_ADD_BUTTON, NULL, NULL);
            break;
            case IDC_ADD_BUTTON:
            {
                wchar_t firstName[100], surname[100], city[100];
                GetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), surname, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), firstName, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_CITY), city, 100);

                std::string last = toUtf8(surname);
                std::string first = toUtf8(firstName);
                std::string town = toUtf8(city);
                const char* sql = "INSERT INTO Members (surname, firstName, city) VALUES (?, ?, ?);";

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
                    SetWindowTextW(GetDlgItem(hWnd, IDC_CITY), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_ROW), L"");
                }

                sqlite3_finalize(stmt);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display new rows
                return 0;
            }
            break;
            case ID_MEMBER_REMOVE:
            {
                CreateWindow(L"STATIC", L"Surname:", WS_VISIBLE | WS_CHILD,
                    600, 20, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 20, 200, 20, hWnd, (HMENU)IDC_SURNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"First name:", WS_VISIBLE | WS_CHILD,
                    600, 50, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 50, 200, 20, hWnd, (HMENU)IDC_FIRSTNAME, NULL, NULL);

                CreateWindow(L"STATIC", L"City:", WS_VISIBLE | WS_CHILD,
                    600, 80, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 80, 200, 20, hWnd, (HMENU)IDC_CITY, NULL, NULL);

                CreateWindow(L"STATIC", L"Row:", WS_VISIBLE | WS_CHILD,
                    600, 110, 80, 20, hWnd, NULL, NULL, NULL);
                CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
                    710, 110, 200, 20, hWnd, (HMENU)IDC_ROW, NULL, NULL);
                
                ShowWindow(hAddButton, SW_HIDE);
                hRemoveButton = CreateWindow(L"BUTTON", L"Remove Member", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    610, 140, 200, 30, hWnd, (HMENU)IDC_REMOVE_BUTTON, NULL, NULL);
            }
            break;
            case IDC_REMOVE_BUTTON:
            {
                wchar_t firstName[100], surname[100], city[100];
                std::string last;
                std::string first;
                std::string town; 
                
                GetWindowTextW(GetDlgItem(hWnd, IDC_SURNAME), surname, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_FIRSTNAME), firstName, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_CITY), city, 100);
                GetWindowTextW(GetDlgItem(hWnd, IDC_ROW), row, 100);

                if (row[0]  == L'\0') {
                    last = toUtf8(surname);
                    first = toUtf8(firstName);
                    town = toUtf8(city);
                }
                else {
                    rowNumber = _wtoi(row);
                    std::wstring wlast, wfirst, wtown;
                    ParseMemberEntry(memberList[rowNumber - 1], wlast, wfirst, wtown);
                    last = toUtf8(wlast.c_str());
                    first = toUtf8(wfirst.c_str());
                    town = toUtf8(wtown.c_str());
                }
                                 
                const char* sql = "DELETE FROM Members WHERE surname = ? AND firstName = ? AND city = ?;";

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
                    SetWindowTextW(GetDlgItem(hWnd, IDC_CITY), L"");
                    SetWindowTextW(GetDlgItem(hWnd, IDC_ROW), L"");
                }

                sqlite3_finalize(stmt);
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the rows
            }
            break; 
            case IDC_MORE_BUTTON:
            {
                currentPage++;
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the next page
            }
			break;
            case IDC_BACK_BUTTON:
            {
                if (currentPage > 0) currentPage--;
                InvalidateRect(hWnd, NULL, TRUE);   // Force a repaint to display the previous page
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
                if (sqlite3_open("Members.db", &db) != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to open database", MB_OK | MB_ICONERROR);
                    return 0;
                }
                const char* sql = "SELECT surname, firstName, city FROM Members;";
                rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
                if (rc != SQLITE_OK) {
                    MessageBox(hWnd, utf8ToWstring(sqlite3_errmsg(db)).c_str(), L"Failed to prepare statement: ", MB_OK | MB_ICONERROR);
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
                    const char* city = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 2));
                    memberList[i] = utf8ToWstring(surname) + L", " + utf8ToWstring(firstName) + L", " + utf8ToWstring(city) + L"\n";
                    i++;
                }
                    return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            int x = 20;
            int y = 10;
			std::wstring rowStr;
            
            int totalMembers = 0;
            for (int i = 0; i < 500 && !memberList[i].empty(); ++i) ++totalMembers;
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
                        600, 350, 100, 30, hWnd, (HMENU)IDC_MORE_BUTTON, NULL, NULL);
                } else {
                    ShowWindow(hMoreButton, SW_SHOW);
                }
            } else {
                if (hMoreButton) ShowWindow(hMoreButton, SW_HIDE);
            }

            // Handle the Back button: hide on first 7 pages (currentPage < 7), show otherwise
            if (!hBackButton) {
                hBackButton = CreateWindow(L"BUTTON", L"Back", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
                    700, 350, 100, 30, hWnd, (HMENU)IDC_BACK_BUTTON, NULL, NULL);
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

// Helper function to parse last, first, and town from a memberList entry
void ParseMemberEntry(const std::wstring& entry, std::wstring& last, std::wstring& first, std::wstring& town)
{
    // Format: "surname, firstName, city\n"
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
    while (std::getline(file, line)) {
        std::wstringstream ss(line);
        std::wstring surname, firstname, city;
        if (std::getline(ss, surname, L';') &&
            std::getline(ss, firstname, L';') &&
            std::getline(ss, city, L';')) {
            // Remove possible whitespace
            surname.erase(0, surname.find_first_not_of(L" \t"));
            firstname.erase(0, firstname.find_first_not_of(L" \t"));
            city.erase(0, city.find_first_not_of(L" \t"));

            // Insert into database (reuse your existing DB insert logic)
            std::string last = toUtf8(surname.c_str());
            std::string first = toUtf8(firstname.c_str());
            std::string town = toUtf8(city.c_str());
            const char* sql = "INSERT INTO Members (surname, firstName, city) VALUES (?, ?, ?);";
            sqlite3_stmt* stmt = nullptr;
			int i = static_cast<int>(memberList->size());
            if (sqlite3_open(filePath.c_str(), &db) == SQLITE_OK &&
                sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
                sqlite3_bind_text(stmt, 1, last.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 2, first.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(stmt, 3, town.c_str(), -1, SQLITE_TRANSIENT);
                sqlite3_step(stmt);
                sqlite3_finalize(stmt);
                memberList[i] = surname + L", " + firstname + L", " + city + L"\n";                i++;
                i++;
            }
        }
    }
    file.close();
}
