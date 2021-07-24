#include <Windows.h>
#include <windowsx.h>
#include <tchar.h>
#include <CommCtrl.h>
#include <Shlwapi.h>
#include <fstream>

#include <string>
#include <vector>

#include <SMTP.h>

using namespace std;

LRESULT CALLBACK MyWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
BOOL OnCreate(HWND hWnd, LPCREATESTRUCT lpCreateStruct);
VOID OnCommand(HWND hWnd, INT id, HWND hwndCtl, UINT codeNotify);
VOID OnSendMessage();
string GetServerIP();
vector<string> GetRecipients();

// Идентификатор элемента управления ввода доменного имени сервера.
const int IDC_EDIT_SERVER_NAME = 2001;
// Идентификатор элемента управления ввода IP-адреса сервера.
const int IDC_IP_SERVER_IP = 2002;
// Идентификатор элемента управления ввода почты отправителя (логина).
const int IDC_EDIT_SENDER = 2003;
// Идентификатор элемента управления ввода пароля.
const int IDC_EDIT_PASSWORD = 2004;
// Идентификатор элемента управления списка получателей.
const int IDC_LISTBOX_RECIPIENTS = 2005;
// Идентификатор элемента управления ввода получателя.
const int IDC_EDIT_RECIPIENT = 2006;
// Идентификатор элемента управления кнопки добавления получателя.
const int IDC_BUTTON_ADD_RECIPIENT = 2007;
// Идентификатор элемента управления кнопки удаления получателя.
const int IDC_BUTTON_DELETE_RECIPIENT = 2008;
// Идентификатор элемента управления ввода текста письма.
const int IDC_EDIT_MESSAGE = 2009;
// Идентификатор элемента управления кнопки отправления сообщения.
const int IDC_BUTTON_SEND = 2010;
// Идентификатор элемента управления ввода темы письма.
const int IDC_EDIT_SUBJECT = 2011;
// Идентификатор элемента управления списка прикреплённых файлов.
const int IDC_LISTBOX_FILES = 2012;
// Идентификатор элемента управления кнопки добавления файла.
const int IDC_BUTTON_ADD_FILE = 2013;
// Идентификатор элемента управления кнопки удаления файла.
const int IDC_BUTTON_DELETE_FILE = 2014;

// Пустая строка.
LPCTSTR TEXT_EMPTY = TEXT("");
// Названия оконного класса главного окна.
LPCTSTR MAIN_WINDOW_CLASS = TEXT("SMTP_FORM");
// "Подпись" главного окна.
LPCTSTR MAIN_WINDOW_LABEL = TEXT("ChushkinSMTP");
// Максимальная длина строки, введённой в элемент управления ввода.
const int EDIT_LIMIT = 50;
// Максимальная длина отправляемого текста.
const int MESSAGE_LIMIT = 1000;
// Максимальное количество получателей почты.
const int RECIPIENT_LIMIT = 5;
// Максимальное количество прикреплённых файлов.
const int FILE_LIMIT = 5;
// Максимальнs размер прикрепляемого файла (в Мбайт).
const int FILE_SIZE_LIMIT = 10;

// Дескриптор главного окна.
HWND MainWindow = NULL;
// Количество получателей почты.
int RecipientCount = 0;
// Количество приложенных файлов.
int FileCount = 0;
// Вектор имён файлов.
vector<string> FileNames;

/**
 * Главная функция.
 * @param hInstance - дескриптор экземпляра оконного сообщения.
 * @param nCmdShow - значение, передаваемое функции ShowWindow.
 */
INT WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE, LPTSTR, INT nCmdShow)
{
	InitializeSMTP();
	WNDCLASSEX wcex = { sizeof(WNDCLASSEX) };
	wcex.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
	wcex.lpfnWndProc = MyWindowProc;
	wcex.hInstance = hInstance;
	wcex.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
	wcex.lpszClassName = MAIN_WINDOW_CLASS;
	wcex.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
	RegisterClassEx(&wcex);

	LoadLibrary(TEXT("ComCtl32.dll"));

	MainWindow = CreateWindowEx(0, MAIN_WINDOW_CLASS, MAIN_WINDOW_LABEL, WS_OVERLAPPEDWINDOW & ~WS_SIZEBOX & ~WS_MAXIMIZEBOX, CW_USEDEFAULT, 0, 1110, 640, NULL, NULL, hInstance, NULL);
	ShowWindow(MainWindow, nCmdShow);

	MSG message;
	BOOL result;
	for (;;)
	{
		result = GetMessage(&message, NULL, 0, 0);
		if (result == -1)
		{
			break;
		}
		else if (result == FALSE)
		{
			break;
		}
		else
		{
			TranslateMessage(&message);
			DispatchMessage(&message);
		}
	}

	return (int)message.wParam;
}

/**
 * Оконная процедура главного окна.
 * @param hWnd - дескриптор главного окна.
 * @param uMsg - код сообщения.
 * @param wParam - параметр оконной процедуры.
 * @param lParam - параметр оконной процедуры.
 * @return результат обработки оконного сообщения.
 */
LRESULT CALLBACK MyWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_CREATE:
		return (HANDLE_WM_CREATE(hWnd, wParam, lParam, OnCreate));

	case WM_COMMAND:
		return (HANDLE_WM_COMMAND(hWnd, wParam, lParam, OnCommand));

	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	}

	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

/**
 * Обработка отображение главного окна.
 * @param hWnd - дескриптор главного окна.
 * @param lpCreateStruct - структура, описывающая созданное окна.
 */
BOOL OnCreate(HWND hWnd, LPCREATESTRUCT lpCreateStruct)
{
	CreateWindowEx(0, TEXT("Static"), TEXT("Доменное имя сервера"), WS_CHILD | WS_VISIBLE, 17, 17, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	HWND hTemp = CreateWindowEx(0, TEXT("Edit"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE, 17, 37, 400, 17, hWnd, (HMENU)IDC_EDIT_SERVER_NAME, lpCreateStruct->hInstance, NULL);
	SendMessage(hTemp, EM_SETLIMITTEXT, EDIT_LIMIT, 0);

	CreateWindowEx(0, TEXT("Static"), TEXT("IPv4-адрес сервера"), WS_CHILD | WS_VISIBLE, 17, 57, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	CreateWindowEx(0, TEXT("SysIPAddress32"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE, 17, 77, 400, 20, hWnd, (HMENU)IDC_IP_SERVER_IP, lpCreateStruct->hInstance, NULL);

	CreateWindowEx(0, TEXT("Static"), TEXT("Адрес отправителя (логин)"), WS_CHILD | WS_VISIBLE, 17, 100, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	hTemp = CreateWindowEx(0, TEXT("Edit"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE, 17, 120, 400, 17, hWnd, (HMENU)IDC_EDIT_SENDER, lpCreateStruct->hInstance, NULL);
	SendMessage(hTemp, EM_SETLIMITTEXT, EDIT_LIMIT, 0);

	CreateWindowEx(0, TEXT("Static"), TEXT("Пароль"), WS_CHILD | WS_VISIBLE, 17, 140, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	hTemp = CreateWindowEx(0, TEXT("Edit"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE | ES_PASSWORD, 17, 160, 400, 17, hWnd, (HMENU)IDC_EDIT_PASSWORD, lpCreateStruct->hInstance, NULL);
	SendMessage(hTemp, EM_SETLIMITTEXT, EDIT_LIMIT, 0);

	CreateWindowEx(0, TEXT("Static"), TEXT("Получатели"), WS_CHILD | WS_VISIBLE, 17, 180, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	CreateWindowEx(0, TEXT("ListBox"), NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_HSCROLL | LBS_NOTIFY, 17, 200, 400, RECIPIENT_LIMIT * 17, hWnd, (HMENU)IDC_LISTBOX_RECIPIENTS, lpCreateStruct->hInstance, NULL);

	CreateWindowEx(0, TEXT("Static"), TEXT("Добавление получателя"), WS_CHILD | WS_VISIBLE, 17, 288, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	hTemp = CreateWindowEx(0, TEXT("Edit"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE, 17, 308, 400, 17, hWnd, (HMENU)IDC_EDIT_RECIPIENT, lpCreateStruct->hInstance, NULL);
	SendMessage(hTemp, EM_SETLIMITTEXT, EDIT_LIMIT, 0);

	CreateWindowEx(0, TEXT("Button"), TEXT("Добавить"), WS_CHILD | WS_VISIBLE, 17, 337, 175, 23, hWnd, (HMENU)IDC_BUTTON_ADD_RECIPIENT, lpCreateStruct->hInstance, NULL);
	CreateWindowEx(0, TEXT("Button"), TEXT("Удалить"), WS_CHILD | WS_VISIBLE, 242, 337, 175, 23, hWnd, (HMENU)IDC_BUTTON_DELETE_RECIPIENT, lpCreateStruct->hInstance, NULL);

	CreateWindowEx(0, TEXT("Static"), TEXT("Прикреплённые файлы"), WS_CHILD | WS_VISIBLE, 17, 362, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	CreateWindowEx(0, TEXT("ListBox"), NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_HSCROLL | LBS_NOTIFY | WS_HSCROLL, 17, 382, 400, FILE_LIMIT * 17, hWnd, (HMENU)IDC_LISTBOX_FILES, lpCreateStruct->hInstance, NULL);

	CreateWindowEx(0, TEXT("Button"), TEXT("Добавить"), WS_CHILD | WS_VISIBLE, 17, 473, 175, 23, hWnd, (HMENU)IDC_BUTTON_ADD_FILE, lpCreateStruct->hInstance, NULL);
	CreateWindowEx(0, TEXT("Button"), TEXT("Удалить"), WS_CHILD | WS_VISIBLE, 242, 473, 175, 23, hWnd, (HMENU)IDC_BUTTON_DELETE_FILE, lpCreateStruct->hInstance, NULL);

	CreateWindowEx(0, TEXT("Static"), TEXT("Тема письма"), WS_CHILD | WS_VISIBLE, 477, 17, 400, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	hTemp = CreateWindowEx(0, TEXT("Edit"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE, 477, 37, 400, 17, hWnd, (HMENU)IDC_EDIT_SUBJECT, lpCreateStruct->hInstance, NULL);
	SendMessage(hTemp, EM_SETLIMITTEXT, EDIT_LIMIT, 0);

	CreateWindowEx(0, TEXT("Static"), TEXT("Текст письма"), WS_CHILD | WS_VISIBLE, 477, 57, 600, 17, hWnd, NULL, lpCreateStruct->hInstance, NULL);
	hTemp = CreateWindowEx(0, TEXT("Edit"), TEXT_EMPTY, WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL, 477, 77, 600, 17 * 27, hWnd, (HMENU)IDC_EDIT_MESSAGE, lpCreateStruct->hInstance, NULL);
	SendMessage(hTemp, EM_SETLIMITTEXT, MESSAGE_LIMIT, 0);

	CreateWindowEx(0, TEXT("Button"), TEXT("Отправить"), WS_CHILD | WS_VISIBLE, 902, 560, 175, 23, hWnd, (HMENU)IDC_BUTTON_SEND, lpCreateStruct->hInstance, NULL);

	return TRUE;
}

/**
 * Обработка нажатия на элемент управления.
 * @param hWnd - дескриптор главного окна.
 * @param id - идентификатор элемента управления.
 * @param hwndCtl - дескриптор нажатого элемента управления.
 * @param codeNotify - код уведомления.
 */
VOID OnCommand(HWND hWnd, INT id, HWND hwndCtl, UINT codeNotify)
{
	switch (id)
	{
	case IDC_BUTTON_ADD_RECIPIENT:
	{
		TCHAR recipientName[EDIT_LIMIT];
		GetDlgItemText(MainWindow, IDC_EDIT_RECIPIENT, recipientName, EDIT_LIMIT);
		if (!_tcscmp(recipientName, TEXT_EMPTY))
		{
			return;
		}

		if (RecipientCount == RECIPIENT_LIMIT)
		{
			MessageBox(MainWindow, TEXT("Достигнуто максимальное кол-во получателей. "), TEXT("Ошибка"), MB_OK | MB_ICONERROR | MB_DEFBUTTON1);

			return;
		}

		HWND recipientsListBox = GetDlgItem(hWnd, IDC_LISTBOX_RECIPIENTS);
		SendMessage(recipientsListBox, LB_ADDSTRING, NULL, (LPARAM)recipientName);
		++RecipientCount;

		return;
	}

	case IDC_BUTTON_DELETE_RECIPIENT:
	{
		HWND recipientsListBox = GetDlgItem(hWnd, IDC_LISTBOX_RECIPIENTS);
		INT index = SendMessage(recipientsListBox, LB_GETCURSEL, 0, 0);
		if (index > -1)
		{
			--RecipientCount;
			SendMessage(recipientsListBox, LB_DELETESTRING, index, 0);
		}

		return;
	}

	case IDC_BUTTON_ADD_FILE:
	{
		if (FileCount == FILE_LIMIT)
		{
			MessageBox(MainWindow, TEXT("Достигнуто максимальное кол-во прикрепляемых файлов. "), TEXT("Ошибка"), MB_OK | MB_ICONERROR | MB_DEFBUTTON1);

			return;
		}

		TCHAR fileName[MAX_PATH];
		OPENFILENAME ofn;
		ZeroMemory(&ofn, sizeof(ofn));
		ofn.lStructSize = sizeof(ofn);
		ofn.hwndOwner = NULL;
		ofn.lpstrFile = fileName;
		ofn.lpstrFile[0] = '\0';
		ofn.nMaxFile = MAX_PATH;
		ofn.lpstrFilter = TEXT("Все\0*.*\0");
		ofn.nFilterIndex = 0;
		ofn.lpstrFileTitle = NULL;
		ofn.nMaxFileTitle = 0;
		ofn.lpstrInitialDir = NULL;
		ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
		if (GetOpenFileName(&ofn) == TRUE)
		{
			ifstream file(ofn.lpstrFile);
			file.seekg(0, std::ios::end);
			size_t size = file.tellg();
			if (size > FILE_SIZE_LIMIT * 1024 * 1024)
			{
				MessageBox(MainWindow, TEXT("Прикрепляемый файл слишком большой. "), TEXT("Ошибка"), MB_OK | MB_ICONERROR | MB_DEFBUTTON1);

				return;
			}

			HWND fileListBox = GetDlgItem(hWnd, IDC_LISTBOX_FILES);
			SendMessage(fileListBox, LB_ADDSTRING, NULL, (LPARAM)PathFindFileName(ofn.lpstrFile));
			string fileName(ofn.lpstrFile);
			FileNames.push_back(fileName);
			++FileCount;
		}


		return;
	}

	case IDC_BUTTON_DELETE_FILE:
	{
		HWND fileListBox = GetDlgItem(hWnd, IDC_LISTBOX_FILES);
		INT index = SendMessage(fileListBox, LB_GETCURSEL, 0, 0);
		if (index > -1)
		{
			--FileCount;
			SendMessage(fileListBox, LB_DELETESTRING, index, 0);
			FileNames.erase(FileNames.begin() + index);
		}

		return;
	}

	case IDC_BUTTON_SEND:
		OnSendMessage();

		return;
	}
}

/**
 * Обработка нажатия кнопки отправки письма.
 */
VOID OnSendMessage()
{ 
	TCHAR serverName[EDIT_LIMIT];
	GetDlgItemText(MainWindow, IDC_EDIT_SERVER_NAME, serverName, EDIT_LIMIT);
	
	string serverIP = GetServerIP();

	TCHAR sender[EDIT_LIMIT];
	GetDlgItemText(MainWindow, IDC_EDIT_SENDER, sender, EDIT_LIMIT);

	TCHAR password[EDIT_LIMIT];
	GetDlgItemText(MainWindow, IDC_EDIT_PASSWORD, password, EDIT_LIMIT);

	vector<string> recipients = GetRecipients();

	TCHAR subject[EDIT_LIMIT];
	GetDlgItemText(MainWindow, IDC_EDIT_SUBJECT, subject, EDIT_LIMIT);

	TCHAR message[MESSAGE_LIMIT];
	GetDlgItemText(MainWindow, IDC_EDIT_MESSAGE, message, MESSAGE_LIMIT);

	string res = SendEmail(serverName, serverIP, recipients, subject, sender, password, message, FileNames);
	if (!res.empty())
	{
		MessageBox(MainWindow, res.c_str(), TEXT("Ошибка"), MB_OK | MB_ICONERROR | MB_DEFBUTTON1);

		return;
	}

	MessageBox(MainWindow, TEXT("Отправка письма успешно завешена. "), TEXT("Отправка"), MB_OK | MB_ICONINFORMATION | MB_DEFBUTTON1);
}

/**
 * Получение IP-адреса SMTP-сервера, введённого пользователем.
 * @return IP-адрес SMTP-сервера вида ip1.ip2.ip3.ip4 .
 * @note Если пользователь не вввёл IP-адрес возвращается пустая строка.
 */
string GetServerIP()
{
	HWND IPEdit = GetDlgItem(MainWindow, IDC_IP_SERVER_IP);
	DWORD serverIP;
	SendMessage(IPEdit, IPM_GETADDRESS, 0, (LPARAM)&serverIP);
	if (serverIP == 0)
		return TEXT_EMPTY;

	int ip1 = FIRST_IPADDRESS(serverIP);
	int ip2 = SECOND_IPADDRESS(serverIP);
	int ip3 = THIRD_IPADDRESS(serverIP);
	int ip4 = FOURTH_IPADDRESS(serverIP);
	string ip = to_string(ip1) + "." + to_string(ip2) + "." + to_string(ip3) + "." + to_string(ip4);

	return ip;
}

/**
 * "Выборка" получателей письма, введённых пользователем.
 * @return получатели письма.
 */
vector<string> GetRecipients()
{
	HWND recipientsListBox = GetDlgItem(MainWindow, IDC_LISTBOX_RECIPIENTS);
	INT recipientsCount = SendMessage(recipientsListBox, LB_GETCOUNT, 0, 0);
	vector<string> recipients;
	for (INT i = 0; i < recipientsCount; ++i)
	{
		TCHAR recipient[EDIT_LIMIT];
		SendMessage(recipientsListBox, LB_GETTEXT, i, (LPARAM)recipient);
		string recipientString(recipient);
		recipients.push_back(recipientString);
	}

	return recipients;
}