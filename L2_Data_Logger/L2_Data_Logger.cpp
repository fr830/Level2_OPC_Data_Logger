// L2_Data_Logger.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "L2_Data_Logger.h"
#include "opccomn.h"
#include "opcda.h"
#include "OPCDataCallback.h"
#include <comdef.h>
#include "TagDefinition.h"
#include "TagDefinition_FS_PLC.h"
#include "TagDefinition_ULS_PLC.h"

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name
double m_PrevAsh, m_CurAsh, m_PrevAsh_10MIN, m_CurAsh_10MIN;
double m_PrevAsh_Y18, m_CurAsh_Y18, m_PrevAsh_10MIN_Y18, m_CurAsh_10MIN_Y18;

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

extern "C"
{
	const IID IID_IOPCServer = { 0x39c13a4d, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCServerPublicGroups = { 0x39c13a4e, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCBrowseServerAddressSpace = { 0x39c13a4f, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCGroupStateMgt = { 0x39c13a50, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCPublicGroupStateMgt = { 0x39c13a51, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCSyncIO = { 0x39c13a52, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCAsyncIO = { 0x39c13a53, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCItemMgt = { 0x39c13a54, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IEnumOPCItemAttributes = { 0x39c13a55, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
	const IID IID_IOPCDataCallback = { 0x39c13a70, 0x011e, 0x11d0, { 0x96, 0x75, 0x00, 0x20, 0xaf, 0xd8, 0xad, 0xb3 } };
}

UINT_PTR tim_ind = 0;
UINT_PTR tim_ind_10MIN = 0;

#define TID_1MIN   10
#define TID_10MIN  12

int no_of_tags_read_Y37_1_MIN =  no_of_tags_Y37_Analyser_1_MIN;
int no_of_tags_read_Y37_10_MIN = no_of_tags_Y37_Analyser_10_MIN;
int no_of_tags_read_Y18_1_MIN = no_of_tags_Y18_Analyser_1_MIN;

int no_of_tags_read_PLC_FS_ASYNCH = no_of_tags_FS_PLC_ASYNCH;
int no_of_tags_read_PLC_ULS_SYNCH = no_of_tags_ULS_PLC_SYNCH;

IUnknown* lpUnknown = NULL;
IOPCServer* m_OpcServer_Thermo = NULL;
IUnknown* pUnkGrp = NULL;

//// Interfaces for Y37 Item  //////////

///ForY37,  1 MIN ANALYSIS
IOPCItemMgt		*pOPCItemMgt_Read_Y37 = NULL;
OPCITEMDEF		*ItemArray_Read_Y37 = NULL;
OPCITEMRESULT	*AddResults_Read_Y37;
OPCHANDLE *ItemHandles_Read_Y37 = (OPCHANDLE*)CoTaskMemAlloc(no_of_tags_Y37_Analyser_1_MIN*sizeof(OPCHANDLE));
IOPCSyncIO		*pOPCSyncIO_read_Y37 = NULL;
OPCITEMSTATE* ItemState_Read_Y37 = (OPCITEMSTATE*)CoTaskMemAlloc(no_of_tags_Y37_Analyser_1_MIN*sizeof(OPCITEMSTATE));

void Analyzer_Db_Connection(DWORD, OPCITEMSTATE*); //Anlayzer database connection for 1 MIN
void Analyzer_Db_Connection_10MIN(DWORD, OPCITEMSTATE*); //Anlayzer database connection for 10 MIN

///For 10 MIN ANALYSIS
IOPCItemMgt		*pOPCItemMgt_Read_Y37_10MIN = NULL;
OPCITEMDEF		*ItemArray_Read_Y37_10MIN = NULL;
OPCITEMRESULT	*AddResults_Read_Y37_10MIN;
OPCHANDLE *ItemHandles_Read_Y37_10MIN = (OPCHANDLE*)CoTaskMemAlloc(no_of_tags_read_Y37_10_MIN*sizeof(OPCHANDLE));
IOPCSyncIO		*pOPCSyncIO_read_Y37_10MIN = NULL;
OPCITEMSTATE* ItemState_Read_Y37_10MIN = (OPCITEMSTATE*)CoTaskMemAlloc(no_of_tags_read_Y37_10_MIN*sizeof(OPCITEMSTATE));


///ForY18,  1 MIN ANALYSIS
IOPCItemMgt		*pOPCItemMgt_Read_Y18 = NULL;
OPCITEMDEF		*ItemArray_Read_Y18 = NULL;
OPCITEMRESULT	*AddResults_Read_Y18;
OPCHANDLE *ItemHandles_Read_Y18 = (OPCHANDLE*)CoTaskMemAlloc(no_of_tags_Y18_Analyser_1_MIN*sizeof(OPCHANDLE));
IOPCSyncIO		*pOPCSyncIO_read_Y18 = NULL;
OPCITEMSTATE* ItemState_Read_Y18 = (OPCITEMSTATE*)CoTaskMemAlloc(no_of_tags_Y18_Analyser_1_MIN*sizeof(OPCITEMSTATE));

void Analyzer_Db_Connection_Y18(DWORD, OPCITEMSTATE*); //Anlayzer database connection for 1 MIN
void Analyzer_COMMON_Db_Connection_Y18(_variant_t ); //Anlayzer database connection for 1 MIN

//void Analyzer_Db_Connection_10MIN(DWORD, OPCITEMSTATE*); //Anlayzer database connection for 10 MIN



//OPC Iterfaces for FS PLC
IOPCServer* m_OpcServer_FS_PLC = NULL;
IUnknown* pUnkGrp_FS_PLC = NULL;

//OPC Interfaces for FS PLC Asynch Read
IOPCItemMgt		*pOPCItemMgt_Read_FS_PLC_Asynch = NULL;
OPCITEMDEF		*ItemArray_Read_FS_PLC_Asynch = NULL;
OPCITEMRESULT	*AddResults_FS_PLC_Asynch;
OPCHANDLE *ItemHandles_Read_FS_PLC_Asynch = (OPCHANDLE*)CoTaskMemAlloc(no_of_tags_read_PLC_FS_ASYNCH*sizeof(OPCHANDLE));

IConnectionPointContainer* pCPC_PLC = NULL;
IConnectionPoint* pConnPt_PLC = NULL;
IOPCAsyncIO2* pOPCAsyncIO2_PLC = NULL;

//OPC Iterfaces for ULS PLC
//IOPCServer* m_OpcServer_FS_PLC = NULL;
IUnknown* pUnkGrp_ULS_PLC = NULL;

//OPC Interfaces for ULS PLC Asynch Read
IOPCItemMgt		*pOPCItemMgt_Read_ULS_PLC_Synch = NULL;
OPCITEMDEF		*ItemArray_Read_ULS_PLC_Synch = NULL;
OPCITEMRESULT	*AddResults_ULS_PLC_Synch;
OPCHANDLE *ItemHandles_Read_ULS_PLC_Synch = (OPCHANDLE*)CoTaskMemAlloc(no_of_tags_read_PLC_ULS_SYNCH*sizeof(OPCHANDLE));

IOPCSyncIO		*pOPCSyncIO_read_ULS_PLC_Synch = NULL;
OPCITEMSTATE* ItemState_Read_ULS_PLC_Synch = (OPCITEMSTATE*)CoTaskMemAlloc(no_of_tags_read_PLC_ULS_SYNCH*sizeof(OPCITEMSTATE));

void ULS_PLC_Db_Connection(DWORD, OPCITEMSTATE*); //Anlayzer database connection for ULS PLC

//IConnectionPointContainer* pCPC_PLC_ULS = NULL;
//IConnectionPoint* pConnPt_PLC_ULS = NULL;
//IOPCAsyncIO2* pOPCAsyncIO2_PLC_ULS = NULL;




int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPTSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_L2_DATA_LOGGER, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_L2_DATA_LOGGER));

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
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
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	//wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_L2_DATA_LOGGER));
	wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SAIL_ICON));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_L2_DATA_LOGGER);
	wcex.lpszClassName	= szWindowClass;
	//wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));
	wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SAIL_ICON));

	return RegisterClassEx(&wcex);
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
   HWND hWnd;

   hInst = hInstance; // Store instance handle in our global variable

  // hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);
   hWnd = CreateWindow(szWindowClass, _T("CHP_L2_DATA_LOGGER"), WS_MINIMIZE + WS_DISABLED, CW_USEDEFAULT, CW_USEDEFAULT, 0, NULL, NULL, NULL, hInstance, NULL);
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
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message)
	{
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		EndPaint(hWnd, &ps);
		break;
	case WM_CREATE:
		//Sleep(20000);
		CreateOPCConnection_Y37();  /// Procedure for creating OPC client and related other procedure ///
		
		CreateOPCConnection_PLC();  /// Procedure for creating OPC client and related other procedure ///
		PostMessage(hWnd, WM_USER + 18, 0, 0);
		PostMessage(hWnd, WM_USER + 20, 0, 0);
		break;
	case (WM_USER + 18) :
		Read_Y37_Data();     //Reading Y37 Elemental Analyzer data in 1 minute interval
		Read_Y18_Data_1MIN();  //Reading Y18 Elemental Analyzer data in 1 minute interval
		Silo_Calculation();		//Silo Stock Calulation in 1 minute interval
		Silo_Calculation_WF();  //Silo Stock Calculation for Weigh feeder Section for 1 minute interval
		//Ash_Calculation();	// Ash calculation for Y37 and Y38
		//if (!tim_ind)
		KillTimer(hWnd, TID_1MIN);
		
		tim_ind = SetTimer(hWnd, TID_1MIN, 60000, NULL);  //Starting 1 MIN TIMER
		break;
	case (WM_USER + 20) :
		//Read_Y37_Data_10MIN(); //Reading Y37 Elemental Analyzer data in 10 minute interval
		//if (!tim_ind_10MIN)
		KillTimer(hWnd, TID_10MIN);
		tim_ind_10MIN = SetTimer(hWnd, TID_10MIN, 600000, NULL); //Starting 10 MIN TIMER
		break;
	case WM_TIMER:
		switch (wParam)
		{
			case TID_1MIN:
				PostMessage(hWnd, WM_USER + 18, 0, 0);
			return 0;
			
			case TID_10MIN:
				PostMessage(hWnd, WM_USER + 20, 0, 0);
			return 0;			
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

BOOL CreateOPCConnection_Y37()  /// Procedure for creating OPC client and related other procedure ////////
{
	HRESULT hr;
	hr = CreateOPCServer_Thermo(); // Connecting the ThermoFisher server
	hr = CreateOPCGroup_Y37(); // Create the group
	hr = AddOPCItems_Read_Y37(); //Create OPC Items 1 MIN
	hr = AddOPCItems_Read_Y37_10MIN(); // Create OPC Items 10 MIN

	hr = AddOPCItems_Read_Y18_1MIN(); //Create OPC Items 1 MIN for Y18 analyser
	
	m_PrevAsh = 0;
	m_CurAsh = 0;
	m_PrevAsh_10MIN = 0;
	m_CurAsh_10MIN = 0;
	m_PrevAsh_Y18 = 0;
	m_CurAsh_Y18 = 0;

	return TRUE;
}

HRESULT CreateOPCServer_Thermo()  /// Procedure for creating OPC client and related other procedure ////////
{
	CLSID			clsidSimaticNETOPC;
	HRESULT			hr;

	hr = ::CoInitialize(NULL);

	///Retriving SIMATIC NET OPC Server CLSID
	if (CLSIDFromProgID(PROGID_SERVER_THERMO, &clsidSimaticNETOPC) != NOERROR)
	{
		MessageBox(NULL, _T("Unable to Get THERMO ECA OPC Server CLSID"), _T("Error"), MB_OK);
		return FALSE;
	}

	hr = CoCreateInstance(clsidSimaticNETOPC, NULL, CLSCTX_LOCAL_SERVER, IID_IUnknown, (LPVOID*)&lpUnknown);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("THERMO OPC CoCreateInstance Failed"), _T("Error"), MB_OK);
		return FALSE;
	}


	hr = lpUnknown->QueryInterface(IID_IOPCServer, (LPVOID*)&m_OpcServer_Thermo);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("THERMO OPC Server Creation Failed"), _T("Error"), MB_OK);
		m_OpcServer_Thermo->Release();
		return FALSE;
	}

	return hr;
}

HRESULT CreateOPCGroup_Y37()  /// Procedure for creating OPC client and related other procedure ////////
{
	HRESULT hr;
	DWORD dwRequestedUpdateRate;
	DWORD dwRevisedUpdateRate;
	OPCHANDLE hServerGroup;
	OPCHANDLE hClientGroup;
	LONG	lTimeBias = 240L;
	BOOL	GroupActiveState = TRUE;
	wchar_t	szwGroupName[80];
	wcscpy_s(szwGroupName, L"Y37_ANALYSER"); ///OPC Group name definition
	hClientGroup = 0x100;
	dwRequestedUpdateRate = 1000;

	// Add a group to the server
	hr = m_OpcServer_Thermo->AddGroup(szwGroupName, GroupActiveState, dwRequestedUpdateRate,
		hClientGroup, &lTimeBias, NULL, 0,
		&hServerGroup, &dwRevisedUpdateRate,
		IID_IUnknown, (LPUNKNOWN*)&pUnkGrp);
	
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Add OPC Group Y37-THERMO Failed"), _T("Error"), MB_OK);
		m_OpcServer_Thermo->Release();
		return FALSE;
	}
	return hr;
}

HRESULT AddOPCItems_Read_Y37()
{
	// *********** Code for Reading Scheduling PLC  Data   **********************////
	
	HRESULT hr;


	hr = pUnkGrp->QueryInterface(IID_IOPCItemMgt, (LPVOID*)&pOPCItemMgt_Read_Y37);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("IUnknown - Failed to get IOPCItemMgt_read pointer for Y37"), _T("Error"), MB_OK);
		CoUninitialize();
		exit(1);
		return FALSE;
	}

	ItemArray_Read_Y37 = (OPCITEMDEF*)CoTaskMemAlloc(no_of_tags_Y37_Analyser_1_MIN * sizeof(OPCITEMDEF));
	HRESULT			*AddErrors;

	for (int n = 0; n < no_of_tags_Y37_Analyser_1_MIN; n++)
	{
		ItemArray_Read_Y37[n].szItemID = (LPWSTR)CoTaskMemAlloc(sizeof(Tag_Y37_1_MIN[n]));
		ItemArray_Read_Y37[n].szAccessPath = L""; //szwAccessPath
		ItemArray_Read_Y37[n].szItemID = (LPWSTR)Tag_Y37_1_MIN[n];
		ItemArray_Read_Y37[n].bActive = TRUE;
		ItemArray_Read_Y37[n].hClient = 0x500 + n;
		ItemArray_Read_Y37[n].dwBlobSize = 0;
		ItemArray_Read_Y37[n].pBlob = NULL;
		ItemArray_Read_Y37[n].vtRequestedDataType = VT_EMPTY;
	}


	// Add the item to the group for read

	hr = pOPCItemMgt_Read_Y37->AddItems(no_of_tags_Y37_Analyser_1_MIN, ItemArray_Read_Y37, &AddResults_Read_Y37, &AddErrors);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("AddItem Failed for Y37 Analuser"), _T("Error"), MB_OK);
		CoTaskMemFree(AddErrors);
		CoTaskMemFree(AddResults_Read_Y37);
		pOPCItemMgt_Read_Y37->Release();
		exit(1);
	}
	if (hr == S_OK)
	{
		for (int n = 0; n < no_of_tags_Y37_Analyser_1_MIN; n++)
			ItemHandles_Read_Y37[n] = AddResults_Read_Y37[n].hServer;
	}

	CoTaskMemFree(AddErrors);
	
	return hr;
}

HRESULT AddOPCItems_Read_Y37_10MIN()
{
	// *********** Code for Reading Scheduling PLC  Data   **********************////

	HRESULT hr;


	hr = pUnkGrp->QueryInterface(IID_IOPCItemMgt, (LPVOID*)&pOPCItemMgt_Read_Y37_10MIN);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("IUnknown - Failed to get IOPCItemMgt_read_10MIN pointer for Y37 "), _T("Error"), MB_OK);
		CoUninitialize();
		exit(1);
		return FALSE;
	}

	ItemArray_Read_Y37_10MIN = (OPCITEMDEF*)CoTaskMemAlloc(no_of_tags_Y37_Analyser_10_MIN * sizeof(OPCITEMDEF));
	HRESULT			*AddErrors_10MIN;

	for (int n = 0; n < no_of_tags_Y37_Analyser_10_MIN; n++)
	{
		ItemArray_Read_Y37_10MIN[n].szItemID = (LPWSTR)CoTaskMemAlloc(sizeof(Tag_Y37_10_MIN[n]));
		ItemArray_Read_Y37_10MIN[n].szAccessPath = L""; //szwAccessPath
		ItemArray_Read_Y37_10MIN[n].szItemID = (LPWSTR)Tag_Y37_10_MIN[n];
		ItemArray_Read_Y37_10MIN[n].bActive = TRUE;
		ItemArray_Read_Y37_10MIN[n].hClient = 0x2000 + n;
		ItemArray_Read_Y37_10MIN[n].dwBlobSize = 0;
		ItemArray_Read_Y37_10MIN[n].pBlob = NULL;
		ItemArray_Read_Y37_10MIN[n].vtRequestedDataType = VT_EMPTY;
	}


	// Add the item to the group for read

	hr = pOPCItemMgt_Read_Y37_10MIN->AddItems(no_of_tags_Y37_Analyser_10_MIN, ItemArray_Read_Y37_10MIN, &AddResults_Read_Y37_10MIN, &AddErrors_10MIN);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("AddItem_10MIN Failed for Y37 Analuser"), _T("Error"), MB_OK);
		CoTaskMemFree(AddErrors_10MIN);
		CoTaskMemFree(AddResults_Read_Y37_10MIN);
		pOPCItemMgt_Read_Y37_10MIN->Release();
		exit(1);
	}
	if (hr == S_OK)
	{
		for (int n = 0; n < no_of_tags_Y37_Analyser_1_MIN; n++)
			ItemHandles_Read_Y37_10MIN[n] = AddResults_Read_Y37_10MIN[n].hServer;
	}

	CoTaskMemFree(AddErrors_10MIN);

	return hr;
}


HRESULT AddOPCItems_Read_Y18_1MIN()
{
	// *********** Code for Reading Scheduling PLC  Data   **********************////

	HRESULT hr;


	hr = pUnkGrp->QueryInterface(IID_IOPCItemMgt, (LPVOID*)&pOPCItemMgt_Read_Y18);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("IUnknown - Failed to get IOPCItemMgt_read_1MIN pointer for Y18 "), _T("Error"), MB_OK);
		CoUninitialize();
		exit(1);
		return FALSE;
	}

	ItemArray_Read_Y18 = (OPCITEMDEF*)CoTaskMemAlloc(no_of_tags_Y18_Analyser_1_MIN * sizeof(OPCITEMDEF));
	HRESULT			*AddErrors_Y18_1MIN;

	for (int n = 0; n < no_of_tags_Y18_Analyser_1_MIN; n++)
	{
		ItemArray_Read_Y18[n].szItemID = (LPWSTR)CoTaskMemAlloc(sizeof(Tag_Y18_1_MIN[n]));
		ItemArray_Read_Y18[n].szAccessPath = L""; //szwAccessPath
		ItemArray_Read_Y18[n].szItemID = (LPWSTR)Tag_Y18_1_MIN[n];
		ItemArray_Read_Y18[n].bActive = TRUE;
		ItemArray_Read_Y18[n].hClient = 0x7000 + n;
		ItemArray_Read_Y18[n].dwBlobSize = 0;
		ItemArray_Read_Y18[n].pBlob = NULL;
		ItemArray_Read_Y18[n].vtRequestedDataType = VT_EMPTY;
	}


	// Add the item to the group for read

	hr = pOPCItemMgt_Read_Y18->AddItems(no_of_tags_Y18_Analyser_1_MIN, ItemArray_Read_Y18, &AddResults_Read_Y18, &AddErrors_Y18_1MIN);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("AddItem_1MIN Failed for Y18 Analyser"), _T("Error"), MB_OK);
		CoTaskMemFree(AddErrors_Y18_1MIN);
		CoTaskMemFree(AddResults_Read_Y18);
		pOPCItemMgt_Read_Y18->Release();
		exit(1);
	}
	if (hr == S_OK)
	{
		for (int n = 0; n < no_of_tags_Y18_Analyser_1_MIN; n++)
			ItemHandles_Read_Y18[n] = AddResults_Read_Y18[n].hServer;
	}

	CoTaskMemFree(AddErrors_Y18_1MIN);
	
	return hr;
}



HRESULT Read_Y37_Data(void)
{
	HRESULT hr;
	HRESULT			*ReadErrors = NULL;
	DWORD			dwTransactionId = 0;

	//MessageBoxW(NULL, _T("Timer 1"), _T("Testing"), MB_OK);

	hr = pOPCItemMgt_Read_Y37->QueryInterface(IID_IOPCSyncIO, (void**)&pOPCSyncIO_read_Y37);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Unable to get IOPCSyncIO_read pointer for Y37 Analyser"), _T("Error"), MB_OK);
		CoUninitialize();
		//exit (1);
	}


	//Reading PLC data
	hr = pOPCSyncIO_read_Y37->Read(OPC_DS_DEVICE, no_of_tags_Y37_Analyser_1_MIN, ItemHandles_Read_Y37, &ItemState_Read_Y37, &ReadErrors);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Failed to Read Y37 Analyzer Data"), _T("Error"), MB_OK);
		//CoUninitialize();
		//exit (1);
	}

	//_variant_t var = ItemState_Read_Y37[25].vDataValue;
	//stat = ItemState_Read_Y37[0].vDataValue.iVal;
	m_CurAsh = ItemState_Read_Y37[0].vDataValue.dblVal;


	/*for (int n = 0; n < no_of_tags_Y37_Analyser_1_MIN; n++)
		double d = ItemState_Read_Y37[n].vDataValue.dblVal;*/

	//Reading Pushing /Charging Time to database
	if (hr == S_OK)
		if (m_PrevAsh != m_CurAsh) //If reading valid ash value
			Analyzer_Db_Connection(no_of_tags_Y37_Analyser_1_MIN, ItemState_Read_Y37); //Save data to database

	m_PrevAsh = m_CurAsh;

	for (int n = 0; n <no_of_tags_Y37_Analyser_1_MIN; n++)
	{
		ItemState_Read_Y37[n].vDataValue.vt = VT_NULL;
		ItemState_Read_Y37[n].ftTimeStamp.dwHighDateTime = NULL;
		ItemState_Read_Y37[n].ftTimeStamp.dwLowDateTime = NULL;
		ItemState_Read_Y37[n].hClient = NULL;
		ItemState_Read_Y37[n].wQuality = NULL;
		ItemState_Read_Y37[n].wReserved = NULL;
		VariantClear(&ItemState_Read_Y37[n].vDataValue);
	}

	pOPCSyncIO_read_Y37->Release();

	return hr;

}

HRESULT Read_Y37_Data_10MIN(void)
{
	HRESULT hr;
	HRESULT			*ReadErrors = NULL;
	DWORD			dwTransactionId = 0;
	//MessageBoxW(NULL, _T("Timer 10"), _T("Testing 10"), MB_OK);
	
	hr = pOPCItemMgt_Read_Y37_10MIN->QueryInterface(IID_IOPCSyncIO, (void**)&pOPCSyncIO_read_Y37_10MIN);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Unable to get IOPCSyncIO_read_10MIN pointer for Y37 Analyser"), _T("Error"), MB_OK);
		CoUninitialize();
		//exit (1);
	}


	//Reading PLC data
	hr = pOPCSyncIO_read_Y37_10MIN->Read(OPC_DS_DEVICE, no_of_tags_Y37_Analyser_10_MIN, ItemHandles_Read_Y37_10MIN, &ItemState_Read_Y37_10MIN, &ReadErrors);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Failed to Read Y37 Analyzer Data 10MIN"), _T("Error"), MB_OK);
		//CoUninitialize();
		//exit (1);
	}

	
	m_CurAsh_10MIN =  ItemState_Read_Y37_10MIN[0].vDataValue.dblVal;


	/*for (int n = 0; n < no_of_tags_Y37_Analyser_1_MIN; n++)
	double d = ItemState_Read_Y37[n].vDataValue.dblVal;*/

	//Writing data to database
	if (hr == S_OK)
		if (m_PrevAsh_10MIN != m_CurAsh_10MIN) //If reading valid ash value
			Analyzer_Db_Connection_10MIN(no_of_tags_Y37_Analyser_10_MIN, ItemState_Read_Y37_10MIN); //Save data to database


	m_PrevAsh_10MIN = m_CurAsh_10MIN;


	for (int n = 0; n <no_of_tags_Y37_Analyser_10_MIN; n++)
	{
		ItemState_Read_Y37_10MIN[n].vDataValue.vt = VT_NULL;
		ItemState_Read_Y37_10MIN[n].ftTimeStamp.dwHighDateTime = NULL;
		ItemState_Read_Y37_10MIN[n].ftTimeStamp.dwLowDateTime = NULL;
		ItemState_Read_Y37_10MIN[n].hClient = NULL;
		ItemState_Read_Y37_10MIN[n].wQuality = NULL;
		ItemState_Read_Y37_10MIN[n].wReserved = NULL;
		VariantClear(&ItemState_Read_Y37_10MIN[n].vDataValue);
	}

	pOPCSyncIO_read_Y37_10MIN->Release();

	return hr;
}

HRESULT Read_Y18_Data_1MIN(void)
{
	HRESULT hr;
	HRESULT			*ReadErrors = NULL;
	DWORD			dwTransactionId = 0;

	//MessageBoxW(NULL, _T("Timer 1"), _T("Testing"), MB_OK);

	hr = pOPCItemMgt_Read_Y18->QueryInterface(IID_IOPCSyncIO, (void**)&pOPCSyncIO_read_Y18);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Unable to get IOPCSyncIO_read pointer for Y18 Analyser 1 MIN"), _T("Error"), MB_OK);
		CoUninitialize();
		//exit (1);
	}


	//Reading PLC data
	hr = pOPCSyncIO_read_Y18->Read(OPC_DS_DEVICE, no_of_tags_Y18_Analyser_1_MIN, ItemHandles_Read_Y18, &ItemState_Read_Y18, &ReadErrors);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Failed to Read Y18 1MIN Analyzer Data"), _T("Error"), MB_OK);
		//CoUninitialize();
		//exit (1);
	}

	//_variant_t var = ItemState_Read_Y37[25].vDataValue;
	//stat = ItemState_Read_Y37[0].vDataValue.iVal;
	m_CurAsh_Y18 = ItemState_Read_Y18[0].vDataValue.dblVal;
	
	if (hr == S_OK)
	{
		{
			if (m_PrevAsh_Y18 != m_CurAsh_Y18) //If reading valid ash value
				Analyzer_Db_Connection_Y18(no_of_tags_Y18_Analyser_1_MIN, ItemState_Read_Y18); //Save data to database
		}
		
		Analyzer_COMMON_Db_Connection_Y18(ItemState_Read_Y18[3].vDataValue);
	}


	m_PrevAsh_Y18 = m_CurAsh_Y18;

	for (int n = 0; n <no_of_tags_Y18_Analyser_1_MIN; n++)
	{
		ItemState_Read_Y18[n].vDataValue.vt = VT_NULL;
		ItemState_Read_Y18[n].ftTimeStamp.dwHighDateTime = NULL;
		ItemState_Read_Y18[n].ftTimeStamp.dwLowDateTime = NULL;
		ItemState_Read_Y18[n].hClient = NULL;
		ItemState_Read_Y18[n].wQuality = NULL;
		ItemState_Read_Y18[n].wReserved = NULL;
		VariantClear(&ItemState_Read_Y18[n].vDataValue);
	}

	pOPCSyncIO_read_Y18->Release();

	return hr;

}

void Analyzer_Db_Connection(DWORD Count, OPCITEMSTATE* pValues)
{
	CString m_strConnection;
	CString m_strCmdText;
	_RecordsetPtr m_pRs;

	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Analyzer_Db;Trusted=True;");

	COleDateTime time_stamp;
	_variant_t var, var1;

	
	time_stamp = pValues[0].ftTimeStamp;
	var = time_stamp;

	COleDateTime curtime = COleDateTime::GetCurrentTime();
	var1 = curtime;
	
	//////////////////Retriving data for Coal Spillage //////////////
	m_strCmdText = _T("Select * from dbo.Y37_Analyzer");
	m_pRs.CreateInstance(_uuidof(Recordset));
	m_pRs->CursorLocation = adUseServer;
	m_pRs->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	m_pRs->AddNew();

	m_pRs->Fields->GetItem("Computer_Time")->Value = var1;
	m_pRs->Fields->GetItem("Time_Stamp")->Value = var;
	m_pRs->Fields->GetItem("ASH_AR_1_MIN")->Value = pValues[0].vDataValue;
	m_pRs->Fields->GetItem("MOIST_AR_1_MIN")->Value = pValues[1].vDataValue;
	m_pRs->Fields->GetItem("CV_AR_1_MIN")->Value = pValues[2].vDataValue;
	m_pRs->Fields->GetItem("TPH_AR_1_MIN")->Value = pValues[3].vDataValue;
	m_pRs->Fields->GetItem("SIO2_AR_1_MIN")->Value = pValues[4].vDataValue;
	m_pRs->Fields->GetItem("AL2O3_AR_1_MIN")->Value = pValues[5].vDataValue;
	m_pRs->Fields->GetItem("FE2O3_AR_1_MIN")->Value = pValues[6].vDataValue;
	m_pRs->Fields->GetItem("CAO_AR_1_MIN")->Value = pValues[7].vDataValue;
	m_pRs->Fields->GetItem("NA2O_AR_1_MIN")->Value = pValues[8].vDataValue;
	m_pRs->Fields->GetItem("K2O_AR_1_MIN")->Value = pValues[9].vDataValue;
	m_pRs->Fields->GetItem("TIO2_AR_1_MIN")->Value = pValues[10].vDataValue;
	m_pRs->Fields->GetItem("N_AR_1_MIN")->Value = pValues[11].vDataValue;
	m_pRs->Fields->GetItem("CL_AR_1_MIN")->Value = pValues[12].vDataValue;
	m_pRs->Fields->GetItem("S_AR_1_MIN")->Value = pValues[13].vDataValue;

	m_pRs->Fields->GetItem("ASH_DB_1_MIN")->Value = pValues[14].vDataValue;
	m_pRs->Fields->GetItem("SIO2_DB_1_MIN")->Value = pValues[15].vDataValue;
	m_pRs->Fields->GetItem("AL2O3_DB_1_MIN")->Value = pValues[16].vDataValue;
	m_pRs->Fields->GetItem("FE2O3_DB_1_MIN")->Value = pValues[17].vDataValue;
	m_pRs->Fields->GetItem("CAO_DB_1_MIN")->Value = pValues[18].vDataValue;
	m_pRs->Fields->GetItem("NA2O_DB_1_MIN")->Value = pValues[19].vDataValue;
	m_pRs->Fields->GetItem("K2O_DB_1_MIN")->Value = pValues[20].vDataValue;
	m_pRs->Fields->GetItem("TIO2_DB_1_MIN")->Value = pValues[21].vDataValue;
	m_pRs->Fields->GetItem("N_DB_1_MIN")->Value = pValues[22].vDataValue;
	m_pRs->Fields->GetItem("CL_DB_1_MIN")->Value = pValues[23].vDataValue;
	m_pRs->Fields->GetItem("S_DB_1_MIN")->Value = pValues[24].vDataValue;

	m_pRs->Update();
	



	m_pRs->Close();
}


void Analyzer_Db_Connection_10MIN(DWORD Count, OPCITEMSTATE* pValues)
{
	CString m_strConnection;
	CString m_strCmdText;
	_RecordsetPtr m_pRs;

	
	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Analyzer_Db;Trusted=True;");

	COleDateTime time_stamp;
	_variant_t var, var1;

	time_stamp = pValues[0].ftTimeStamp;
	var = time_stamp;

	COleDateTime curtime = COleDateTime::GetCurrentTime();
	var1 = curtime;

	//////////////////Retriving data for Coal Spillage //////////////
	m_strCmdText = _T("Select * from dbo.Y37_Analyzer_10_MIN");
	m_pRs.CreateInstance(_uuidof(Recordset));
	m_pRs->CursorLocation = adUseServer;
	m_pRs->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	m_pRs->AddNew();

	m_pRs->Fields->GetItem("Computer_Time")->Value = var1;
	m_pRs->Fields->GetItem("Time_Stamp")->Value = var;
	m_pRs->Fields->GetItem("ASH_AR_10_MIN")->Value = pValues[0].vDataValue;
	m_pRs->Fields->GetItem("MOIST_AR_10_MIN")->Value = pValues[1].vDataValue;
	m_pRs->Fields->GetItem("CV_AR_10_MIN")->Value = pValues[2].vDataValue;
	m_pRs->Fields->GetItem("TPH_AR_10_MIN")->Value = pValues[3].vDataValue;
	m_pRs->Fields->GetItem("SIO2_AR_10_MIN")->Value = pValues[4].vDataValue;
	m_pRs->Fields->GetItem("AL2O3_AR_10_MIN")->Value = pValues[5].vDataValue;
	m_pRs->Fields->GetItem("FE2O3_AR_10_MIN")->Value = pValues[6].vDataValue;
	m_pRs->Fields->GetItem("CAO_AR_10_MIN")->Value = pValues[7].vDataValue;
	m_pRs->Fields->GetItem("NA2O_AR_10_MIN")->Value = pValues[8].vDataValue;
	m_pRs->Fields->GetItem("K2O_AR_10_MIN")->Value = pValues[9].vDataValue;
	m_pRs->Fields->GetItem("TIO2_AR_10_MIN")->Value = pValues[10].vDataValue;
	m_pRs->Fields->GetItem("N_AR_10_MIN")->Value = pValues[11].vDataValue;
	m_pRs->Fields->GetItem("CL_AR_10_MIN")->Value = pValues[12].vDataValue;
	m_pRs->Fields->GetItem("S_AR_10_MIN")->Value = pValues[13].vDataValue;

	m_pRs->Fields->GetItem("ASH_DB_10_MIN")->Value = pValues[14].vDataValue;
	m_pRs->Fields->GetItem("SIO2_DB_10_MIN")->Value = pValues[15].vDataValue;
	m_pRs->Fields->GetItem("AL2O3_DB_10_MIN")->Value = pValues[16].vDataValue;
	m_pRs->Fields->GetItem("FE2O3_DB_10_MIN")->Value = pValues[17].vDataValue;
	m_pRs->Fields->GetItem("CAO_DB_10_MIN")->Value = pValues[18].vDataValue;
	m_pRs->Fields->GetItem("NA2O_DB_10_MIN")->Value = pValues[19].vDataValue;
	m_pRs->Fields->GetItem("K2O_DB_10_MIN")->Value = pValues[20].vDataValue;
	m_pRs->Fields->GetItem("TIO2_DB_10_MIN")->Value = pValues[21].vDataValue;
	m_pRs->Fields->GetItem("N_DB_10_MIN")->Value = pValues[22].vDataValue;
	m_pRs->Fields->GetItem("CL_DB_10_MIN")->Value = pValues[23].vDataValue;
	m_pRs->Fields->GetItem("S_DB_10_MIN")->Value = pValues[24].vDataValue;

	m_pRs->Update();




	m_pRs->Close();
}

void Analyzer_Db_Connection_Y18(DWORD Count, OPCITEMSTATE* pValues)
{
	CString m_strConnection;
	CString m_strCmdText;
	_RecordsetPtr m_pRs;

	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Analyzer_Db;Trusted=True;");

	COleDateTime time_stamp;
	_variant_t var, var1;


	time_stamp = pValues[0].ftTimeStamp;
	var = time_stamp;

	COleDateTime curtime = COleDateTime::GetCurrentTime();
	var1 = curtime;

	//////////////////Retriving data for Coal Spillage //////////////
	m_strCmdText = _T("Select * from dbo.Y18_Analyzer_1_MIN");
	m_pRs.CreateInstance(_uuidof(Recordset));
	m_pRs->CursorLocation = adUseServer;
	m_pRs->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	m_pRs->AddNew();

	m_pRs->Fields->GetItem("Computer_Time")->Value = var1;
	m_pRs->Fields->GetItem("Time_Stamp")->Value = var;
	m_pRs->Fields->GetItem("ASH_AR_1_MIN")->Value = pValues[0].vDataValue;
	m_pRs->Fields->GetItem("MOIST_AR_1_MIN")->Value = pValues[1].vDataValue;
	m_pRs->Fields->GetItem("CV_AR_1_MIN")->Value = pValues[2].vDataValue;
	m_pRs->Fields->GetItem("TPH_AR_1_MIN")->Value = pValues[3].vDataValue;
	m_pRs->Fields->GetItem("SIO2_AR_1_MIN")->Value = pValues[4].vDataValue;
	m_pRs->Fields->GetItem("AL2O3_AR_1_MIN")->Value = pValues[5].vDataValue;
	m_pRs->Fields->GetItem("FE2O3_AR_1_MIN")->Value = pValues[6].vDataValue;
	m_pRs->Fields->GetItem("CAO_AR_1_MIN")->Value = pValues[7].vDataValue;
	m_pRs->Fields->GetItem("NA2O_AR_1_MIN")->Value = pValues[8].vDataValue;
	m_pRs->Fields->GetItem("K2O_AR_1_MIN")->Value = pValues[9].vDataValue;
	m_pRs->Fields->GetItem("TIO2_AR_1_MIN")->Value = pValues[10].vDataValue;
	m_pRs->Fields->GetItem("N_AR_1_MIN")->Value = pValues[11].vDataValue;
	m_pRs->Fields->GetItem("CL_AR_1_MIN")->Value = pValues[12].vDataValue;
	m_pRs->Fields->GetItem("S_AR_1_MIN")->Value = pValues[13].vDataValue;

	m_pRs->Fields->GetItem("ASH_DB_1_MIN")->Value = pValues[14].vDataValue;
	m_pRs->Fields->GetItem("SIO2_DB_1_MIN")->Value = pValues[15].vDataValue;
	m_pRs->Fields->GetItem("AL2O3_DB_1_MIN")->Value = pValues[16].vDataValue;
	m_pRs->Fields->GetItem("FE2O3_DB_1_MIN")->Value = pValues[17].vDataValue;
	m_pRs->Fields->GetItem("CAO_DB_1_MIN")->Value = pValues[18].vDataValue;
	m_pRs->Fields->GetItem("NA2O_DB_1_MIN")->Value = pValues[19].vDataValue;
	m_pRs->Fields->GetItem("K2O_DB_1_MIN")->Value = pValues[20].vDataValue;
	m_pRs->Fields->GetItem("TIO2_DB_1_MIN")->Value = pValues[21].vDataValue;
	m_pRs->Fields->GetItem("N_DB_1_MIN")->Value = pValues[22].vDataValue;
	m_pRs->Fields->GetItem("CL_DB_1_MIN")->Value = pValues[23].vDataValue;
	m_pRs->Fields->GetItem("S_DB_1_MIN")->Value = pValues[24].vDataValue;

	m_pRs->Update();




	m_pRs->Close();
}

void Analyzer_COMMON_Db_Connection_Y18(_variant_t var_y18_tph)
{
	CString m_strConnection;
	CString m_strCmdText;
	_RecordsetPtr m_pRs;

	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Analyzer_Db;Trusted=True;");


	//////////////////Opening Database //////////////
	m_strCmdText = _T("Select * from dbo.Y18_COMMON");
	m_pRs.CreateInstance(_uuidof(Recordset));
	m_pRs->CursorLocation = adUseServer;
	m_pRs->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	if (var_y18_tph.vt != VT_NULL)
		m_pRs->Update(("Y18_TPH"), var_y18_tph);


	m_pRs->Close();
}

BOOL CreateOPCConnection_PLC()  /// Procedure for creating OPC client and related other procedure for PLC ////////
{
	HRESULT hr;
	hr = CreateOPCServer_PLC(); // Connecting the Schneider PLC OPC server
	
	hr = CreateOPCGroup_PLC_FS(); // Create the FS PLC group
	hr = AddOPCItems_Read_Asynch_PLC_FS(); //Create OPC Items PLC Asynch Read
	hr = ReadAsynchData_FS_PLC(); //Reading Asynchrous PLC FS

	hr = CreateOPCGroup_PLC_ULS(); // Create the ULS PLC group
	hr = AddOPCItems_Read_Synch_PLC_ULS(); //Create OPC Items PLC Synch Read
//	hr = ReadSynchData_ULS_PLC();//Reading Synchrous PLC ULS
//	hr = ReadAsynchData_ULS_PLC(); //Reading Asynchrous PLC ULS

	return TRUE;
}

HRESULT CreateOPCServer_PLC()  /// Procedure for creating OPC client and related other procedure for PLC ////////
{
	CLSID			clsidSimaticNETOPC;
	HRESULT			hr;

	hr = ::CoInitialize(NULL);

	///Retriving SIMATIC NET OPC Server CLSID
	if (CLSIDFromProgID(PROGID_SERVER_PLC, &clsidSimaticNETOPC) != NOERROR)
	{
		MessageBox(NULL, _T("Unable to Get Schneider PLC OPC Server CLSID"), _T("Error"), MB_OK);
		return FALSE;
	}

	hr = CoCreateInstance(clsidSimaticNETOPC, NULL, CLSCTX_LOCAL_SERVER, IID_IUnknown, (LPVOID*)&lpUnknown);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Schneider PLC OPC CoCreateInstance Failed"), _T("Error"), MB_OK);
		return FALSE;
	}

	hr = lpUnknown->QueryInterface(IID_IOPCServer, (LPVOID*)&m_OpcServer_FS_PLC);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Schneider PLC OPC Server Creation Failed"), _T("Error"), MB_OK);
		m_OpcServer_Thermo->Release();
		return FALSE;
	}

	return hr;
}

HRESULT CreateOPCGroup_PLC_FS()  /// Creating OPC Group for FS PLC////////
{
	HRESULT hr;
	DWORD dwRequestedUpdateRate;
	DWORD dwRevisedUpdateRate;
	OPCHANDLE hServerGroup;
	OPCHANDLE hClientGroup;
	LONG	lTimeBias = 240L;
	BOOL	GroupActiveState = TRUE;
	wchar_t	szwGroupName[80];
	wcscpy_s(szwGroupName, L"FS_PLC_GR"); ///OPC Group name definition
	hClientGroup = 0x100;
	dwRequestedUpdateRate = 2000;

	// Add a group to the server
	hr = m_OpcServer_FS_PLC->AddGroup(szwGroupName, GroupActiveState, dwRequestedUpdateRate,
		hClientGroup, &lTimeBias, NULL, 0,
		&hServerGroup, &dwRevisedUpdateRate,
		IID_IUnknown, (LPUNKNOWN*)&pUnkGrp_FS_PLC);

	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Add OPC FS PLC Failed"), _T("Error"), MB_OK);
		m_OpcServer_FS_PLC->Release();
		return FALSE;
	}
	return hr;
}

HRESULT CreateOPCGroup_PLC_ULS()  /// Creating OPC Group for ULS PLC////////
{
	HRESULT hr;
	DWORD dwRequestedUpdateRate;
	DWORD dwRevisedUpdateRate;
	OPCHANDLE hServerGroup;
	OPCHANDLE hClientGroup;
	LONG	lTimeBias = 240L;
	BOOL	GroupActiveState = TRUE;
	wchar_t	szwGroupName[80];
	wcscpy_s(szwGroupName, L"ULS_PLC_GR"); ///OPC Group name definition
	hClientGroup = 0x200;
	dwRequestedUpdateRate = 2000;

	// Add a group to the server
	hr = m_OpcServer_FS_PLC->AddGroup(szwGroupName, GroupActiveState, dwRequestedUpdateRate,
		hClientGroup, &lTimeBias, NULL, 0,
		&hServerGroup, &dwRevisedUpdateRate,
		IID_IUnknown, (LPUNKNOWN*)&pUnkGrp_ULS_PLC);

	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Add OPC ULS PLC Failed"), _T("Error"), MB_OK);
		m_OpcServer_FS_PLC->Release();
		return FALSE;
	}
	return hr;
}



HRESULT AddOPCItems_Read_Asynch_PLC_FS()
{
	// *********** Code for Reading Scheduling PLC  Data   **********************////

	HRESULT hr;


	hr = pUnkGrp_FS_PLC->QueryInterface(IID_IOPCItemMgt, (LPVOID*)&pOPCItemMgt_Read_FS_PLC_Asynch);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("IUnknown - Failed to get IOPCItemMgt_read_FS_PLC_Asynch pointer for FS PLC "), _T("Error"), MB_OK);
		CoUninitialize();
		exit(1);
		return FALSE;
	}

	ItemArray_Read_FS_PLC_Asynch = (OPCITEMDEF*)CoTaskMemAlloc(no_of_tags_read_PLC_FS_ASYNCH * sizeof(OPCITEMDEF));
	HRESULT			*AddErrors_PLC_FS_Asynch;

	for (int n = 0; n < no_of_tags_read_PLC_FS_ASYNCH; n++)
	{
		ItemArray_Read_FS_PLC_Asynch[n].szItemID = (LPWSTR)CoTaskMemAlloc(sizeof(Tag_FS_PLC_ASYNCH[n]));
		ItemArray_Read_FS_PLC_Asynch[n].szAccessPath = L"CHP_FS"; //szwAccessPath
		ItemArray_Read_FS_PLC_Asynch[n].szItemID = (LPWSTR)Tag_FS_PLC_ASYNCH[n];
		ItemArray_Read_FS_PLC_Asynch[n].bActive = TRUE;
		ItemArray_Read_FS_PLC_Asynch[n].hClient = 0x3000 + n;
		ItemArray_Read_FS_PLC_Asynch[n].dwBlobSize = 0;
		ItemArray_Read_FS_PLC_Asynch[n].pBlob = NULL;
		ItemArray_Read_FS_PLC_Asynch[n].vtRequestedDataType = VT_EMPTY;
	}


	// Add the item to the group for read

	hr = pOPCItemMgt_Read_FS_PLC_Asynch->AddItems(no_of_tags_read_PLC_FS_ASYNCH, ItemArray_Read_FS_PLC_Asynch, &AddResults_FS_PLC_Asynch, &AddErrors_PLC_FS_Asynch);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("AddItem_failed for FS PLC Asynch Read"), _T("Error"), MB_OK);
		CoTaskMemFree(AddErrors_PLC_FS_Asynch);
		CoTaskMemFree(AddResults_FS_PLC_Asynch);
		pOPCItemMgt_Read_FS_PLC_Asynch->Release();
		exit(1);
	}
	if (hr == S_OK)
	{
		for (int n = 0; n < no_of_tags_read_PLC_FS_ASYNCH; n++)
			ItemHandles_Read_FS_PLC_Asynch[n] = AddResults_FS_PLC_Asynch[n].hServer;
	}

	CoTaskMemFree(AddErrors_PLC_FS_Asynch);

	return hr;
}

HRESULT AddOPCItems_Read_Synch_PLC_ULS()
{
	// *********** Code for Reading Scheduling PLC  Data   **********************////

	HRESULT hr;


	hr = pUnkGrp_ULS_PLC->QueryInterface(IID_IOPCItemMgt, (LPVOID*)&pOPCItemMgt_Read_ULS_PLC_Synch);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("IUnknown - Failed to get IOPCItemMgt_read_ULS_PLC_Asynch pointer for ULS PLC "), _T("Error"), MB_OK);
		CoUninitialize();
		exit(1);
		return FALSE;
	}

	ItemArray_Read_ULS_PLC_Synch = (OPCITEMDEF*)CoTaskMemAlloc(no_of_tags_read_PLC_ULS_SYNCH * sizeof(OPCITEMDEF));
	HRESULT			*AddErrors_PLC_ULS_Synch;

	for (int n = 0; n < no_of_tags_read_PLC_ULS_SYNCH; n++)
	{
		ItemArray_Read_ULS_PLC_Synch[n].szItemID = (LPWSTR)CoTaskMemAlloc(sizeof(Tag_ULS_PLC_SYNCH[n]));
		ItemArray_Read_ULS_PLC_Synch[n].szAccessPath = L"CHP_ULS"; //szwAccessPath
		ItemArray_Read_ULS_PLC_Synch[n].szItemID = (LPWSTR)Tag_ULS_PLC_SYNCH[n];
		ItemArray_Read_ULS_PLC_Synch[n].bActive = TRUE;
		ItemArray_Read_ULS_PLC_Synch[n].hClient = 0x5000 + n;
		ItemArray_Read_ULS_PLC_Synch[n].dwBlobSize = 0;
		ItemArray_Read_ULS_PLC_Synch[n].pBlob = NULL;
		ItemArray_Read_ULS_PLC_Synch[n].vtRequestedDataType = VT_EMPTY;
	}


	// Add the item to the group for read

	hr = pOPCItemMgt_Read_ULS_PLC_Synch->AddItems(no_of_tags_read_PLC_ULS_SYNCH, ItemArray_Read_ULS_PLC_Synch, &AddResults_ULS_PLC_Synch, &AddErrors_PLC_ULS_Synch);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("AddItem_failed for ULS PLC Synch Read"), _T("Error"), MB_OK);
		CoTaskMemFree(AddErrors_PLC_ULS_Synch);
		CoTaskMemFree(AddResults_ULS_PLC_Synch);
		pOPCItemMgt_Read_ULS_PLC_Synch->Release();
		exit(1);
	}
	if (hr == S_OK)
	{
		for (int n = 0; n < no_of_tags_read_PLC_ULS_SYNCH; n++)
			ItemHandles_Read_ULS_PLC_Synch[n] = AddResults_ULS_PLC_Synch[n].hServer;
	}

	
	CoTaskMemFree(AddErrors_PLC_ULS_Synch);

	return hr;
}

HRESULT ReadAsynchData_FS_PLC()
{
	HRESULT hr;
	// Get pointer the Connection Point Container
	hr = pUnkGrp_FS_PLC->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC_PLC);
	if (FAILED(hr))
	{
		pUnkGrp_FS_PLC->Release();
		m_OpcServer_FS_PLC->Release();
		CoUninitialize();
		exit(1);
	}

	// Get pointer to the Data Callback connection point	
	hr = pCPC_PLC->FindConnectionPoint(IID_IOPCDataCallback, (IConnectionPoint**)&pConnPt_PLC);
	if (FAILED(hr)){
		m_OpcServer_FS_PLC->Release();
		CoUninitialize();
		exit(1);
	}

	// Don't need this interface anymore. Get rid of it.
	pCPC_PLC->Release();

	// Create a sink object.
	COPCDataCallback*  pOPCDataCallback = NULL;
	//COPCDataCallback*  pOPCDataCallback = (COPCDataCallback*)CoTaskMemAlloc(sizeof(COPCDataCallback));
	pOPCDataCallback = new COPCDataCallback;



	DWORD dwCookie;
	// Start the advise. Pass the address of the sink to the OPC server.
	hr = pConnPt_PLC->Advise(pOPCDataCallback, &dwCookie);
	if (FAILED(hr))
	{
		pConnPt_PLC->Release();
		m_OpcServer_FS_PLC->Release();
		CoUninitialize();
		exit(1);
	}

	return hr;

}

/*HRESULT ReadAsynchData_ULS_PLC()
{
	HRESULT hr;
	// Get pointer the Connection Point Container
	hr = pUnkGrp_ULS_PLC->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC_PLC_ULS);
	if (FAILED(hr))
	{
		pUnkGrp_ULS_PLC->Release();
		m_OpcServer_FS_PLC->Release();
		CoUninitialize();
		exit(1);
	}

	// Get pointer to the Data Callback connection point	
	hr = pCPC_PLC_ULS->FindConnectionPoint(IID_IOPCDataCallback, (IConnectionPoint**)&pConnPt_PLC_ULS);
	if (FAILED(hr)){
		m_OpcServer_FS_PLC->Release();
		CoUninitialize();
		exit(1);
	}

	// Don't need this interface anymore. Get rid of it.
	pCPC_PLC_ULS->Release();

	// Create a sink object.
	COPCDataCallback*  pOPCDataCallback_ULS = NULL;
	//COPCDataCallback*  pOPCDataCallback = (COPCDataCallback*)CoTaskMemAlloc(sizeof(COPCDataCallback));
	pOPCDataCallback_ULS = new COPCDataCallback;

	DWORD dwCookie;
	// Start the advise. Pass the address of the sink to the OPC server.
	hr = pConnPt_PLC_ULS->Advise(pOPCDataCallback_ULS, &dwCookie);
	if (FAILED(hr))
	{
		pConnPt_PLC_ULS->Release();
		m_OpcServer_FS_PLC->Release();
		CoUninitialize();
		exit(1);
	}

	return hr;

}*/

HRESULT ReadSynchData_ULS_PLC(void)
{
	HRESULT hr;
	
	HRESULT			*ReadErrors = NULL;
	DWORD			dwTransactionId = 0;

	
	hr = pOPCItemMgt_Read_ULS_PLC_Synch->QueryInterface(IID_IOPCSyncIO, (void**)&pOPCSyncIO_read_ULS_PLC_Synch);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Unable to get IOPCSyncIO_read pointer for ULS PLC Synchrounous Read"), _T("Error"), MB_OK);
		CoUninitialize();
		//exit (1);
	}


	//Reading PLC data
	hr = pOPCSyncIO_read_ULS_PLC_Synch->Read(OPC_DS_DEVICE, no_of_tags_read_PLC_ULS_SYNCH, ItemHandles_Read_ULS_PLC_Synch, &ItemState_Read_ULS_PLC_Synch, &ReadErrors);
	if (hr != S_OK)
	{
		MessageBox(NULL, _T("Failed to Read ULS PLC Synchronous Data"), _T("Error"), MB_OK);
		//CoUninitialize();
		//exit (1);
	}

	if (hr == S_OK)
		ULS_PLC_Db_Connection(no_of_tags_read_PLC_ULS_SYNCH, ItemState_Read_ULS_PLC_Synch);  ///Saving Data to database

	for (int n = 0; n <no_of_tags_read_PLC_ULS_SYNCH; n++)
	{
		ItemState_Read_ULS_PLC_Synch[n].vDataValue.vt = VT_NULL;
		ItemState_Read_ULS_PLC_Synch[n].ftTimeStamp.dwHighDateTime = NULL;
		ItemState_Read_ULS_PLC_Synch[n].ftTimeStamp.dwLowDateTime = NULL;
		ItemState_Read_ULS_PLC_Synch[n].hClient = NULL;
		ItemState_Read_ULS_PLC_Synch[n].wQuality = NULL;
		ItemState_Read_ULS_PLC_Synch[n].wReserved = NULL;
		VariantClear(&ItemState_Read_ULS_PLC_Synch[n].vDataValue);
	}

	pOPCSyncIO_read_ULS_PLC_Synch->Release();

	return hr;
}

void ULS_PLC_Db_Connection(DWORD Count, OPCITEMSTATE* pValues)
{
	///////////OPENING SILO DETAILS DATABASE ///////////////////
	CString m_strConnection_silodb;
	CString m_strCmdText_silodb;
	_RecordsetPtr m_pRs_silodb;


	m_strConnection_silodb = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_silodb = _T("Select * from dbo.Silo_Details");
	m_pRs_silodb.CreateInstance(_uuidof(Recordset));
	m_pRs_silodb->CursorLocation = adUseServer;
	m_pRs_silodb->Open((LPCTSTR)m_strCmdText_silodb, (LPCTSTR)m_strConnection_silodb, adOpenDynamic, adLockOptimistic, adCmdText);
	m_pRs_silodb->MoveFirst();

	bool m_silo_sel_status = 0;
	bool m_prev_silo_sel_status = 0;
	int index = 0;
	bool m_cur_silo_status = 0;

	while (!m_pRs_silodb->EndOfFile)
	{
		m_silo_sel_status = 0;
		m_cur_silo_status = 0;
		
		m_prev_silo_sel_status = m_pRs_silodb->Fields->GetItem(_T("Feed_Status"))->Value.boolVal; //Saving silo previous status
		Sleep(200);

		m_cur_silo_status = pValues[index].vDataValue.boolVal;
		if ((m_prev_silo_sel_status) != (m_cur_silo_status)) //Update data only when data has changed
		{
			//COleDateTime m_stamp_time = COleDateTime::GetCurrentTime();
			COleDateTime m_stamp_time = pValues[index].ftTimeStamp;
		
			if (pValues[index].vDataValue.boolVal)
				m_silo_sel_status = 1;
			_variant_t var;
			var.vt = VT_NULL;


			m_pRs_silodb->Update(("Feed_Status"), pValues[index].vDataValue);//saving current silo status

			if (!m_silo_sel_status) // Silo selection is being reset
			{
				if (m_prev_silo_sel_status)
					m_pRs_silodb->Update(("Feed_Stop_Time"), (_variant_t)m_stamp_time);

			}
			else  // Silo selection is being set
			{
				m_pRs_silodb->Update(("Feed_Start_Time"), (_variant_t)m_stamp_time);
				m_pRs_silodb->Fields->GetItem("Feed_Stop_Time")->Value = var;				
			}

			m_pRs_silodb->Update();
		}

		index++;
		m_pRs_silodb->MoveNext();
	}

	m_pRs_silodb->Close();

	/////////OPENING COMMON ULS PLC DATABASE ////////
	m_strCmdText_silodb = _T("Select * from dbo.Common_ULS");
	m_pRs_silodb.CreateInstance(_uuidof(Recordset));
	m_pRs_silodb->CursorLocation = adUseServer;
	m_pRs_silodb->Open((LPCTSTR)m_strCmdText_silodb, (LPCTSTR)m_strConnection_silodb, adOpenDynamic, adLockOptimistic, adCmdText);

	
	m_pRs_silodb->Update(("Y37_TPH"), pValues[81].vDataValue);	
	m_pRs_silodb->Update(("Y37_TOTALIZED"), pValues[82].vDataValue);	
	m_pRs_silodb->Update(("Y38_TPH"), pValues[83].vDataValue);	
	m_pRs_silodb->Update(("Y38_TOTALIZED"), pValues[84].vDataValue);	
	m_pRs_silodb->Update(("Y1_P1"), pValues[85].vDataValue);	
	m_pRs_silodb->Update(("Y1_P2"), pValues[86].vDataValue);	
	m_pRs_silodb->Update(("Y1_P3"), pValues[87].vDataValue);	
	m_pRs_silodb->Update(("Y1_P4"), pValues[88].vDataValue);	
	m_pRs_silodb->Update(("Y1_P5"), pValues[89].vDataValue);	
	m_pRs_silodb->Update(("Y1_P6"), pValues[90].vDataValue);	
	m_pRs_silodb->Update(("Y2_P1"), pValues[91].vDataValue);	
	m_pRs_silodb->Update(("Y2_P2"), pValues[92].vDataValue);	
	m_pRs_silodb->Update(("Y2_P3"), pValues[93].vDataValue);	
	m_pRs_silodb->Update(("Y2_P4"), pValues[94].vDataValue);	
	m_pRs_silodb->Update(("Y2_P5"), pValues[95].vDataValue);	
	m_pRs_silodb->Update(("Y2_P6"), pValues[96].vDataValue);
	m_pRs_silodb->Update(("Y1_RUN"), pValues[97].vDataValue);
	m_pRs_silodb->Update(("Y2_RUN"), pValues[98].vDataValue);

	m_pRs_silodb->Close();
}

void Silo_Calculation(void)
{
	ReadSynchData_ULS_PLC();//Reading Synchrous PLC ULS
	Sleep(2000);

	bool Y1_P1, Y1_P2, Y1_P3, Y1_P4, Y1_P5, Y1_P6, Y2_P1, Y2_P2, Y2_P3, Y2_P4, Y2_P5, Y2_P6;
	Y1_P1 = Y1_P2 = Y1_P3 = Y1_P4 = Y1_P5 = Y1_P6 = Y2_P1 = Y2_P2 = Y2_P3 = Y2_P4 = Y2_P5 = Y2_P6 = 0;

	float Y37_TPH_TOTAL, Y38_TPH_TOTAL, Y37_TPH_TOTAL_PREV, Y38_TPH_TOTAL_PREV;
	Y37_TPH_TOTAL = Y38_TPH_TOTAL = 0;
	Y37_TPH_TOTAL_PREV = Y38_TPH_TOTAL_PREV = 0;

	CString m_Silo_Ref_Y37 = _T("");
	CString m_Silo_Ref_Y38 = _T("");
	float m_SiloStock_Y37, m_SiloStock_Y38, m_add_stock_Y37, m_add_stock_Y38;
	m_SiloStock_Y37 = m_SiloStock_Y38 = m_add_stock_Y37 = m_add_stock_Y38 = 0;
	
	///////////OPENING ULS COMMON DATABASE ///////////////////
	m_strConnection_uls = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_uls = _T("Select * from Common_ULS");
	m_pRs_uls.CreateInstance(_uuidof(Recordset));
	m_pRs_uls->CursorLocation = adUseServer;
	m_pRs_uls->Open((LPCTSTR)m_strCmdText_uls, (LPCTSTR)m_strConnection_uls, adOpenDynamic, adLockOptimistic, adCmdText);

	Y37_TPH_TOTAL = m_pRs_uls->Fields->GetItem("Y37_TOTALIZED")->Value.fltVal;
	Y38_TPH_TOTAL = m_pRs_uls->Fields->GetItem("Y38_TOTALIZED")->Value.fltVal;
	Y37_TPH_TOTAL_PREV = m_pRs_uls->Fields->GetItem("Y37_TOTALIZED_PREV")->Value.fltVal;
	Y38_TPH_TOTAL_PREV = m_pRs_uls->Fields->GetItem("Y38_TOTALIZED_PREV")->Value.fltVal;

	Y1_P1 = m_pRs_uls->Fields->GetItem("Y1_P1")->Value.boolVal;
	Y1_P2 = m_pRs_uls->Fields->GetItem("Y1_P2")->Value.boolVal;
	Y1_P3 = m_pRs_uls->Fields->GetItem("Y1_P3")->Value.boolVal;
	Y1_P4 = m_pRs_uls->Fields->GetItem("Y1_P4")->Value.boolVal;
	Y1_P5 = m_pRs_uls->Fields->GetItem("Y1_P5")->Value.boolVal;
	Y1_P6 = m_pRs_uls->Fields->GetItem("Y1_P6")->Value.boolVal;

	Y2_P1 = m_pRs_uls->Fields->GetItem("Y2_P1")->Value.boolVal;
	Y2_P2 = m_pRs_uls->Fields->GetItem("Y2_P2")->Value.boolVal;
	Y2_P3 = m_pRs_uls->Fields->GetItem("Y2_P3")->Value.boolVal;
	Y2_P4 = m_pRs_uls->Fields->GetItem("Y2_P4")->Value.boolVal;
	Y2_P5 = m_pRs_uls->Fields->GetItem("Y2_P5")->Value.boolVal;
	Y2_P6 = m_pRs_uls->Fields->GetItem("Y2_P6")->Value.boolVal;

	if (Y37_TPH_TOTAL != Y37_TPH_TOTAL_PREV) //Checking for Y37 material flow
	{
		///////// Selection for Silo No. ////////////
		if (Y1_P1)
			m_Silo_Ref_Y37 = _T("1");
		if (Y1_P2)
			m_Silo_Ref_Y37 = GetSiloNumber_P2();
		if (Y1_P3)
			m_Silo_Ref_Y37 = _T("2");
		if (Y1_P4)
			m_Silo_Ref_Y37 = GetSiloNumber_P4();
		if (Y1_P5)
			m_Silo_Ref_Y37 = _T("2A");
		if (Y1_P6)
			m_Silo_Ref_Y37 = GetSiloNumber_P6();

		//////// Adding Coal Stock to the selected Silo /////////////

		if (m_Silo_Ref_Y37 != _T(""))    //Checking if Silo is selected or not
		{
			if (Y37_TPH_TOTAL > Y37_TPH_TOTAL_PREV)  //Totalizer not reset
				m_add_stock_Y37 = Y37_TPH_TOTAL - Y37_TPH_TOTAL_PREV;
			else //Totalizer reset
			{
				if ((Y37_TPH_TOTAL < 500) && (Y37_TPH_TOTAL_PREV > 9999000))  //Totalizer reset recently
					m_add_stock_Y37 = (9999990 - Y37_TPH_TOTAL_PREV) + Y37_TPH_TOTAL;
			}

			m_SiloStock_Y37 = AddSiloStock(m_Silo_Ref_Y37, m_add_stock_Y37);  //Adding additional coal to Silo and saving in Silo database

			m_pRs_uls->Update(("Y37_TOTALIZED_PREV"), (_variant_t)Y37_TPH_TOTAL); //Saving Current Totalized value to Previous Totalized value

		}
	}

	if (Y38_TPH_TOTAL != Y38_TPH_TOTAL_PREV) //Checking for Y37 material flow
	{
		///////// Selection for Silo No. ////////////
		if (Y2_P1)
			m_Silo_Ref_Y38 = _T("1");
		if (Y2_P2)
			m_Silo_Ref_Y38 = GetSiloNumber_P2();
		if (Y2_P3)
			m_Silo_Ref_Y38 = _T("2");
		if (Y2_P4)
			m_Silo_Ref_Y38 = GetSiloNumber_P4();
		if (Y2_P5)
			m_Silo_Ref_Y38 = _T("2A");
		if (Y2_P6)
			m_Silo_Ref_Y38 = GetSiloNumber_P6();

		//////// Adding Coal Stock to the selected Silo /////////////

		if (m_Silo_Ref_Y38 != _T(""))    //Checking if Silo is selected or not
		{
			if (Y38_TPH_TOTAL > Y38_TPH_TOTAL_PREV)  //Totalizer not reset
				m_add_stock_Y38 = Y38_TPH_TOTAL - Y38_TPH_TOTAL_PREV;
			else //Totalizer reset
			{
				if ((Y38_TPH_TOTAL < 500) && (Y38_TPH_TOTAL_PREV > 9999000))  //Totalizer reset recently
					m_add_stock_Y38 = (9999990 - Y38_TPH_TOTAL_PREV) + Y38_TPH_TOTAL;
			}

			m_SiloStock_Y38 = AddSiloStock(m_Silo_Ref_Y38, m_add_stock_Y38);  //Adding additional coal to Silo and saving in Silo database

			m_pRs_uls->Update(("Y38_TOTALIZED_PREV"), (_variant_t)Y38_TPH_TOTAL); //Saving Current Totalized value to Previous Totalized value

		}
	}
	m_pRs_uls->Close();
}

CString GetSiloNumber_P2(void)
{
	
	CString m_SiloNumber = _T("");
	///////////OPENING SILO DETAILS DATABASE ///////////////////
	m_strConnection_silo = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_silo = _T("Select * from dbo.Silo_Details");
	m_pRs_silo.CreateInstance(_uuidof(Recordset));
	m_pRs_silo->CursorLocation = adUseServer;
	m_pRs_silo->Open((LPCTSTR)m_strCmdText_silo, (LPCTSTR)m_strConnection_silo, adOpenDynamic, adLockOptimistic, adCmdText);
	
	m_pRs_silo->MoveFirst();
	m_pRs_silo->Move(1); // Initial Ref Silo Number 3

	for (int i = 0; i <= 25; ++i)  //For Silo Number 3 to 53
	{
		if (m_pRs_silo->Fields->GetItem("Feed_Status")->Value.boolVal)  /////Checking  if Silo is manually selected in PLC-HMI or not
			m_SiloNumber = m_pRs_silo->Fields->GetItem("Silo_No")->Value;		

		m_pRs_silo->MoveNext();
	}

	m_pRs_silo->Close();

	return(m_SiloNumber);
}

CString GetSiloNumber_P4(void)
{
	CString m_SiloNumber = _T("");
	///////////OPENING SILO DETAILS DATABASE ///////////////////
	m_strConnection_silo = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_silo = _T("Select * from dbo.Silo_Details");
	m_pRs_silo.CreateInstance(_uuidof(Recordset));
	m_pRs_silo->CursorLocation = adUseServer;
	m_pRs_silo->Open((LPCTSTR)m_strCmdText_silo, (LPCTSTR)m_strConnection_silo, adOpenDynamic, adLockOptimistic, adCmdText);

	m_pRs_silo->MoveFirst();
	m_pRs_silo->Move(28); // Initial Ref Silo Number 4

	
	for (int i = 0; i <= 25; ++i)  //For Silo Number 4 to 54
	{
	
		if (m_pRs_silo->Fields->GetItem("Feed_Status")->Value.boolVal)     /////Checking  if Silo is manually selected in PLC-HMI or not
			m_SiloNumber = m_pRs_silo->Fields->GetItem("Silo_No")->Value;
			
		m_pRs_silo->MoveNext();
	}

	m_pRs_silo->Close();

	return(m_SiloNumber);
}

CString GetSiloNumber_P6(void)
{
	CString m_SiloNumber = _T("");
	///////////OPENING SILO DETAILS DATABASE ///////////////////
	m_strConnection_silo = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_silo = _T("Select * from dbo.Silo_Details");
	m_pRs_silo.CreateInstance(_uuidof(Recordset));
	m_pRs_silo->CursorLocation = adUseServer;
	m_pRs_silo->Open((LPCTSTR)m_strCmdText_silo, (LPCTSTR)m_strConnection_silo, adOpenDynamic, adLockOptimistic, adCmdText);

	m_pRs_silo->MoveFirst();
	m_pRs_silo->Move(55); // Initial Ref Silo Number 4A

	CString str;
	for (int i = 0; i <= 25; ++i)  //For Silo Number 4A to 54A
	{
		if (m_pRs_silo->Fields->GetItem("Feed_Status")->Value.boolVal)     /////Checking  if Silo is manually selected in PLC-HMI or not
			m_SiloNumber = m_pRs_silo->Fields->GetItem("Silo_No")->Value;

		m_pRs_silo->MoveNext();
	}

	m_pRs_silo->Close();

	return(m_SiloNumber);
}

float AddSiloStock(CString m_SiloNumber, float m_add_stock)
{
	float m_SiloStock = 0;
	_variant_t var;

	///////////OPENING SILO DETAILS DATABASE ///////////////////
	m_strConnection_silo = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_silo = _T("Select * from dbo.Silo_Details");
	m_pRs_silo.CreateInstance(_uuidof(Recordset));
	m_pRs_silo->CursorLocation = adUseServer;
	m_pRs_silo->Open((LPCTSTR)m_strCmdText_silo, (LPCTSTR)m_strConnection_silo, adOpenDynamic, adLockOptimistic, adCmdText);
	m_pRs_silo->MoveFirst();

	CString str;

	if (m_add_stock >= 0)
	{
		while (!m_pRs_silo->EndOfFile)
		{
			str = m_pRs_silo->Fields->GetItem("Silo_No")->Value;
			if (str == m_SiloNumber)
			{
				var = m_pRs_silo->Fields->GetItem("Silo_Stock")->Value;
				m_SiloStock = var.dblVal;
				m_SiloStock = m_SiloStock + m_add_stock;

				if (m_SiloStock > 0)
					m_pRs_silo->Update(("Silo_Stock"), (_variant_t)m_SiloStock);  //Saving new silo stock to database
			}
			m_pRs_silo->MoveNext();
		}
	}
	m_pRs_silo->Close();

	return(m_SiloStock);
}

void Silo_Calculation_WF()
{
	CString m_strConnection_WF, m_strCmdText_WF, m_strCmdText1_WF;
	_RecordsetPtr m_pRs_WF, m_pRs1_WF;
	bool Y9_STATUS, Y10_STATUS, Y10A_STATUS;
	Y9_STATUS = Y10_STATUS = Y10A_STATUS = 0;
	_variant_t var;
	double m_silo_discharge_rate = 0;
	
	double m_totalizedVal = 0;
	double m_totalized_prevVal = 0;
	double diff = 0;
	double m_silo_stock = 0;

	m_strConnection_WF = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_WF = _T("Select * from dbo.Silo_Details ORDER BY Si_No ASC ");
	m_pRs_WF.CreateInstance(_uuidof(Recordset));
	m_pRs_WF->CursorLocation = adUseServer;
	m_pRs_WF->Open((LPCTSTR)m_strCmdText_WF, (LPCTSTR)m_strConnection_WF, adOpenDynamic, adLockOptimistic, adCmdText);

	m_strCmdText1_WF = _T("Select * from dbo.Common_FS");
	m_pRs1_WF.CreateInstance(_uuidof(Recordset));
	m_pRs1_WF->CursorLocation = adUseServer;
	m_pRs1_WF->Open((LPCTSTR)m_strCmdText1_WF, (LPCTSTR)m_strConnection_WF, adOpenDynamic, adLockOptimistic, adCmdText);

	Y9_STATUS = m_pRs1_WF->Fields->GetItem(_T("Y9_RUN"))->Value.boolVal;
	Y10_STATUS = m_pRs1_WF->Fields->GetItem(_T("Y10_RUN"))->Value.boolVal;
	Y10A_STATUS = m_pRs1_WF->Fields->GetItem(_T("Y10A_RUN"))->Value.boolVal;

	m_pRs1_WF->Close();

	if (Y9_STATUS) //Checking Y9 conveyor runing status
	{
		for (int n_index = 0; n_index <= 26; n_index++)
		{
			var = m_pRs_WF->Fields->GetItem(_T("WF_START"))->Value;
			if (var.boolVal)
			{
				var = m_pRs_WF->Fields->GetItem(_T("TPH_PV"))->Value;
				m_silo_discharge_rate = var.dblVal;

				if (m_silo_discharge_rate >= 20)  //For valid signal only
				{
										
					var = m_pRs_WF->Fields->GetItem(_T("WF_TOTALIZED_PREV"))->Value;
					m_totalized_prevVal = var.dblVal;
					
					var = m_pRs_WF->Fields->GetItem(_T("WF_TOTALIZED"))->Value;
					m_totalizedVal = var.dblVal;

					if (m_totalizedVal > m_totalized_prevVal) 
					{	
						diff = m_totalizedVal - m_totalized_prevVal;
						m_silo_stock = m_pRs_WF->Fields->GetItem(_T("Silo_Stock"))->Value.dblVal;
						m_silo_stock = m_silo_stock - diff; // deduct material outflow from current silo stock

						if (m_silo_stock < 0)
							m_silo_stock = 0;

						m_pRs_WF->Update(("Silo_Stock"), (variant_t)m_silo_stock); // update db
						m_pRs_WF->Update(("WF_TOTALIZED_PREV"), (variant_t)m_totalizedVal); // update db with current totalized to prev_totalized val

					}
									
				}
			}

			m_pRs_WF->MoveNext();
		}
	} // Y9 ends 

	// reset recordset for Y10
	m_pRs_WF->MoveFirst();
	m_pRs_WF->Move(26);
	if (Y10_STATUS) //Checking Y10 conveyor runing status
	{
		for (int n_index = 0; n_index <= 26; n_index++)
		{
			var = m_pRs_WF->Fields->GetItem(_T("WF_START"))->Value;
			if (var.boolVal)
			{
				var = m_pRs_WF->Fields->GetItem(_T("TPH_PV"))->Value;
				m_silo_discharge_rate = var.dblVal;

				if (m_silo_discharge_rate >= 20)  //For valid signal only
				{

					var = m_pRs_WF->Fields->GetItem(_T("WF_TOTALIZED_PREV"))->Value;
					m_totalized_prevVal = var.dblVal;

					var = m_pRs_WF->Fields->GetItem(_T("WF_TOTALIZED"))->Value;
					m_totalizedVal = var.dblVal;

					if (m_totalizedVal > m_totalized_prevVal)
					{
						diff = m_totalizedVal - m_totalized_prevVal;
						m_silo_stock = m_pRs_WF->Fields->GetItem(_T("Silo_Stock"))->Value.dblVal;
						m_silo_stock = m_silo_stock - diff; // deduct material outflow from current silo stock

						if (m_silo_stock < 0)
							m_silo_stock = 0;

						m_pRs_WF->Update(("Silo_Stock"), (variant_t)m_silo_stock); // update db
						m_pRs_WF->Update(("WF_TOTALIZED_PREV"), (variant_t)m_totalizedVal); // update db with current totalized to prev_totalized val

					}

				}
			}

			m_pRs_WF->MoveNext();
		}
	} // Y10 ends

	// reset recordset for Y10A
	m_pRs_WF->MoveFirst();
	m_pRs_WF->Move(53);
	if (Y10_STATUS) //Checking Y10A conveyor runing status
	{
		for (int n_index = 0; n_index <= 26; n_index++)
		{
			var = m_pRs_WF->Fields->GetItem(_T("WF_START"))->Value;
			if (var.boolVal)
			{
				var = m_pRs_WF->Fields->GetItem(_T("TPH_PV"))->Value;
				m_silo_discharge_rate = var.dblVal;

				if (m_silo_discharge_rate >= 20)  //For valid signal only
				{

					var = m_pRs_WF->Fields->GetItem(_T("WF_TOTALIZED_PREV"))->Value;
					m_totalized_prevVal = var.dblVal;

					var = m_pRs_WF->Fields->GetItem(_T("WF_TOTALIZED"))->Value;
					m_totalizedVal = var.dblVal;

					if (m_totalizedVal > m_totalized_prevVal)
					{
						diff = m_totalizedVal - m_totalized_prevVal;
						m_silo_stock = m_pRs_WF->Fields->GetItem(_T("Silo_Stock"))->Value.dblVal;
						m_silo_stock = m_silo_stock - diff; // deduct material outflow from current silo stock

						if (m_silo_stock < 0)
							m_silo_stock = 0;

						m_pRs_WF->Update(("Silo_Stock"), (variant_t)m_silo_stock); // update db
						m_pRs_WF->Update(("WF_TOTALIZED_PREV"), (variant_t)m_totalizedVal); // update db with current totalized to prev_totalized val

					}

				}
			}

			m_pRs_WF->MoveNext();
		}
	} // Y10A ends


	m_pRs_WF->Close();

}

void Ash_Calculation()
{
	CString m_strConnection;
	CString m_strCmdText;
	_RecordsetPtr m_pRs_ash, m_pRs_tph, m_pRs_silo, m_pRs_conv, m_pRs_silo_1, m_pRs_silo_2;

	double ash_dry_basis = 0;
	double tph_y1 = 0;		// tph for y1, taken from analyzer
	double tph_y2 = 0;		// tph for y2, taken from PLC

	double coal_ash = 0;
	double silo_stock = 0; 
	double ash_temp = 0;

	CString silo_1 = _T("");  // 2 silos selected for feed in; 1 for y37; 2 for y38
	CString silo_2 = _T("");

	int silo_select_count = 1;

	//bool y37_run = FALSE;
	//bool y38_run = FALSE;
	bool y1_run = FALSE;
	bool y2_run = FALSE;
	bool feed_status = FALSE;

	CString silo_temp = _T("");

	// check y37 and y38 status; better use y1 and y2 correspondingly
	// analyzer on y37; so if y37 runs assume the same coal on y38, else dont calculate
	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText = _T("Select * from dbo.Common_ULS");
	m_pRs_conv.CreateInstance(_uuidof(Recordset));
	m_pRs_conv->CursorLocation = adUseServer;
	m_pRs_conv->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	y1_run = m_pRs_conv->Fields->GetItem(_T("Y1_RUN"))->Value.boolVal;
	y2_run = m_pRs_conv->Fields->GetItem(_T("Y2_RUN"))->Value.boolVal;

	m_pRs_conv->Close();

	if (y1_run == TRUE )  // y1 running ; TO DO: enumerate the rest 
	{	// identify silos with feed status 1
		m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
		m_strCmdText = _T("Select * from dbo.Silo_Details");
		m_pRs_silo.CreateInstance(_uuidof(Recordset));
		m_pRs_silo->CursorLocation = adUseServer;
		m_pRs_silo->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

		
		while (!m_pRs_silo->EndOfFile)
		{
			feed_status = m_pRs_silo->Fields->GetItem(_T("Feed_Status"))->Value.boolVal;
			silo_temp = m_pRs_silo->Fields->GetItem(_T("Silo_No"))->Value;

			if (feed_status && (ValidatePath(silo_temp) != _T("")))
			{
				if (silo_select_count == 1)
				{
					silo_1 = silo_temp;
				}
					
				else 
					silo_2 = silo_temp;
			}
			
			silo_select_count++;
			m_pRs_silo->MoveNext();
		}
		
		m_pRs_silo->Close();
	}

	// get latest ash_dry_basis for silo1 and silo2 from y37 analyser; analyser is only on y37
	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Analyzer_Db;Trusted=True;");
	m_strCmdText = _T("Select * from dbo.Y37_Analyzer ORDER BY Computer_Time DESC "); // TODO: order by time stamp desc; DONE
	m_pRs_ash.CreateInstance(_uuidof(Recordset));
	m_pRs_ash->CursorLocation = adUseServer;
	m_pRs_ash->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	ash_dry_basis = m_pRs_ash->Fields->GetItem(_T("ASH_DB_1_MIN"))->Value.dblVal;  // get ash dry basis analyzer; DONE
	if (ash_dry_basis < 3.0) 
	{
		m_pRs_ash->Close();		// if ash is less than threshold, silently return
		return;
	}

	//tph_y1 = m_pRs_ash->Fields->GetItem(_T("TPH_AR_1_MIN"))->Value.dblVal;		// tph from analyzer only for y1/y37
																			// TODO: take tph from plc
	m_pRs_ash->Close();

	// get latest thp_y1 and tph_y2 from Common_ULS db ; tph_y1 is also available from the analyzer, but better use this db
	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Analyzer_Db;Trusted=True;");
	m_strCmdText = _T("Select * from dbo.Common_ULS "); // TODO: order by time stamp desc; DONE
	m_pRs_tph.CreateInstance(_uuidof(Recordset));
	m_pRs_tph->CursorLocation = adUseServer;
	m_pRs_tph->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);
	
	tph_y1 = m_pRs_tph->Fields->GetItem(_T("Y37_TPH"))->Value.dblVal;	// tph for y1 or y37
	tph_y2 = m_pRs_tph->Fields->GetItem(_T("Y38_TPH"))->Value.dblVal;	// tph for y2 or y38

	m_pRs_tph->Close();

	if (silo_1 != _T(""))				// Check if silo1 is fed by y1 or y2; so check the path
	{
		m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True;");
		m_strCmdText = _T("Select * from dbo.Silo_Details");					// parameterise silo1 in query
		m_pRs_silo_1.CreateInstance(_uuidof(Recordset));
		m_pRs_silo_1->CursorLocation = adUseServer;
		m_pRs_silo_1->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

		while (!m_pRs_silo_1->EndOfFile)
		{
			if (m_pRs_silo_1->Fields->GetItem(_T("Silo_No"))->Value == silo_1) 
			{
				// for silo_1
				coal_ash = m_pRs_silo_1->Fields->GetItem(_T("Coal_Ash"))->Value.dblVal;  // get ash saved from db for the silo
				silo_stock = m_pRs_silo_1->Fields->GetItem(_T("Silo_Stock"))->Value.dblVal;		// silo stock from db for the silo
				// calculate weighted avg and update the ash
				ash_temp = ((ash_dry_basis * tph_y1 / 60) + (coal_ash * silo_stock)) / (tph_y1 / 60 + silo_stock);
				m_pRs_silo_1->Update(("Coal_Ash"), (variant_t)ash_temp);
			}
		}
		
		m_pRs_silo_1->Close();

	}
	

	// assuming the update is done in 1 minute interval
	
	if (y1_run && !y2_run)
	{

	}


}

// update this function to match its definition elsewhere 
// finally move this to a common header file
CString ValidatePath(CString m_Silo_No)
{
	CString m_Path = _T("");

	bool Y1_P1, Y1_P2, Y1_P3, Y1_P4, Y1_P5, Y1_P6, Y2_P1, Y2_P2, Y2_P3, Y2_P4, Y2_P5, Y2_P6;
	Y1_P1 = Y1_P2 = Y1_P3 = Y1_P4 = Y1_P5 = Y1_P6 = Y2_P1 = Y2_P2 = Y2_P3 = Y2_P4 = Y2_P5 = Y2_P6 = 0;


	///////////OPENING ULS COMMON DATABASE ///////////////////
	CString  m_strConnection_uls, m_strCmdText_uls;
	_RecordsetPtr m_pRs_uls;

	m_strConnection_uls = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText_uls = _T("Select * from Common_ULS");
	m_pRs_uls.CreateInstance(_uuidof(Recordset));
	m_pRs_uls->CursorLocation = adUseServer;
	m_pRs_uls->Open((LPCTSTR)m_strCmdText_uls, (LPCTSTR)m_strConnection_uls, adOpenDynamic, adLockOptimistic, adCmdText);

	/////////READING SELECTED PATHS ///////////////
	Y1_P1 = m_pRs_uls->Fields->GetItem("Y1_P1")->Value.boolVal;
	Y1_P2 = m_pRs_uls->Fields->GetItem("Y1_P2")->Value.boolVal;
	Y1_P3 = m_pRs_uls->Fields->GetItem("Y1_P3")->Value.boolVal;
	Y1_P4 = m_pRs_uls->Fields->GetItem("Y1_P4")->Value.boolVal;
	Y1_P5 = m_pRs_uls->Fields->GetItem("Y1_P5")->Value.boolVal;
	Y1_P6 = m_pRs_uls->Fields->GetItem("Y1_P6")->Value.boolVal;

	Y2_P1 = m_pRs_uls->Fields->GetItem("Y2_P1")->Value.boolVal;
	Y2_P2 = m_pRs_uls->Fields->GetItem("Y2_P2")->Value.boolVal;
	Y2_P3 = m_pRs_uls->Fields->GetItem("Y2_P3")->Value.boolVal;
	Y2_P4 = m_pRs_uls->Fields->GetItem("Y2_P4")->Value.boolVal;
	Y2_P5 = m_pRs_uls->Fields->GetItem("Y2_P5")->Value.boolVal;
	Y2_P6 = m_pRs_uls->Fields->GetItem("Y2_P6")->Value.boolVal;

	m_pRs_uls->Close();
	////////////////////////////////////////////////////////////

	if ((Y1_P1) || (Y2_P1))
	{
		if (m_Silo_No == _T("1"))
		{
			if (Y1_P1)
				m_Path = _T("Y1_P1");
			else
				m_Path = _T("Y2_P1");
		}
	}

	if ((Y1_P2) || (Y2_P2))
	{
		if ((m_Silo_No == _T("3")) || (m_Silo_No == _T("5")) || (m_Silo_No == _T("7")) || (m_Silo_No == _T("9")) || (m_Silo_No == _T("11")) || (m_Silo_No == _T("13")) || (m_Silo_No == _T("15")) || (m_Silo_No == _T("17")) || (m_Silo_No == _T("19")) || (m_Silo_No == _T("21")) || (m_Silo_No == _T("23")) || (m_Silo_No == _T("25")) || (m_Silo_No == _T("27")) || (m_Silo_No == _T("29")) || (m_Silo_No == _T("31")) || (m_Silo_No == _T("33")) || (m_Silo_No == _T("35")) || (m_Silo_No == _T("37")) || (m_Silo_No == _T("39")) || (m_Silo_No == _T("41")) || (m_Silo_No == _T("43")) || (m_Silo_No == _T("45")) || (m_Silo_No == _T("47")) || (m_Silo_No == _T("49")) || (m_Silo_No == _T("51")) || (m_Silo_No == _T("53")))
		{
			if (Y1_P2)
				m_Path = _T("Y1_P2");
			else
				m_Path = _T("Y2_P2");
		}
	}

	if ((Y1_P3) || (Y2_P3))
	{
		if (m_Silo_No == _T("2"))
		{
			if (Y1_P3)
				m_Path = _T("Y1_P3");
			else
				m_Path = _T("Y2_P3");
		}
	}

	if ((Y1_P4) || (Y2_P4))
	{
		if ((m_Silo_No == _T("4")) || (m_Silo_No == _T("6")) || (m_Silo_No == _T("8")) || (m_Silo_No == _T("10")) || (m_Silo_No == _T("12")) || (m_Silo_No == _T("14")) || (m_Silo_No == _T("16")) || (m_Silo_No == _T("18")) || (m_Silo_No == _T("20")) || (m_Silo_No == _T("22")) || (m_Silo_No == _T("24")) || (m_Silo_No == _T("26")) || (m_Silo_No == _T("28")) || (m_Silo_No == _T("30")) || (m_Silo_No == _T("32")) || (m_Silo_No == _T("34")) || (m_Silo_No == _T("36")) || (m_Silo_No == _T("38")) || (m_Silo_No == _T("40")) || (m_Silo_No == _T("42")) || (m_Silo_No == _T("44")) || (m_Silo_No == _T("46")) || (m_Silo_No == _T("48")) || (m_Silo_No == _T("50")) || (m_Silo_No == _T("52")) || (m_Silo_No == _T("54")))
		{
			if (Y1_P4)
				m_Path = _T("Y1_P4");
			else
				m_Path = _T("Y2_P4");
		}
	}

	if ((Y1_P5) || (Y2_P5))
	{
		if (m_Silo_No == _T("2A"))
		{
			if (Y1_P3)
				m_Path = _T("Y1_P5");
			else
				m_Path = _T("Y2_P5");
		}
	}

	if ((Y1_P6) || (Y2_P6))
	{
		if ((m_Silo_No == _T("4A")) || (m_Silo_No == _T("6A")) || (m_Silo_No == _T("8A")) || (m_Silo_No == _T("10A")) || (m_Silo_No == _T("12A")) || (m_Silo_No == _T("14A")) || (m_Silo_No == _T("16A")) || (m_Silo_No == _T("18A")) || (m_Silo_No == _T("20A")) || (m_Silo_No == _T("22A")) || (m_Silo_No == _T("24A")) || (m_Silo_No == _T("26A")) || (m_Silo_No == _T("28A")) || (m_Silo_No == _T("30A")) || (m_Silo_No == _T("32A")) || (m_Silo_No == _T("34A")) || (m_Silo_No == _T("36A")) || (m_Silo_No == _T("38A")) || (m_Silo_No == _T("40A")) || (m_Silo_No == _T("42A")) || (m_Silo_No == _T("44A")) || (m_Silo_No == _T("46A")) || (m_Silo_No == _T("48A")) || (m_Silo_No == _T("50A")) || (m_Silo_No == _T("52A")) || (m_Silo_No == _T("54A")))
		{
			if (Y1_P6)
				m_Path = _T("Y1_P6");
			else
				m_Path = _T("Y2_P6");
		}
	}

	return(m_Path);
}
