#pragma once
#include "opcda.h"
#include "opccomn.h"

class COPCDataCallback :
	public IOPCDataCallback
{
public:
	COPCDataCallback(void);
	~COPCDataCallback(void);

	///////////////////////////

	private:
        ULONG       m_cRef;     //Reference count

   
        //IUnknown members
        STDMETHODIMP         QueryInterface(REFIID, void**);
        STDMETHODIMP_(DWORD) AddRef(void);
        STDMETHODIMP_(DWORD) Release(void);

        //IOPCDataCallback members
        STDMETHODIMP OnDataChange(DWORD, OPCHANDLE, HRESULT, HRESULT, DWORD,
								  OPCHANDLE*, VARIANT*, WORD*, FILETIME*, HRESULT*);
		STDMETHODIMP OnReadComplete(DWORD, OPCHANDLE, HRESULT, HRESULT, DWORD,
								    OPCHANDLE*, VARIANT*, WORD*, FILETIME*, HRESULT*);
        STDMETHODIMP OnWriteComplete(DWORD , OPCHANDLE, HRESULT , DWORD,OPCHANDLE*, HRESULT*);
        STDMETHODIMP OnCancelComplete(DWORD, OPCHANDLE);
        
		void SaveDatabase(int, int, long, VARIANT*, FILETIME*, long );

		CString m_strConnection, m_strCmdText, m_strCmdText1, m_strCmdText2;
		_RecordsetPtr m_pRs, m_pRs1, m_pRs2;
		int GetRow(long Row_Index);
		int GetCol(long Col_Index);
		
};
