#include "StdAfx.h"
#include ".\opcdatacallback.h"
//#include "utils.h"
#include <stdio.h>
#include <string.h>


_variant_t var;
CString str, str1;

COPCDataCallback::COPCDataCallback(void)
{
	m_cRef=0;
	
	m_strConnection = _T("DRIVER={SQL Server};Server=BLENDINGPC-2;Database=Coal_Blending_Db;Trusted=True");
	m_strCmdText = _T("Select * from dbo.Silo_Details");
	m_pRs.CreateInstance(_uuidof(Recordset));
	m_pRs->CursorLocation = adUseServer;
	m_pRs->Open((LPCTSTR)m_strCmdText, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);
	m_pRs->MoveFirst();

	m_strCmdText1 = _T("Select * from dbo.Common_FS");
	m_pRs1.CreateInstance(_uuidof(Recordset));
	m_pRs1->CursorLocation = adUseServer;
	m_pRs1->Open((LPCTSTR)m_strCmdText1, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);

	m_strCmdText2 = _T("Select * from Common_ULS");
	m_pRs2.CreateInstance(_uuidof(Recordset));
	m_pRs2->CursorLocation = adUseServer;
	m_pRs2->Open((LPCTSTR)m_strCmdText2, (LPCTSTR)m_strConnection, adOpenDynamic, adLockOptimistic, adCmdText);
	
	//fp = ::fopen("E:\\myfile.txt", "w");
	
	 
	return;
}

COPCDataCallback::~COPCDataCallback(void)
{
	//m_pRs->Close();
	//::fclose(fp);
	
	return;
}



/*
 * COPCDataCallback::QueryInterface
 * COPCDataCallback::AddRef
 * COPCDataCallback::Release
 *
 * Purpose:
 *  Non-delegating IUnknown members for COPCDataCallback.
 */

STDMETHODIMP COPCDataCallback::QueryInterface(REFIID riid
    , void** ppv)
{
    *ppv=NULL;

//    if (IID_IUnknown==riid || IID_IOPCDataCallback==riid)
//      *ppv=this;

    if (IID_IUnknown==riid)
		 *ppv=reinterpret_cast<IUnknown*>(this);
	else if (IID_IOPCDataCallback==riid)
		*ppv = (IOPCDataCallback*)this;

    if (NULL!=*ppv)
        {
        ((LPUNKNOWN)*ppv)->AddRef();
        return NOERROR;
        }

    return ResultFromScode(E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) COPCDataCallback::AddRef(void)
{
    return ++m_cRef;
}

STDMETHODIMP_(ULONG) COPCDataCallback::Release(void)
{
    if (0!=--m_cRef)
        return m_cRef;

    delete this;
    return 0;
}


        
/*
 * COPCDataCallback::OnDataChange
 *
 * Purpose: Called by the OPC server when a group is in subscription mode. If a data value changes,
 * the new value and other information is returned. If there is a problem with
 * an item, error information is returned.
 *  
 */

STDMETHODIMP COPCDataCallback::OnDataChange(DWORD dwTransid, OPCHANDLE hGroup, HRESULT hrMasterquality,
											HRESULT hrMastererror, DWORD dwCount,
								            OPCHANDLE *phClientItems, VARIANT *pvValues,
											WORD *pwQualities, FILETIME *pftTimeStamps,
											HRESULT *pErrors)
{
	int row, col;
	row = col = 0;

	DWORD n = 0;
	/* var = pvValues[53];
	 float f = var.fltVal;*/
//COleDateTime m_time;

	for (n=0; n < dwCount; n++)
	{
	
		switch(pwQualities[n])
		{
			case OPC_QUALITY_GOOD:
				
				row = GetRow(phClientItems[n]);
				col = GetCol(phClientItems[n]);
				SaveDatabase(row, col, n, pvValues, pftTimeStamps, phClientItems[n]);
			break;		
		}		
		
	}

	 //m_time = (COleDateTime)*pftTimeStamps;
		//	int date = m_time.GetDay();
	//const TCHAR* str5 = (LPCTSTR)str;
 
	// str1 = m_time.Format("%D:%M:%Y:%H:%M:%S", m_time);
 //char c;
 /*
 int x = 0;
 while (x < str.GetLength())
 {
	 //str5[x] = (char*)(LPCTSTR)str.GetAt(x);
	
	  x++;
 }
fp = ::fopen("E:\\myfile.txt", "w");
Sleep(500);

 if (fp)
 {
	 //::fprintf(fp, str5);
	 ::fwrite(str5, str.GetLength(), 1, fp);
	::fclose(fp);
 }*/

    return NOERROR;
}
STDMETHODIMP COPCDataCallback::OnReadComplete(DWORD dwTransid, OPCHANDLE hGroup, HRESULT hrMasterquality,
											HRESULT hrMastererror, DWORD dwCount,
								            OPCHANDLE *phClientItems, VARIANT *pvValues,
											WORD *pwQualities, FILETIME *pftTimeStamps,
											HRESULT *pErrors)
{



    return NOERROR;
}



/*
 * COPCDataCallback::OnWriteComplete
 *
 * Purpose: Called when an asynchronous write is performed. If successful, the 
 * transaction id and other information is returned. If the write fails, error
 * information is returned.
 */

STDMETHODIMP COPCDataCallback::OnWriteComplete(DWORD dwTransid, OPCHANDLE hGroup, HRESULT hrMastererror, DWORD dwCount,OPCHANDLE* phClientItems, HRESULT* pErrors)
{
	
    return NOERROR;
}


/*
 * COPCDataCallback::OnCancelComplete
 *
 * Purpose: Called by OPC server after a transaction is cancelled. It returns the 
 * transaction id and the group handle.
 *
 */
STDMETHODIMP COPCDataCallback::OnCancelComplete(DWORD dwTransid, OPCHANDLE hGroup)
{

	return NOERROR;
}



/*
 * COPCDataCallback::OnReadComplete
 *
 * Purpose: Called by the OPC server after an asynchronous read is performed. If successful, the current data value and 
 * and other information is returned. If the read fails, error information is returned.
 *
 */



void COPCDataCallback::SaveDatabase(int row, int col, long ref, VARIANT* pvValues, FILETIME* pftTimeStamps, long address)
{
	m_pRs->MoveFirst();
	if (row <= 80)
		m_pRs->Move(row);
	
	if (col == 10) // for silo feed status updating in database
	{
		bool m_silo_sel_status = 0;
		bool m_prev_silo_sel_status = 0;

		m_prev_silo_sel_status = m_pRs->Fields->GetItem(_T("Feed_Status"))->Value.boolVal; //Saving silo previous status
		Sleep(200);
		
		if (m_prev_silo_sel_status != pvValues[ref].boolVal) //Update data only when data has changed
		{	
			COleDateTime m_stamp_time;
			m_stamp_time = pftTimeStamps[ref];

			if (pvValues[ref].boolVal)
				m_silo_sel_status = 1;
			_variant_t var;
			var.vt = VT_NULL;


			m_pRs->Update(("Feed_Status"), pvValues[ref]);//saving current silo status

			if (!m_silo_sel_status) // Silo selection is being reset
			{
				if (m_prev_silo_sel_status)
					m_pRs->Update(("Feed_Stop_Time"), (_variant_t)m_stamp_time);

			}
			else  // Silo selection is being set
			{
				m_pRs->Update(("Feed_Start_Time"), (_variant_t)m_stamp_time);
				m_pRs->Fields->GetItem("Feed_Stop_Time")->Value = var;
				//m_pRs->Fields->Delete(7);		//Delete Previous Feed Stop Time
				m_pRs->Update();
			}
		}
	}

	if (col == 11) // for ULS Common Parameters
	{
		if (address == 20561)
			m_pRs2->Update(("Y37_TPH"), pvValues[ref]);

		if (address == 20562)
			m_pRs2->Update(("Y37_TOTALIZED"), pvValues[ref]);

		if (address == 20563)
			m_pRs2->Update(("Y38_TPH"), pvValues[ref]);

		if (address == 20564)
			m_pRs2->Update(("Y38_TOTALIZED"), pvValues[ref]);

		if (address == 20565)
			m_pRs2->Update(("Y1_P1"), pvValues[ref]);

		if (address == 20566)
			m_pRs2->Update(("Y1_P2"), pvValues[ref]);

		if (address == 20567)
			m_pRs2->Update(("Y1_P3"), pvValues[ref]);

		if (address == 20568)
			m_pRs2->Update(("Y1_P4"), pvValues[ref]);

		if (address == 20569)
			m_pRs2->Update(("Y1_P5"), pvValues[ref]);

		if (address == 20570)
			m_pRs2->Update(("Y1_P6"), pvValues[ref]);


		if (address == 20571)
			m_pRs2->Update(("Y2_P1"), pvValues[ref]);

		if (address == 20572)
			m_pRs2->Update(("Y2_P2"), pvValues[ref]);

		if (address == 20573)
			m_pRs2->Update(("Y2_P3"), pvValues[ref]);

		if (address == 20574)
			m_pRs2->Update(("Y2_P4"), pvValues[ref]);

		if (address == 20575)
			m_pRs2->Update(("Y2_P5"), pvValues[ref]);

		if (address == 20576)
			m_pRs2->Update(("Y2_P6"), pvValues[ref]);


	}

	if (col == 1) // for weigh feeder PV
	{
		double m_wf_pv = 0;
		m_wf_pv = pvValues[ref].fltVal;
		float m_zero = 0;
		
		if (m_wf_pv >= 0) //saving only valid values
			m_pRs->Update(("TPH_PV"), pvValues[ref]);
		else
			m_pRs->Update(("TPH_PV"), (_variant_t) m_zero);
		
	}

	if (col == 2) // for weigh feeder SP_HMI
	{
		double m_wf_sp_hmi = 0;
		m_wf_sp_hmi = pvValues[ref].fltVal;

		if (m_wf_sp_hmi >= 0) //saving only valid values
			m_pRs->Update(("TPH_SP_HMI"), pvValues[ref]);

	}

	if (col == 3) // for weigh feeder START PB
	{
		bool m_silo_sel_status = 0;
		bool m_prev_silo_sel_status = 0;
		_variant_t var_stop;
		var_stop.vt = VT_NULL;

		m_prev_silo_sel_status = m_pRs->Fields->GetItem(_T("WF_START"))->Value.boolVal; //Saving silo previous status
		Sleep(50);

		if (m_prev_silo_sel_status != pvValues[ref].boolVal) //Update data only when data has changed
		{
			COleDateTime m_stamp_time;
			m_stamp_time = pftTimeStamps[ref];

			m_pRs->Update(("WF_START"), pvValues[ref]);//saving current silo status

			if (!m_prev_silo_sel_status && pvValues[ref].boolVal)
			{
				m_pRs->Update(("WF_START_TIME"), (_variant_t)m_stamp_time); //Saving WF START TIME
				m_pRs->Update(("WF_STOP_TIME"), var_stop); //Seting WF STop TIME to NULL
			}
			
		}
	}

	if (col == 4) // for weigh feeder STOP PB
	{
		bool m_silo_sel_status = 0;
		bool m_prev_silo_sel_status = 0;

		m_prev_silo_sel_status = m_pRs->Fields->GetItem(_T("WF_STOP"))->Value.boolVal; //Saving silo previous status
		Sleep(50);

		if (m_prev_silo_sel_status != pvValues[ref].boolVal) //Update data only when data has changed
		{
			COleDateTime m_stamp_time;
			m_stamp_time = pftTimeStamps[ref];

			m_pRs->Update(("WF_STOP"), pvValues[ref]);//saving current silo status

			if (!m_prev_silo_sel_status && pvValues[ref].boolVal)
				m_pRs->Update(("WF_STOP_TIME"), (_variant_t)m_stamp_time); //Saving WF STop TIME

		}
	}

	if (col == 5) // for weigh feeder TOTALIZED FLOW
	{
		double m_wf_pv = 0;
		m_wf_pv = pvValues[ref].fltVal;

		if (m_wf_pv > 20) //saving only valid values
			m_pRs->Update(("WF_TOTALIZED"), pvValues[ref]);

	}

	if (col == 6) // for weigh feeder COMMON PARAMETERS
	{
		if (address == 12693)
			m_pRs1->Update(("Y9_RUN"), pvValues[ref]);

		if (address == 12694)
			m_pRs1->Update(("Y10_RUN"), pvValues[ref]);

		if (address == 12695)
			m_pRs1->Update(("Y10A_RUN"), pvValues[ref]);
	}
	
}
int COPCDataCallback::GetRow(long Row_Index)
{
	int increment = 0;
	long offset = 0;

	///////////////  FOR FS PLC ///////////////////////
	////////////////////////////////////////////////////

	if ((Row_Index >= 12288) && (Row_Index <= 12368)) ///FROM  3000 HEX + 80 = 3050 HEX for 81 Weigh feeder PV
		increment = Row_Index - 12288;
		
	if ((Row_Index >= 12369) && (Row_Index <= 12449)) ///FOR  3051 HEX + 80 =   30A1 HEX for 81 Weigh feeder SP
		increment = Row_Index - 12369;

	if ((Row_Index >= 12450) && (Row_Index <= 12530)) ///FOR  30A2 HEX + 80 =   30F2 HEX for 81 Weigh feeder START PB
		increment = Row_Index - 12450;

	if ((Row_Index >= 12531) && (Row_Index <= 12611)) ///FOR  30A2 HEX + 80 =   3143 HEX for 81 Weigh feeder STOP PB
		increment = Row_Index - 12531;

	if ((Row_Index >= 12612) && (Row_Index <= 12692)) ///FOR  3144 HEX + 80 =   3194 HEX for 81 Weigh feeder TOTALIZED FLOW
		increment = Row_Index - 12612;

	if ((Row_Index >= 12693) && (Row_Index <= 12695)) ///FOR  3195 HEX + 2 =   3197 HEX for 81 Weigh feeder 3 Nos. COMMON PARAMETERS
		increment = Row_Index - 12693;

	
	///////////////  FOR ULS PLC ///////////////////////
	////////////////////////////////////////////////////

	if ((Row_Index >= 20480) && (Row_Index <= 20560)) ///FOR  5000 HEX + 80 =   5050 HEX for 81 ULS SILO SELECTION
		increment = Row_Index - 20480;

			
	return increment;

}

int COPCDataCallback::GetCol(long Col_Index)
{
	int remain = 0;
	

	///////////////  FOR FS PLC ///////////////////////
	////////////////////////////////////////////////////

	if ((Col_Index >= 12288) && (Col_Index <= 12368))  ///FROM  3000 HEX + 80 = 3050 HEX for 81 silo PV
		remain = 1;

	if ((Col_Index >= 12369) && (Col_Index <= 12449)) ///FROM  3000 HEX + 80 = 3050 HEX for 81 silo SP-HMI
		remain = 2;

	if ((Col_Index >= 12450) && (Col_Index <= 12530)) ///FOR  30A2 HEX + 80 =   30F2 HEX for 81 Weigh feeder START PB
		remain = 3;

	if ((Col_Index >= 12531) && (Col_Index <= 12611)) ///FOR  30A2 HEX + 80 =   3143 HEX for 81 Weigh feeder STOP PB
		remain = 4;


	if ((Col_Index >= 12612) && (Col_Index <= 12692)) ///FOR  3144 HEX + 80 =   3194 HEX for 81 Weigh feeder TOTALIZED FLOW
		remain = 5;

	if ((Col_Index >= 12693) && (Col_Index <= 12695)) ///FOR  3195 HEX + 2 =   3197 HEX for 81 Weigh feeder 3 Nos. COMMON PARAMETERS
		remain = 6;


	///////////////  FOR ULS PLC ///////////////////////
	////////////////////////////////////////////////////
	
	
	if ((Col_Index >= 20480) && (Col_Index <= 20560)) ///FOR  5000 HEX + 80 =   5050 HEX for 81 ULS SILO SELECTION
		remain = 10;

	if ((Col_Index >= 20561) && (Col_Index <= 20576)) ///FOR  5051 HEX + 16 =   5060 HEX for 16 ULS COMMON PARAMETERS
		remain = 11;

	return remain;
}


