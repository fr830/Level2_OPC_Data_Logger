#pragma once

#include "resource.h"

#define PROGID_SERVER_THERMO		L"ThermoScientific.Datapool" /// Class id OPC driver of Thermofisher Elemental Analyzer
#define PROGID_SERVER_PLC			L"Schneider-Aut.OFS" /// Class id OPC driver of Schneider PLC

BOOL CreateOPCConnection_Y37(); /// Procedure for creating OPC client and related other procedure ////////

HRESULT CreateOPCServer_Thermo();  /// Procedure for creating OPC client and related other procedure ////////
HRESULT CreateOPCGroup_Y37();  // Create the group 

HRESULT AddOPCItems_Read_Y37(void);  //Creating OPC Items for 1 MIN CURRENT ANALYSIS

HRESULT Read_Y37_Data(void); //Reading Y37 Analyzer data for 1 MIN CURRENT ANALYSIS
HRESULT AddOPCItems_Read_Y37_10MIN(void);  //Creating OPC Items for 10 MIN ROLLING ANALYSIS
HRESULT Read_Y37_Data_10MIN(void); //Reading Y37 Analyzer data for 10 MIN ROLLING ANALYSIS

HRESULT Read_Y18_Data(void); //Reading Y18 Analyzer data for 1 MIN CURRENT ANALYSIS
HRESULT AddOPCItems_Read_Y18_1MIN(void);  //Creating OPC Items for 1 MIN ROLLING ANALYSIS Y18
HRESULT Read_Y18_Data_1MIN(void); //Reading Y37 Analyzer data for 1 MIN ROLLING ANALYSIS Y18



BOOL CreateOPCConnection_PLC(); /// Procedure for creating OPC client and related other procedure for PLC ////////
HRESULT CreateOPCServer_PLC();  /// Procedure for creating OPC client and related other procedure ////////
HRESULT CreateOPCGroup_PLC_FS();  // Create the items for Asynch Read for FS PLC
HRESULT CreateOPCGroup_PLC_ULS();  // Create the items for Asynch Read for ULS PLC

HRESULT AddOPCItems_Read_Asynch_PLC_FS(void); //Creating OPC Items for FS PLC Asynch_read
HRESULT ReadAsynchData_FS_PLC(void); //Reading FS PLC data asynchronous

HRESULT AddOPCItems_Read_Synch_PLC_ULS(void); //Creating OPC Items for ULS PLC Asynch_read
//HRESULT ReadAsynchData_ULS_PLC(void); //Reading ULS PLC data asynchronous
HRESULT ReadSynchData_ULS_PLC(void); //Reading ULS PLC data Synchronous

void Silo_Calculation(void); //Silo Stock calculation for Unloadind section
void Silo_Calculation_WF(void); //Silo Stock calculation for Feeding Section Weigh Feeders
void Ash_Calculation(void); // Ash calculation for Y37 and Y38
CString ValidatePath(CString);



CString m_strConnection_silo, m_strConnection_uls;
CString m_strCmdText_silo, m_strCmdText_uls;
_RecordsetPtr m_pRs_silo, m_pRs_uls;

CString GetSiloNumber_P2(void); //Return Silo Number for Path 2
CString GetSiloNumber_P4(void); //Return Silo Number for Path 4
CString GetSiloNumber_P6(void); //Return Silo Number for Path 6

float AddSiloStock(CString, float); //Adding Coal Stock to Selected Unloading Silo



