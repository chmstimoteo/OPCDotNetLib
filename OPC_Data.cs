/*=====================================================================
  File:      OPC_Data.cs

  Summary:   OPC DA custom interface

-----------------------------------------------------------------------
  This file is part of the Viscom OPC Code Samples.

  Copyright(c) 2001 Viscom (www.viscomvisual.com) All rights reserved.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
======================================================================*/

/*
Notes:
	An interface declared with ComImport can expose HRESULTs to C#,
	this is done by [PreserveSig]

	midl attribute 'pointer_unique' is simulated by passing an array[1]


*/

using System;
using System.Runtime.InteropServices;
using System.Reflection;

using OPC.Common;


namespace OPC.Data.Interface
{
public enum OPCDATASOURCE
	{
	OPC_DS_CACHE	= 1,
	OPC_DS_DEVICE	= 2
	}

public enum OPCBROWSETYPE
	{
	OPC_BRANCH	= 1,
	OPC_LEAF	= 2,
	OPC_FLAT	= 3
	}

public enum OPCNAMESPACETYPE
	{
	OPC_NS_HIERARCHIAL	= 1,
	OPC_NS_FLAT			= 2
	}

public enum OPCBROWSEDIRECTION
	{
	OPC_BROWSE_UP		= 1,
	OPC_BROWSE_DOWN		= 2,
	OPC_BROWSE_TO		= 3
	}

	[Flags]
public enum OPCACCESSRIGHTS
	{
	OPC_READABLE	= 1,
	OPC_WRITEABLE	= 2
	}

public enum OPCEUTYPE
	{
	OPC_NOENUM		= 0,
	OPC_ANALOG		= 1,
	OPC_ENUMERATED	= 2
	}

public enum OPCSERVERSTATE
	{
	OPC_STATUS_RUNNING		= 1,
	OPC_STATUS_FAILED		= 2,
	OPC_STATUS_NOCONFIG		= 3,
	OPC_STATUS_SUSPENDED	= 4,
	OPC_STATUS_TEST			= 5
	}

public enum OPCENUMSCOPE
	{
	OPC_ENUM_PRIVATE_CONNECTIONS	= 1,
	OPC_ENUM_PUBLIC_CONNECTIONS		= 2,
	OPC_ENUM_ALL_CONNECTIONS		= 3,
	OPC_ENUM_PRIVATE				= 4,
	OPC_ENUM_PUBLIC					= 5,
	OPC_ENUM_ALL					= 6
	}





//****************************************************
// OPC Quality flags
	[Flags]
public enum OPC_QUALITY_MASKS : short
	{
	LIMIT_MASK			= 0x0003,
	STATUS_MASK			= 0x00FC,
	MASTER_MASK			= 0x00C0,
	}

	[Flags]
public enum OPC_QUALITY_MASTER : short
	{
	QUALITY_BAD			= 0x0000,
	QUALITY_UNCERTAIN	= 0x0040,
	ERROR_QUALITY_VALUE	= 0x0080,		// non standard!
	QUALITY_GOOD		= 0x00C0,
	}

	[Flags]
public enum OPC_QUALITY_STATUS : short
	{
	BAD					= 0x0000,	// STATUS_MASK Values for Quality = BAD
	CONFIG_ERROR		= 0x0004,
	NOT_CONNECTED		= 0x0008,
	DEVICE_FAILURE		= 0x000c,
	SENSOR_FAILURE		= 0x0010,
	LAST_KNOWN			= 0x0014,
	COMM_FAILURE		= 0x0018,
	OUT_OF_SERVICE		= 0x001C,

	UNCERTAIN			= 0x0040,	// STATUS_MASK Values for Quality = UNCERTAIN
	LAST_USABLE			= 0x0044,
	SENSOR_CAL			= 0x0050,
	EGU_EXCEEDED		= 0x0054,
	SUB_NORMAL			= 0x0058,

	OK					= 0x00C0,	// STATUS_MASK Value for Quality = GOOD
	LOCAL_OVERRIDE		= 0x00D8
	}

	[Flags]
public enum OPC_QUALITY_LIMIT
	{
	LIMIT_OK			= 0x0000,
	LIMIT_LOW			= 0x0001,
	LIMIT_HIGH			= 0x0002,
	LIMIT_CONST			= 0x0003
	}


public enum OPC_PROPS
	{
	OPC_PROP_CDT			= 1,
	OPC_PROP_VALUE			= 2,
	OPC_PROP_QUALITY		= 3,
	OPC_PROP_TIME			= 4,
	OPC_PROP_RIGHTS			= 5,
	OPC_PROP_SCANRATE		= 6,

	OPC_PROP_UNIT			= 100,
	OPC_PROP_DESC			= 101,
	OPC_PROP_HIEU			= 102,
	OPC_PROP_LOEU			= 103,
	OPC_PROP_HIRANGE		= 104,
	OPC_PROP_LORANGE		= 105,
	OPC_PROP_CLOSE			= 106,
	OPC_PROP_OPEN			= 107,
	OPC_PROP_TIMEZONE		= 108,

	OPC_PROP_FGC			= 200,
	OPC_PROP_BGC			= 201,
	OPC_PROP_BLINK			= 202,
	OPC_PROP_BMP			= 203,
	OPC_PROP_SND			= 204,
	OPC_PROP_HTML			= 205,
	OPC_PROP_AVI			= 206,

	OPC_PROP_ALMSTAT		= 300,
	OPC_PROP_ALMHELP		= 301,
	OPC_PROP_ALMAREAS		= 302,
	OPC_PROP_ALMPRIMARYAREA	= 303,
	OPC_PROP_ALMCONDITION	= 304,
	OPC_PROP_ALMLIMIT		= 305,
	OPC_PROP_ALMDB			= 306,
	OPC_PROP_ALMHH			= 307,
	OPC_PROP_ALMH			= 308,
	OPC_PROP_ALML			= 309,
	OPC_PROP_ALMLL			= 310,
	OPC_PROP_ALMROC			= 311,
	OPC_PROP_ALMDEV			= 312
	}


// ------------------ SERVER level structs ------------------

	[StructLayout(LayoutKind.Sequential, Pack=2, CharSet=CharSet.Unicode)]
public class SERVERSTATUS
	{
	public long		ftStartTime;
	public long		ftCurrentTime;
	public long		ftLastUpdateTime;

		[MarshalAs(UnmanagedType.U4)]
	public OPCSERVERSTATE	eServerState;

	public int		dwGroupCount;
	public int		dwBandWidth;
	public short	wMajorVersion;
	public short	wMinorVersion;
	public short	wBuildNumber;
	public short	wReserved;

		[MarshalAs(UnmanagedType.LPWStr)]
	public string	szVendorInfo;
	};








// ------------------ INTERNAL item level structs ------------------

	[StructLayout(LayoutKind.Sequential, Pack=2, CharSet=CharSet.Unicode)]
internal class OPCITEMDEFintern
	{
		[MarshalAs(UnmanagedType.LPWStr)]
	public string	szAccessPath;

		[MarshalAs(UnmanagedType.LPWStr)]
	public string	szItemID;

		[MarshalAs(UnmanagedType.Bool)]
	public bool		bActive;

	public int		hClient;
	public int		dwBlobSize;
	public IntPtr	pBlob;

	public short	vtRequestedDataType;
	
	public short	wReserved;
	};




	[StructLayout(LayoutKind.Sequential, Pack=2)]
internal class OPCITEMRESULTintern
	{
	public int		hServer					= 0;
	public short	vtCanonicalDataType		= 0;
	public short	wReserved				= 0;

		[MarshalAs(UnmanagedType.U4)]
	public OPCACCESSRIGHTS	dwAccessRights	= 0;

	public int		dwBlobSize				= 0;
	public int		pBlob					= 0;
	};







// ----------------------------------------------------------------- SERVER
	[ComVisible(true), ComImport,
	Guid("39c13a4d-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCServer
	{
	void AddGroup(
		[In, MarshalAs(UnmanagedType.LPWStr) ]					string		szName,
		[In, MarshalAs(UnmanagedType.Bool) ]					bool		bActive,
		[In]													int			dwRequestedUpdateRate,
		[In]													int			hClientGroup,
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]		int[]		pTimeBias,
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]		float[]		pPercentDeadband,
		[In]													int			dwLCID,
		[Out]													out	int			phServerGroup,
		[Out]													out	int			pRevisedUpdateRate,
		[In]													ref Guid		riid,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]				out	object		ppUnk );

	void GetErrorString(
		[In]											int			dwError,
		[In]											int			dwLocale,
		[Out, MarshalAs(UnmanagedType.LPWStr) ]		out	string		ppString );

	void GetGroupByName(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string		szName,
		[In]										ref Guid		riid,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object		ppUnk );

	void GetStatus(
		[Out, MarshalAs(UnmanagedType.LPStruct) ]	out	SERVERSTATUS	ppServerStatus );

	void RemoveGroup(
		[In]											int			hServerGroup,
		[In, MarshalAs(UnmanagedType.Bool) ]			bool		bForce );

		[PreserveSig]
	int CreateGroupEnumerator(										// may return S_FALSE
		[In]											int			dwScope,
		[In]										ref Guid		riid,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object		ppUnk );

	}	// interface IOPCServer




// ----------------------------------------------------------------- Public Groups
	[ComVisible(true), ComImport,
	Guid("39c13a4e-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCServerPublicGroups
	{
	void GetPublicGroupByName(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string		szName,
		[In]										ref Guid		riid,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object		ppUnk );

	void RemovePublicGroup(
		[In]											int			hServerGroup,
		[In, MarshalAs(UnmanagedType.Bool) ]			bool		bForce );

	}

// ----------------------------------------------------------------- ServerAddressSpace Browsing
	[ComVisible(true), ComImport,
	Guid("39c13a4f-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCBrowseServerAddressSpace
	{
	void QueryOrganization(
		[Out, MarshalAs(UnmanagedType.U4) ]			out	OPCNAMESPACETYPE	pNameSpaceType );

	void ChangeBrowsePosition(
		[In,  MarshalAs(UnmanagedType.U4) ]				OPCBROWSEDIRECTION	dwBrowseDirection,
		[In,  MarshalAs(UnmanagedType.LPWStr) ]			string				szName );

		[PreserveSig]
	int BrowseOPCItemIDs(
		[In,  MarshalAs(UnmanagedType.U4) ]				OPCBROWSETYPE		dwBrowseFilterType,
		[In,  MarshalAs(UnmanagedType.LPWStr) ]			string				szFilterCriteria,
		[In,  MarshalAs(UnmanagedType.U2) ]				short				vtDataTypeFilter,
		[In,  MarshalAs(UnmanagedType.U4) ]				OPCACCESSRIGHTS		dwAccessRightsFilter,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object				ppUnk );

	void GetItemID(
		[In,  MarshalAs(UnmanagedType.LPWStr) ]			string				szItemDataID,
		[Out, MarshalAs(UnmanagedType.LPWStr) ]		out	string				szItemID );

		[PreserveSig]
	int BrowseAccessPaths(
		[In,  MarshalAs(UnmanagedType.LPWStr) ]			string				szItemID,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object				ppUnk );
	}


// ----------------------------------------------------------------- Item Properties
	[ComVisible(true), ComImport,
	Guid("39c13a72-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCItemProperties
	{
	void QueryAvailableProperties(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string		szItemID,
		[Out]										out int			dwCount,
		[Out]										out IntPtr		ppPropertyIDs,
		[Out]										out IntPtr		ppDescriptions,
		[Out]										out	IntPtr		ppvtDataTypes );

		[PreserveSig]
	int GetItemProperties(
		[In, MarshalAs(UnmanagedType.LPWStr) ]						string		szItemID,
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=1)]	int[]		pdwPropertyIDs,
		[Out]													out IntPtr		ppvData,
		[Out]													out	IntPtr		ppErrors );

		[PreserveSig]
	int LookupItemIDs(
		[In, MarshalAs(UnmanagedType.LPWStr) ]						string		szItemID,
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=1)]	int[]		pdwPropertyIDs,
		[Out]													out IntPtr		ppszNewItemIDs,
		[Out]													out	IntPtr		ppErrors );
	}






































// ----------------------------------------------------------------- GroupStateMgt
	[ComVisible(true),
	Guid("39c13a50-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCGroupStateMgt
	{
	void GetState(
		[Out]										out	int			pUpdateRate,
		[Out, MarshalAs(UnmanagedType.Bool) ]		out	bool		pActive,
		[Out, MarshalAs(UnmanagedType.LPWStr) ]		out	string		ppName,
		[Out]										out	int			pTimeBias,
		[Out]										out	float		pPercentDeadband,
		[Out]										out	int			pLCID,
		[Out]										out	int			phClientGroup,
		[Out]										out	int			phServerGroup );

	void SetState(
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]											int[]	pRequestedUpdateRate,
		[Out]																					out	int	pRevisedUpdateRate,
		[In, MarshalAs(UnmanagedType.LPArray, ArraySubType=UnmanagedType.Bool, SizeConst=1)]		bool[]	pActive,
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]											int[]	pTimeBias,
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]		float[]	pPercentDeadband,
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]		int[]	pLCID,
		[In, MarshalAs(UnmanagedType.LPArray, SizeConst=1)]		int[]	phClientGroup );

	void SetName(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string		szName );

	void CloneGroup(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string		szName,
		[In]										ref Guid		riid,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object		ppUnk );

	}	// interface IOPCGroupStateMgt





// ----------------------------------------------------------------- Public Group StateMgt
	[ComVisible(true),
	Guid("39c13a51-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCPublicGroupStateMgt
	{
	void GetState(
		[Out, MarshalAs(UnmanagedType.Bool) ]		out	bool		pPublic );

	void MoveToPublic();
	}







// ----------------------------------------------------------------- Item Mgmt
	[ComVisible(true), ComImport,
	Guid("39c13a54-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCItemMgt
	{
		[PreserveSig]
	int AddItems(
		[In]											int			dwCount,
		[In]											IntPtr		pItemArray,
		[Out]										out IntPtr		ppAddResults,
		[Out]										out	IntPtr		ppErrors );

		[PreserveSig]
	int ValidateItems(
		[In]											int			dwCount,
		[In]											IntPtr		pItemArray,
		[In, MarshalAs(UnmanagedType.Bool) ]			bool		bBlobUpdate,
		[Out]										out	IntPtr		ppValidationResults,
		[Out]										out	IntPtr		ppErrors );

		[PreserveSig]
	int RemoveItems(
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phServer,
		[Out]													out	IntPtr		ppErrors );

		[PreserveSig]
	int SetActiveState(
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phServer,
		[In, MarshalAs(UnmanagedType.Bool) ]						bool		bActive,
		[Out]													out	IntPtr		ppErrors );

		[PreserveSig]
	int SetClientHandles(
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phServer,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phClient,
		[Out]													out	IntPtr		ppErrors );

		[PreserveSig]
	int SetDatatypes(
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phServer,
		[In]														IntPtr		pRequestedDatatypes,
		[Out]													out	IntPtr		ppErrors );

		[PreserveSig]
	int CreateEnumerator(
		[In]										ref Guid		riid,
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object		ppUnk );

	}	// interface IOPCItemMgt



// ----------------------------------------------------------------- Sync IO
	[ComVisible(true), ComImport,
	Guid("39c13a52-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCSyncIO
	{
		[PreserveSig]
	int Read(
		[In, MarshalAs(UnmanagedType.U4) ]							OPCDATASOURCE	dwSource,
		[In]														int				dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=1)]	int[]			phServer,
		[Out]													out IntPtr			ppItemValues,
		[Out]													out	IntPtr			ppErrors );

		[PreserveSig]
	int Write(
		[In]														int				dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]			phServer,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	object[]		pItemValues,
		[Out]													out	IntPtr			ppErrors );

	}	// interface IOPCSyncIO



// ----------------------------------------------------------------- Async IO
	[ComVisible(true), ComImport,
	Guid("39c13a71-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCAsyncIO2
	{
		[PreserveSig]
	int Read(
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phServer,
		[In]														int			dwTransactionID,
		[Out]													out int			pdwCancelID,
		[Out]													out	IntPtr		ppErrors );

		[PreserveSig]
	int Write(
		[In]														int			dwCount,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	int[]		phServer,
		[In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=0)]	object[]	pItemValues,
		[In]														int			dwTransactionID,
		[Out]													out int			pdwCancelID,
		[Out]													out	IntPtr		ppErrors );

	void Refresh2(
		[In, MarshalAs(UnmanagedType.U4) ]				OPCDATASOURCE	dwSource,
		[In]											int				dwTransactionID,
		[Out]										out int				pdwCancelID );

	void Cancel2(
		[In]											int				dwCancelID );

	void SetEnable(
		[In, MarshalAs(UnmanagedType.Bool) ]			bool			bEnable );

	void GetEnable(
		[Out, MarshalAs(UnmanagedType.Bool) ]		out	bool			pbEnable );

	}	// interface IOPCAsyncIO2









// ----------------------------------------------------------------- Async Callback
	[ComVisible(true), ComImport,
	Guid("39c13a70-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCDataCallback
	{
	void OnDataChange(
		[In]											int				dwTransid,
		[In]											int				hGroup,
		[In]											int				hrMasterquality,
		[In]											int				hrMastererror,
		[In]											int				dwCount,
		[In]											IntPtr			phClientItems,
		[In]											IntPtr			pvValues,
		[In]											IntPtr			pwQualities,
		[In]											IntPtr			pftTimeStamps,
		[In]											IntPtr			ppErrors );

	void OnReadComplete(
		[In]											int				dwTransid,
		[In]											int				hGroup,
		[In]											int				hrMasterquality,
		[In]											int				hrMastererror,
		[In]											int				dwCount,
		[In]											IntPtr			phClientItems,
		[In]											IntPtr			pvValues,
		[In]											IntPtr			pwQualities,
		[In]											IntPtr			pftTimeStamps,
		[In]											IntPtr			ppErrors );
	
	void OnWriteComplete(
		[In]											int				dwTransid,
		[In]											int				hGroup,
		[In]											int				hrMastererr,
		[In]											int				dwCount,
		[In]											IntPtr			pClienthandles,
		[In]											IntPtr			ppErrors );

	void OnCancelComplete(
		[In]											int				dwTransid,
		[In]											int				hGroup );

	}	// interface IOPCDataCallback




// ----------------------------------------------------------------- Enum Item Attributes
	[ComVisible(true), ComImport,
	Guid("39c13a55-011e-11d0-9675-0020afd8adb3"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IEnumOPCItemAttributes
	{
	void Next(
		[In]											int		celt,
		[Out]										out	IntPtr	ppItemArray,
		[Out]										out int		pceltFetched );

	void Skip(
		[In]											int		celt );

	void Reset();

	void Clone(
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object		ppUnk );

	}	// interface IEnumOPCItemAttributes


}	// namespace OPC.Data.Interface
