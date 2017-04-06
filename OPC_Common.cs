/*=====================================================================
  File:      OPC_Common.cs

  Summary:   OPC common custom interface

-----------------------------------------------------------------------
  This file is part of the Viscom OPC Code Samples.

  Copyright(c) 2001 Viscom (www.viscomvisual.com) All rights reserved.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
======================================================================*/

using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

namespace OPC.Common
{







// --------------------------- Opc Server lister (OPCEnum) ------------------------

	[StructLayout(LayoutKind.Sequential)]
public class OpcServers
	{
	public string	ProgID;
	public string	ServerName;
	public Guid		ClsID;

	public override string ToString()
		{
		StringBuilder sb = new StringBuilder( "OPCServer: ", 300 );
		sb.AppendFormat( "'{0}' ID={1} [{2}]", ServerName, ProgID, ClsID );
		return sb.ToString();
		}
	}


	[ComVisible(true)]
public class OpcServerList
	{
	public OpcServerList()
		{
		OPCListObj		= null;
		ifList			= null;
		EnumObj			= null;
		ifEnum			= null;
		}
	~OpcServerList()
		{ Dispose(); }

	public void ListAllData20( out OpcServers[] serverslist )
		{					// CATID_OPCDAServer20
		ListAll( new Guid( "63D5F432-CFE4-11d1-B2C8-0060083BA1FB" ), out serverslist );
		}

	public void ListAll( Guid catid, out OpcServers[] serverslist )
		{
		serverslist = null;
		Dispose();
		Guid	guid = new Guid( "13486D51-4821-11D2-A494-3CB306C10000" );
		Type	typeoflist = Type.GetTypeFromCLSID( guid );
		OPCListObj = Activator.CreateInstance( typeoflist );

		ifList = (IOPCServerList) OPCListObj;
		if( ifList == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );
			
		ifList.EnumClassesOfCategories( 1, ref catid, 0, ref catid, out EnumObj );
		if( EnumObj == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		ifEnum = (IEnumGUID) EnumObj;
		if( ifEnum == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );
		
		int		maxcount = 300;
		IntPtr	ptrGuid = Marshal.AllocCoTaskMem( maxcount * 16 );
		int		count = 0;
		ifEnum.Next( maxcount, ptrGuid, out count );
		if( count < 1 )
			{ Marshal.FreeCoTaskMem( ptrGuid ); return; }

		serverslist = new OpcServers[ count ];

		byte[]	guidbin = new byte[ 16 ];
		int	runGuid = (int) ptrGuid;
		for( int i = 0; i < count; i++ )
			{
			serverslist[i] = new OpcServers();
			Marshal.Copy( (IntPtr) runGuid, guidbin, 0, 16 );
			serverslist[i].ClsID = new Guid( guidbin );
			ifList.GetClassDetails( ref serverslist[i].ClsID,
									out serverslist[i].ProgID, out serverslist[i].ServerName );
			runGuid += 16;
			}

		Marshal.FreeCoTaskMem( ptrGuid );
		Dispose();
		}

	public void Dispose()
		{
		ifEnum = null;
		if( ! (EnumObj == null) )
			{
			int	rc = Marshal.ReleaseComObject( EnumObj );
			EnumObj = null;
			}
		ifList = null;
		if( ! (OPCListObj == null) )
			{
			int	rc = Marshal.ReleaseComObject( OPCListObj );
			OPCListObj = null;
			}
		}

	private object							OPCListObj;
	private IOPCServerList					ifList;

	private object							EnumObj;
	private IEnumGUID						ifEnum;
	}	// class OpcServerList












public class HRESULTS
	{
	public static bool Failed( int hresultcode )
		{	return (hresultcode < 0);	}

	public static bool Succeeded( int hresultcode )
		{	return (hresultcode >= 0);	}

	public const int S_OK		= 0x00000000;
	public const int S_FALSE	= 0x00000001;

	public const int E_NOTIMPL					= unchecked( (int)0x80004001 );		// winerror.h
	public const int E_NOINTERFACE				= unchecked( (int)0x80004002 );
	public const int E_ABORT					= unchecked( (int)0x80004004 );
	public const int E_FAIL						= unchecked( (int)0x80004005 );
	public const int E_OUTOFMEMORY				= unchecked( (int)0x8007000E );
	public const int E_INVALIDARG				= unchecked( (int)0x80070057 );
	
	public const int CONNECT_E_NOCONNECTION		= unchecked( (int)0x80040200 );		// olectl.h
	public const int CONNECT_E_ADVISELIMIT		= unchecked( (int)0x80040201 );

	public const int OPC_E_INVALIDHANDLE		= unchecked( (int)0xC0040001 );		// opcerror.h
	public const int OPC_E_BADTYPE				= unchecked( (int)0xC0040004 );
	public const int OPC_E_PUBLIC				= unchecked( (int)0xC0040005 );
	public const int OPC_E_BADRIGHTS			= unchecked( (int)0xC0040006 );
	public const int OPC_E_UNKNOWNITEMID		= unchecked( (int)0xC0040007 );
	public const int OPC_E_INVALIDITEMID		= unchecked( (int)0xC0040008 );
	public const int OPC_E_INVALIDFILTER		= unchecked( (int)0xC0040009 );
	public const int OPC_E_UNKNOWNPATH			= unchecked( (int)0xC004000A );
	public const int OPC_E_RANGE				= unchecked( (int)0xC004000B );
	public const int OPC_E_DUPLICATENAME		= unchecked( (int)0xC004000C );
	public const int OPC_S_UNSUPPORTEDRATE		= unchecked( (int)0x0004000D );
	public const int OPC_S_CLAMP				= unchecked( (int)0x0004000E );
	public const int OPC_S_INUSE				= unchecked( (int)0x0004000F );
	public const int OPC_E_INVALIDCONFIGFILE	= unchecked( (int)0xC0040010 );
	public const int OPC_E_NOTFOUND				= unchecked( (int)0xC0040011 );
	public const int OPC_E_INVALID_PID			= unchecked( (int)0xC0040203 );

	}	// class HRESULTS




// dummy VARIANT  (workaround)
	[ComVisible(true), StructLayout(LayoutKind.Sequential, Pack=2)]
public class DUMMY_VARIANT
	{
	[DllImport("oleaut32.dll")]
	public static extern void VariantInit( IntPtr addrofvariant );

	[DllImport("oleaut32.dll")]
	public static extern int VariantClear( IntPtr addrofvariant );

	public static string VarEnumToString( VarEnum vevt )
		{
		string	strvt = "";
		short	vtshort = (short) vevt;
		if( vtshort == VT_ILLEGAL )
			return "VT_ILLEGAL";
		
		if( (vtshort & VT_ARRAY) != 0 )
			strvt += "VT_ARRAY | ";

		if( (vtshort & VT_BYREF) != 0 )
			strvt += "VT_BYREF | ";

		if( (vtshort & VT_VECTOR) != 0 )
			strvt += "VT_VECTOR | ";

		VarEnum vtbase = (VarEnum) (vtshort & VT_TYPEMASK);
		strvt += vtbase.ToString();
		return strvt;
		}

	public static short VT_TYPEMASK	= 0x0fff;
	public static short VT_VECTOR		= 0x1000;
	public static short VT_ARRAY		= 0x2000;
	public static short VT_BYREF		= 0x4000;
	public static short VT_ILLEGAL	= unchecked( (short)0xffff );


	public static int ConstSize = 16;

	public short	vt;
	public short	r1;
	public short	r2;
	public short	r3;
	public int		v1;
	public int		v2;
	}	// class DUMMY_VARIANT










// ----------------------------------------------------------------- OPC common
	[ComVisible(true), ComImport,
	Guid("F31DFDE2-07B6-11d2-B2D8-0060083BA1FB"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCCommon
	{
	void SetLocaleID(
		[In]								int	dwLcid );

	void GetLocaleID(
		[Out]							out int	pdwLcid );

		[PreserveSig]
	int QueryAvailableLocaleIDs(
		[Out]							out int				pdwCount,
		[Out]							out	IntPtr			pdwLcid );

	void GetErrorString(
		[In]											int			dwError,
		[Out, MarshalAs(UnmanagedType.LPWStr) ]		out	string		ppString );

	void SetClientName(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string		szName );
	}







// ----------------------------------------------------------------- Common callback
	[ComVisible(true), ComImport,
	Guid("F31DFDE1-07B6-11d2-B2D8-0060083BA1FB"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCShutdown
	{
	void ShutdownRequest(
		[In, MarshalAs(UnmanagedType.LPWStr) ]		string	szReason );
	}



// ----------------------------------------------------------------- Server List enum
	[ComVisible(true), ComImport,
	Guid("13486D50-4821-11D2-A494-3CB306C10000"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IOPCServerList
	{
	void EnumClassesOfCategories(
		[In]											int		cImplemented,	// WARNING ONLY 1!!
		[In]										ref Guid	catidImpl,		// WARNING ONLY 1!!
		[In]											int		cRequired,		// WARNING ONLY 1!!
		[In]										ref Guid	catidReq,		// WARNING ONLY 1!!
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object	ppUnk );

	void GetClassDetails(
		[In]										ref Guid	clsid,
		[Out, MarshalAs(UnmanagedType.LPWStr) ]		out	string	ppszProgID,
		[Out, MarshalAs(UnmanagedType.LPWStr) ]		out	string	ppszUserType );

	void CLSIDFromProgID(
		[In, MarshalAs(UnmanagedType.LPWStr) ]			string	szProgId,
		[Out]										out Guid	clsid );
	}	// interface IOPCServerList



// ----------------------------------------------------------------- Enum GUIDs
	[ComVisible(true), ComImport,
	Guid("0002E000-0000-0000-C000-000000000046"),
	InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
internal interface IEnumGUID
	{
	void Next(
		[In]											int		celt,
		[In]											IntPtr	rgelt,				// ptr to Out-Values!!
		[Out]										out int		pceltFetched );

	void Skip(
		[In]											int		celt );

	void Reset();

	void Clone(
		[Out, MarshalAs(UnmanagedType.IUnknown) ]	out	object	ppUnk );

	}	// interface IEnumGUID


}	// namespace OPC.Common.Interface
