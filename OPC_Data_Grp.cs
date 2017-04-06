/*=====================================================================
  File:      OPC_Data_Grp.cs

  Summary:   OPC DA group interfaces wrapper class

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
using System.Diagnostics;

using OPC.Common;
using OPC.Data.Interface;


namespace OPC.Data
{




// ------------- managed side only structs ----------------------


public class OPCItemDef
	{
	public OPCItemDef() { }
	public OPCItemDef( string id, bool activ, int hclt, VarEnum vt )
		{	ItemID=id; Active=activ; HandleClient=hclt; RequestedDataType = vt; }
		
	public string	AccessPath	= "";
	public string	ItemID;
	public bool		Active;
	public int		HandleClient;
	public byte[]	Blob			= null;
	public VarEnum	RequestedDataType;
	};

public class OPCItemResult
	{
	public int				Error;			// content below only valid if Error=S_OK
	public int				HandleServer;
	public VarEnum			CanonicalDataType;
	public OPCACCESSRIGHTS	AccessRights;
	public byte[]			Blob;
	}

public class OPCItemState
	{
	public int		Error;			// content below only valid if Error=S_OK
	public int		HandleClient;	// always valid for callbacks
	public object	DataValue;
	public long		TimeStamp;
	public short	Quality;

	public override string ToString()
		{
		StringBuilder sb = new StringBuilder( "OPCIST: ", 256 );
		sb.AppendFormat( "error=0x{0:x} hclt=0x{1:x}", Error, HandleClient );
		if( Error == HRESULTS.S_OK )
			{
			sb.AppendFormat( " val={0} time={1} qual=", DataValue, TimeStamp );
			sb.Append( OpcGroup.QualityToString( Quality ) );
			}
		
		return sb.ToString();
		}
	}



public class OPCWriteResult
	{
	public int		Error;
	public int		HandleClient;
	}


public class OPCItemAttributes
	{
	public string			AccessPath;
	public string			ItemID;
	public bool				Active;
	public int				HandleClient;
	public int				HandleServer;
	public OPCACCESSRIGHTS	AccessRights;
	public VarEnum			RequestedDataType;
	public VarEnum			CanonicalDataType;
	public OPCEUTYPE		EUType;
	public object			EUInfo;
	public byte[]			Blob;

	public override string ToString()
		{
		StringBuilder sb = new StringBuilder( "OPCIAT: '", 512 );
		sb.Append( ItemID );	sb.Append( "' ('" );	sb.Append( AccessPath );
		sb.AppendFormat( "') hc=0x{0:x} hs=0x{1:x} act={2}", HandleClient, HandleServer, Active );
		sb.AppendFormat( "\r\n\tacc={0} typr={1} typc={2}", AccessRights, RequestedDataType, CanonicalDataType );
		sb.AppendFormat( "\r\n\teut={0} eui={1}", EUType, EUInfo );
		if( ! (Blob == null) )
			sb.AppendFormat( " blob size={0}", Blob.Length );
		
		return sb.ToString();
		}
	}



public struct OPCGroupState
	{
	public string	Name;
	public bool		Public;
	public int		UpdateRate;
	public bool		Active;
	public int		TimeBias;
	public float	PercentDeadband;
	public int		LocaleID;
	public int		HandleClient;
	public int		HandleServer;
	}





// ----------------- event arguments + handlers ------------------------

						// IOPCAsyncIO2
						
public class DataChangeEventArgs : EventArgs
	{
	public DataChangeEventArgs()
		{
		}

	public int				transactionID;
	public int				groupHandleClient;
	public int				masterQuality;
	public int				masterError;

	public OPCItemState[]	sts;
	}
public delegate void DataChangeEventHandler( object sender, DataChangeEventArgs e );


public class ReadCompleteEventArgs : EventArgs
	{
	public ReadCompleteEventArgs()
		{
		}

	public int				transactionID;
	public int				groupHandleClient;
	public int				masterQuality;
	public int				masterError;

	public OPCItemState[]	sts;
	}
public delegate void ReadCompleteEventHandler( object sender, ReadCompleteEventArgs e );

public class WriteCompleteEventArgs : EventArgs
	{
	public WriteCompleteEventArgs()
		{
		}

	public int				transactionID;
	public int				groupHandleClient;
	public int				masterError;

	public OPCWriteResult[]	res;
	}
public delegate void WriteCompleteEventHandler( object sender, WriteCompleteEventArgs e );


public class CancelCompleteEventArgs : EventArgs
	{
	public CancelCompleteEventArgs( int transactionIDp, int groupHandleClientp )
		{
		transactionID		= transactionID;
		groupHandleClient	= groupHandleClientp;
		}

	public int				transactionID;
	public int				groupHandleClient;
	}
public delegate void CancelCompleteEventHandler( object sender, CancelCompleteEventArgs e );











// ----------------- class OpcGroup ------------------------

public class OpcGroup : IOPCDataCallback
	{
	internal OpcGroup( ref IOPCServer ifServerLink, bool isPublic, string groupName, bool setActive, int requestedUpdateRate )
		{
		ifServer				= ifServerLink;

		state.Name				= groupName;
		state.Public			= isPublic;
		state.UpdateRate		= requestedUpdateRate;
		state.Active			= setActive;
		state.TimeBias			= 0;
		state.PercentDeadband	= 0.0f;
		state.LocaleID			= 0;
		state.HandleClient		= this.GetHashCode();
		state.HandleServer		= 0;

		// marshaling helpers:
		typeOPCITEMDEF			= typeof( OPCITEMDEFintern );
		sizeOPCITEMDEF			= Marshal.SizeOf( typeOPCITEMDEF );
		typeOPCITEMRESULT		= typeof( OPCITEMRESULTintern );
		sizeOPCITEMRESULT		= Marshal.SizeOf( typeOPCITEMRESULT );
		}

	~OpcGroup()
		{ Remove( false ); }


	internal void internalAdd( int[] biasTime, float[] percentDeadband, int localeID )
		{
		Type	typGrpMgt = typeof( IOPCGroupStateMgt );
		Guid	guidGrpTst = typGrpMgt.GUID;

		object	objtemp;
		if( state.Public )
			{
			IOPCServerPublicGroups	ifPubGrps = null;
			ifPubGrps = (IOPCServerPublicGroups)	ifServer;
			if( ifPubGrps == null )
				Marshal.ThrowExceptionForHR( HRESULTS.E_NOINTERFACE );

			ifPubGrps.GetPublicGroupByName( state.Name, ref guidGrpTst, out objtemp );
			ifPubGrps = null;
			}
		else
			{
			ifServer.AddGroup(	state.Name, state.Active, state.UpdateRate, state.HandleClient, biasTime, percentDeadband, state.LocaleID,
								out state.HandleServer, out state.UpdateRate, ref guidGrpTst, out objtemp );
			}
		if( objtemp == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_NOINTERFACE );

		ifMgt = (IOPCGroupStateMgt) objtemp;
		objtemp = null;
		GetStates();
		
		getinterfaces();
		AdviseIOPCDataCallback();
		}

	public void Remove( bool bForce )
		{
		if( ! (callbackcpoint == null) )
			{
			if( callbackcookie != 0 )
				{
				callbackcpoint.Unadvise( callbackcookie );
				callbackcookie = 0;
				}
			int	rc = Marshal.ReleaseComObject( callbackcpoint );
			callbackcpoint = null;
			}

		cpointcontainer	= null;
		ifItems		= null;
		ifSync		= null;
		ifAsync		= null;

		if( ! (ifMgt == null) )
			{
			int	rc = Marshal.ReleaseComObject( ifMgt );
			ifMgt = null;
			}

		if( ! (ifServer == null) )
			{
			if( ! state.Public )
				ifServer.RemoveGroup( state.HandleServer, bForce );
			ifServer	= null;
			}

		state.HandleServer = 0;
		}


	// -------------- IOPCServerPublicGroups + IOPCPublicGroupStateMgt

	public void DeletePublic( bool bForce )
		{
		if( ! state.Public )
			Marshal.ThrowExceptionForHR( HRESULTS.E_FAIL );

		IOPCServerPublicGroups	ifPubGrps = null;
		ifPubGrps = (IOPCServerPublicGroups)	ifServer;
		if( ifPubGrps == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_NOINTERFACE );
		int	serverhandle = state.HandleServer;
		Remove( false );
		ifPubGrps.RemovePublicGroup( serverhandle, bForce );
		ifPubGrps = null;
		}

	public void MoveToPublic()
		{
		if( state.Public )
			Marshal.ThrowExceptionForHR( HRESULTS.E_FAIL );

		IOPCPublicGroupStateMgt		ifPubMgt = null;
		ifPubMgt = (IOPCPublicGroupStateMgt)	ifMgt;
		if( ifPubMgt == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_NOINTERFACE );
		ifPubMgt.MoveToPublic();
		ifPubMgt.GetState( out state.Public );
		ifPubMgt = null;
		}



	// -------------------------- IOPCGroupStateMgt
	
	public void SetName( string newName )
		{
		ifMgt.SetName( newName );
		state.Name = newName;
		}

	public void GetStates()		// like a refresh
		{
		ifMgt.GetState( out	state.UpdateRate, out state.Active, out state.Name, out state.TimeBias, out state.PercentDeadband,
						out state.LocaleID, out state.HandleClient, out state.HandleServer );
		}

	public string Name
		{
		get	{ return state.Name; }
		set	{
			SetName( value );
			}
		}

	public bool Active
		{
		get	{ return state.Active; }
		set	{
			ifMgt.SetState( null, out state.UpdateRate, new bool[1] { value }, null, null, null, null );
			state.Active = value;
			}
		}

	public bool Public
		{
		get	{ return state.Public; }
		}

	public int UpdateRate
		{
		get	{ return state.UpdateRate; }
		set	{
			ifMgt.SetState( new int[1] { value }, out state.UpdateRate, null, null, null, null, null );
			}
		}

	public int TimeBias
		{
		get	{ return state.TimeBias; }
		set	{
			ifMgt.SetState( null, out state.UpdateRate, null, new int[1] { value }, null, null, null );
			state.TimeBias = value;
			}
		}

	public float PercentDeadband
		{
		get	{ return state.PercentDeadband; }
		set	{
			ifMgt.SetState( null, out state.UpdateRate, null, null, new float[1] { value }, null, null );
			state.PercentDeadband = value;
			}
		}

	public int LocaleID
		{
		get	{ return state.LocaleID; }
		set	{
			ifMgt.SetState( null, out state.UpdateRate, null, null, null, new int[1] { value }, null );
			state.LocaleID = value;
			}
		}

	public int HandleClient
		{
		get	{ return state.HandleClient; }
		set	{
			ifMgt.SetState( null, out state.UpdateRate, null, null, null, null, new int[1] { value } );
			state.HandleClient = value;
			}
		}

	public int HandleServer
		{
		get	{ return state.HandleServer; }
		}






	// ------------------------ IOPCItemMgt ---------------
	
	public bool AddItems(	OPCItemDef[] arrDef, out OPCItemResult[] arrRes )
		{
		arrRes			= null;
		bool hasblobs	= false;
		int	count		= arrDef.Length;

		IntPtr	ptrDef = Marshal.AllocCoTaskMem( count * sizeOPCITEMDEF );
		int	runDef = (int) ptrDef;
		OPCITEMDEFintern	idf = new OPCITEMDEFintern();
		idf.wReserved = 0;
		foreach( OPCItemDef d in arrDef )
			{
			idf.szAccessPath		= d.AccessPath;
			idf.szItemID			= d.ItemID;
			idf.bActive				= d.Active;
			idf.hClient				= d.HandleClient;
			idf.vtRequestedDataType	= (short) d.RequestedDataType;
			idf.dwBlobSize = 0;	idf.pBlob = IntPtr.Zero;
			if( d.Blob != null )
				{
				idf.dwBlobSize = d.Blob.Length;
				if( idf.dwBlobSize > 0 )
					{
					hasblobs = true;
					idf.pBlob = Marshal.AllocCoTaskMem( idf.dwBlobSize );
					Marshal.Copy( d.Blob, 0, idf.pBlob, idf.dwBlobSize );
					}
				}

			Marshal.StructureToPtr( idf, (IntPtr) runDef, false );
			runDef += sizeOPCITEMDEF;
			}

		IntPtr	ptrRes;
		IntPtr	ptrErr;
		int	hresult = ifItems.AddItems( count, ptrDef, out ptrRes, out ptrErr );

		runDef = (int) ptrDef;
		if( hasblobs )
			{
			for( int i = 0; i < count; i++ )
				{
				IntPtr blob = (IntPtr) Marshal.ReadInt32( (IntPtr)(runDef + 20) );
				if( blob != IntPtr.Zero )
					Marshal.FreeCoTaskMem( blob );
				Marshal.DestroyStructure( (IntPtr) runDef, typeOPCITEMDEF );
				runDef += sizeOPCITEMDEF;
				}
			}
		else
			{
			for( int i = 0; i < count; i++ )
				{
				Marshal.DestroyStructure( (IntPtr) runDef, typeOPCITEMDEF );
				runDef += sizeOPCITEMDEF;
				}
			}
		Marshal.FreeCoTaskMem( ptrDef );

		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		int	runRes = (int) ptrRes;
		int	runErr = (int) ptrErr;
		if( (runRes == 0) || (runErr == 0) )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		arrRes		= new OPCItemResult[ count ];
		for( int i = 0; i < count; i++ )
			{
			arrRes[i] = new OPCItemResult();
			arrRes[i].Error = Marshal.ReadInt32( (IntPtr) runErr );
			if( HRESULTS.Failed( arrRes[i].Error ) )
				continue;
				
			arrRes[i].HandleServer		= Marshal.ReadInt32( (IntPtr) runRes );
			arrRes[i].CanonicalDataType	= (VarEnum)(int) Marshal.ReadInt16( (IntPtr)(runRes + 4) );
			arrRes[i].AccessRights		= (OPCACCESSRIGHTS) Marshal.ReadInt32( (IntPtr)(runRes + 8) );

			int ptrblob	= Marshal.ReadInt32( (IntPtr)(runRes + 16) );
			if( (ptrblob != 0) )
				{
				int blobsize	= Marshal.ReadInt32( (IntPtr)(runRes + 12) );
				if( blobsize > 0 )
					{
					arrRes[i].Blob = new byte[ blobsize ];
					Marshal.Copy( (IntPtr) ptrblob, arrRes[i].Blob, 0, blobsize );
					}
				Marshal.FreeCoTaskMem( (IntPtr) ptrblob );
				}

			runRes += sizeOPCITEMRESULT;
			runErr += 4;
			}

		Marshal.FreeCoTaskMem( ptrRes );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	// -----------------------------------------------------------------------------------

	public bool ValidateItems(	OPCItemDef[] arrDef, bool blobUpd,
								out OPCItemResult[] arrRes )
		{
		arrRes			= null;
		bool hasblobs	= false;
		int	count		= arrDef.Length;

		IntPtr	ptrDef = Marshal.AllocCoTaskMem( count * sizeOPCITEMDEF );
		int	runDef = (int) ptrDef;
		OPCITEMDEFintern	idf = new OPCITEMDEFintern();
		idf.wReserved = 0;
		foreach( OPCItemDef d in arrDef )
			{
			idf.szAccessPath		= d.AccessPath;
			idf.szItemID			= d.ItemID;
			idf.bActive				= d.Active;
			idf.hClient				= d.HandleClient;
			idf.vtRequestedDataType	= (short) d.RequestedDataType;
			idf.dwBlobSize = 0;	idf.pBlob = IntPtr.Zero;
			if( d.Blob != null )
				{
				idf.dwBlobSize = d.Blob.Length;
				if( idf.dwBlobSize > 0 )
					{
					hasblobs = true;
					idf.pBlob = Marshal.AllocCoTaskMem( idf.dwBlobSize );
					Marshal.Copy( d.Blob, 0, idf.pBlob, idf.dwBlobSize );
					}
				}

			Marshal.StructureToPtr( idf, (IntPtr) runDef, false );
			runDef += sizeOPCITEMDEF;
			}

		IntPtr	ptrRes;
		IntPtr	ptrErr;
		int	hresult = ifItems.ValidateItems( count, ptrDef, blobUpd, out ptrRes, out ptrErr );

		runDef = (int) ptrDef;
		if( hasblobs )
			{
			for( int i = 0; i < count; i++ )
				{
				IntPtr blob = (IntPtr) Marshal.ReadInt32( (IntPtr)(runDef + 20) );
				if( blob != IntPtr.Zero )
					Marshal.FreeCoTaskMem( blob );
				Marshal.DestroyStructure( (IntPtr) runDef, typeOPCITEMDEF );
				runDef += sizeOPCITEMDEF;
				}
			}
		else
			{
			for( int i = 0; i < count; i++ )
				{
				Marshal.DestroyStructure( (IntPtr) runDef, typeOPCITEMDEF );
				runDef += sizeOPCITEMDEF;
				}
			}
		Marshal.FreeCoTaskMem( ptrDef );

		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		int	runRes = (int) ptrRes;
		int	runErr = (int) ptrErr;
		if( (runRes == 0) || (runErr == 0) )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		arrRes		= new OPCItemResult[ count ];
		for( int i = 0; i < count; i++ )
			{
			arrRes[i] = new OPCItemResult();
			arrRes[i].Error = Marshal.ReadInt32( (IntPtr) runErr );
			if( HRESULTS.Failed( arrRes[i].Error ) )
				continue;
				
			arrRes[i].HandleServer		= Marshal.ReadInt32( (IntPtr) runRes );
			arrRes[i].CanonicalDataType	= (VarEnum)(int) Marshal.ReadInt16( (IntPtr)(runRes + 4) );
			arrRes[i].AccessRights		= (OPCACCESSRIGHTS) Marshal.ReadInt32( (IntPtr)(runRes + 8) );

			int ptrblob	= Marshal.ReadInt32( (IntPtr)(runRes + 16) );
			if( (ptrblob != 0) )
				{
				int blobsize	= Marshal.ReadInt32( (IntPtr)(runRes + 12) );
				if( blobsize > 0 )
					{
					arrRes[i].Blob = new byte[ blobsize ];
					Marshal.Copy( (IntPtr) ptrblob, arrRes[i].Blob, 0, blobsize );
					}
				Marshal.FreeCoTaskMem( (IntPtr) ptrblob );
				}

			runRes += sizeOPCITEMRESULT;
			runErr += 4;
			}

		Marshal.FreeCoTaskMem( ptrRes );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	// -----------------------------------------------------------------------------------
	
	public bool RemoveItems( int[] arrHSrv, out int[] arrErr )
		{
		arrErr		= null;
		int	count	= arrHSrv.Length;
		IntPtr		ptrErr;
		int	hresult = ifItems.RemoveItems( count, arrHSrv, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	// -----------------------------------------------------------------------------------
	
	public bool SetActiveState( int[] arrHSrv, bool activate, out int[] arrErr )
		{
		arrErr		= null;
		int	count	= arrHSrv.Length;
		IntPtr		ptrErr;
		int	hresult = ifItems.SetActiveState( count, arrHSrv, activate, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	// -----------------------------------------------------------------------------------
	
	public bool SetClientHandles( int[] arrHSrv, int[] arrHClt, out int[] arrErr )
		{
		arrErr		= null;
		int	count	= arrHSrv.Length;
		if( count != arrHClt.Length )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		IntPtr	ptrErr;
		int	hresult = ifItems.SetClientHandles( count, arrHSrv, arrHClt, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}



	// -----------------------------------------------------------------------------------
	
	public bool SetDatatypes( int[] arrHSrv, VarEnum[] arrVT, out int[] arrErr )
		{
		arrErr		= null;
		int	count	= arrHSrv.Length;
		if( count != arrVT.Length )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		IntPtr ptrVT = Marshal.AllocCoTaskMem( count * 2 );
		int	runVT = (int) ptrVT;
		foreach( VarEnum v in arrVT )
			{
			Marshal.WriteInt16( (IntPtr) runVT, (short) v );
			runVT += 2;
			}

		IntPtr	ptrErr;
		int	hresult = ifItems.SetDatatypes( count, arrHSrv, ptrVT, out ptrErr );

		Marshal.FreeCoTaskMem( ptrVT );

		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	// -----------------------------------------------------------------------------------
	
	public OpcEnumItemAttributes CreateAttrEnumerator()
		{
		Type	typEnuAtt = typeof( IEnumOPCItemAttributes );
		Guid	guidEnuAtt = typEnuAtt.GUID;
		object	objtemp;

		int	hresult = ifItems.CreateEnumerator( ref guidEnuAtt, out objtemp );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );
		if( (hresult == HRESULTS.S_FALSE) || (objtemp == null) )
			return null;

		IEnumOPCItemAttributes	ifenu = (IEnumOPCItemAttributes) objtemp;
		objtemp = null;

		OpcEnumItemAttributes	enu = new OpcEnumItemAttributes( ifenu );
		return enu;
		}











	// ------------------------ IOPCSyncIO ---------------

	public bool Read( OPCDATASOURCE src, int[] arrHSrv, out OPCItemState[] arrStat )
		{
		arrStat		= null;
		int			count = arrHSrv.Length;
		IntPtr		ptrStat;
		IntPtr		ptrErr;
		int	hresult = ifSync.Read( src, count, arrHSrv, out ptrStat, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		int	runErr	= (int) ptrErr;
		int	runStat	= (int) ptrStat;
		if( (runErr == 0) || (runStat == 0) )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		arrStat		= new OPCItemState[ count ];
		for( int i = 0; i < count; i++ )
			{														// WORKAROUND !!!
			arrStat[i]				= new OPCItemState();

			arrStat[i].Error		= Marshal.ReadInt32( (IntPtr) runErr );
			runErr += 4;

			arrStat[i].HandleClient	= Marshal.ReadInt32( (IntPtr) runStat );

			if( HRESULTS.Succeeded( arrStat[i].Error ) )
				{
				short vt = Marshal.ReadInt16( (IntPtr) (runStat + 16) );
				if( vt == (short) VarEnum.VT_ERROR )
					arrStat[i].Error = Marshal.ReadInt32( (IntPtr) (runStat + 24) );

				arrStat[i].TimeStamp	= Marshal.ReadInt64( (IntPtr) (runStat + 4) );
				arrStat[i].Quality		= Marshal.ReadInt16( (IntPtr) (runStat + 12) );
				arrStat[i].DataValue	= Marshal.GetObjectForNativeVariant( (IntPtr) (runStat + 16) );
				DUMMY_VARIANT.VariantClear( (IntPtr) (runStat + 16) );
				}
			else
				arrStat[i].DataValue = null;

			runStat += 32;
			}

		Marshal.FreeCoTaskMem( ptrStat );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	public bool Write( int[] arrHSrv, object[] arrVal, out int[] arrErr )
		{
		arrErr		= null;
		int	count	= arrHSrv.Length;
		if( count != arrVal.Length )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		IntPtr	ptrErr;
		int	hresult = ifSync.Write( count, arrHSrv, arrVal, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


















	// ------------------------ IOPCAsyncIO2 ---------------

	public bool Read( int[] arrHSrv, int transactionID, out int cancelID, out int[] arrErr )
		{
		arrErr		= null;
		cancelID	= 0;
		int	count	= arrHSrv.Length;

		IntPtr	ptrErr;
		int	hresult = ifAsync.Read( count, arrHSrv, transactionID, out cancelID, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}

	public bool Write(	int[] arrHSrv, object[] arrVal, int transactionID,
						out int cancelID, out int[] arrErr )
		{
		arrErr		= null;
		cancelID	= 0;
		int	count	= arrHSrv.Length;
		if( count != arrVal.Length )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		IntPtr	ptrErr;
		int	hresult = ifAsync.Write( count, arrHSrv, arrVal, transactionID, out cancelID, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		arrErr		= new int[ count ];
		Marshal.Copy( ptrErr, arrErr, 0, count );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}

	public void Refresh2( OPCDATASOURCE	sourceMode, int transactionID, out int cancelID )
		{
		ifAsync.Refresh2( sourceMode, transactionID, out cancelID );
		}
	public void Cancel2( int cancelID )
		{
		ifAsync.Cancel2( cancelID );
		}
	public void SetEnable( bool doEnable )
		{
		ifAsync.SetEnable( doEnable );
		}
	public void GetEnable( out bool isEnabled )
		{
		ifAsync.GetEnable( out isEnabled );
		}



	// ------------------------ IOPCDataCallback ---------------

	void IOPCDataCallback.OnDataChange(
			int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount,
			IntPtr phClientItems, IntPtr pvValues, IntPtr pwQualities, IntPtr pftTimeStamps, IntPtr ppErrors )
		{
		Trace.WriteLine( "OpcGroup.OnDataChange" );
		if( (dwCount == 0) || (hGroup != state.HandleClient) )
			return;
		int count = (int) dwCount;

		int	runh = (int) phClientItems;
		int	runv = (int) pvValues;
		int	runq = (int) pwQualities;
		int	runt = (int) pftTimeStamps;
		int	rune = (int) ppErrors;

		DataChangeEventArgs	e = new DataChangeEventArgs();
		e.transactionID		= dwTransid;
		e.groupHandleClient	= hGroup;
		e.masterQuality		= hrMasterquality;
		e.masterError		= hrMastererror;
		e.sts				= new OPCItemState[ count ];

		for( int i = 0; i < count; i++ )
			{
			e.sts[i] = new OPCItemState();
			e.sts[i].Error = Marshal.ReadInt32( (IntPtr) rune );
			rune += 4;

			e.sts[i].HandleClient = Marshal.ReadInt32( (IntPtr) runh );
			runh += 4;

			if( HRESULTS.Succeeded( e.sts[i].Error ) )
				{
				short vt = Marshal.ReadInt16( (IntPtr) runv );
				if( vt == (short) VarEnum.VT_ERROR )
					e.sts[i].Error = Marshal.ReadInt32( (IntPtr) (runv + 8) );

				e.sts[i].DataValue	= Marshal.GetObjectForNativeVariant( (IntPtr) runv );
				e.sts[i].Quality = Marshal.ReadInt16( (IntPtr) runq );
				e.sts[i].TimeStamp = Marshal.ReadInt64( (IntPtr) runt );
				}

			runv += DUMMY_VARIANT.ConstSize;
			runq += 2;
			runt += 8;
			}

		if( DataChanged != null )
			DataChanged( this, e );
		}

	void IOPCDataCallback.OnReadComplete(
			int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount,
			IntPtr phClientItems, IntPtr pvValues, IntPtr pwQualities, IntPtr pftTimeStamps, IntPtr ppErrors )
		{
		Trace.WriteLine( "OpcGroup.OnReadComplete" );
		if( (dwCount == 0) || (hGroup != state.HandleClient) )
			return;
		int count = (int) dwCount;

		int	runh = (int) phClientItems;
		int	runv = (int) pvValues;
		int	runq = (int) pwQualities;
		int	runt = (int) pftTimeStamps;
		int	rune = (int) ppErrors;

		ReadCompleteEventArgs	e = new ReadCompleteEventArgs();
		e.transactionID		= dwTransid;
		e.groupHandleClient	= hGroup;
		e.masterQuality		= hrMasterquality;
		e.masterError		= hrMastererror;
		e.sts				= new OPCItemState[ count ];

		for( int i = 0; i < count; i++ )
			{
			e.sts[i] = new OPCItemState();
			e.sts[i].Error = Marshal.ReadInt32( (IntPtr) rune );
			rune += 4;

			e.sts[i].HandleClient = Marshal.ReadInt32( (IntPtr) runh );
			runh += 4;

			if( HRESULTS.Succeeded( e.sts[i].Error ) )
				{
				short vt = Marshal.ReadInt16( (IntPtr) runv );
				if( vt == (short) VarEnum.VT_ERROR )
					e.sts[i].Error = Marshal.ReadInt32( (IntPtr) (runv + 8) );

				e.sts[i].DataValue	= Marshal.GetObjectForNativeVariant( (IntPtr) runv );
				e.sts[i].Quality = Marshal.ReadInt16( (IntPtr) runq );
				e.sts[i].TimeStamp = Marshal.ReadInt64( (IntPtr) runt );
				}

			runv += DUMMY_VARIANT.ConstSize;
			runq += 2;
			runt += 8;
			}

		if( ReadCompleted != null )
			ReadCompleted( this, e );
		}

	void IOPCDataCallback.OnWriteComplete(
			int dwTransid, int hGroup, int hrMastererr, int dwCount,
			IntPtr pClienthandles, IntPtr ppErrors )
		{
		Trace.WriteLine( "OpcGroup.OnWriteComplete" );
		if( (dwCount == 0) || (hGroup != state.HandleClient) )
			return;
		int count = (int) dwCount;

		int	runh = (int) pClienthandles;
		int	rune = (int) ppErrors;

		WriteCompleteEventArgs	e = new WriteCompleteEventArgs();
		e.transactionID		= dwTransid;
		e.groupHandleClient	= hGroup;
		e.masterError		= hrMastererr;
		e.res				= new OPCWriteResult[ count ];

		for( int i = 0; i < count; i++ )
			{
			e.res[i] = new OPCWriteResult();

			e.res[i].Error = Marshal.ReadInt32( (IntPtr) rune );
			rune += 4;

			e.res[i].HandleClient = Marshal.ReadInt32( (IntPtr) runh );
			runh += 4;
			}

		if( WriteCompleted != null )
			WriteCompleted( this, e );
		}
	
	void IOPCDataCallback.OnCancelComplete( int dwTransid, int hGroup )
		{
		Trace.WriteLine( "OpcGroup.OnCancelComplete" );
		if( hGroup != state.HandleClient )
			return;

		CancelCompleteEventArgs	e = new CancelCompleteEventArgs( dwTransid, hGroup );
		if( CancelCompleted != null )
			CancelCompleted( this, e );
		}




	// -------------------------- events ---------------------

	public event DataChangeEventHandler		DataChanged;
	public event ReadCompleteEventHandler	ReadCompleted;
	public event WriteCompleteEventHandler	WriteCompleted;
	public event CancelCompleteEventHandler	CancelCompleted;




	// -------------------------- helper ---------------------

	public static string QualityToString( short	Quality )
		{
		StringBuilder sb = new StringBuilder( 256 );
		OPC_QUALITY_MASTER		oqm = (OPC_QUALITY_MASTER)	(Quality & (short)OPC_QUALITY_MASKS.MASTER_MASK);
		OPC_QUALITY_STATUS		oqs = (OPC_QUALITY_STATUS)	(Quality & (short)OPC_QUALITY_MASKS.STATUS_MASK);
		OPC_QUALITY_LIMIT		oql = (OPC_QUALITY_LIMIT)	(Quality & (short)OPC_QUALITY_MASKS.LIMIT_MASK);
		sb.AppendFormat( "{0}+{1}+{2}", oqm, oqs, oql );
		return sb.ToString();
		}




	// -------------------------- private ---------------------

	private	void getinterfaces()
		{
		ifItems	= (IOPCItemMgt) ifMgt;
		ifSync	= (IOPCSyncIO) ifMgt;
		ifAsync	= (IOPCAsyncIO2) ifMgt;

		cpointcontainer	= (UCOMIConnectionPointContainer)	ifMgt;
		}

	private void AdviseIOPCDataCallback()
		{
		Type	sinktype = typeof( IOPCDataCallback );
		Guid	sinkguid = sinktype.GUID;
		
		cpointcontainer.FindConnectionPoint( ref sinkguid, out callbackcpoint );
		if( callbackcpoint == null )
			return;

		callbackcpoint.Advise( this, out callbackcookie );
		}



	private	OPCGroupState		state;

	private	IOPCServer			ifServer	= null;
	private	IOPCGroupStateMgt	ifMgt		= null;
	private	IOPCItemMgt			ifItems		= null;
	private IOPCSyncIO			ifSync		= null;
	private IOPCAsyncIO2		ifAsync		= null;

	private	UCOMIConnectionPointContainer	cpointcontainer	= null;
	private	UCOMIConnectionPoint			callbackcpoint	= null;
	private int								callbackcookie	= 0;

	// marshaling helpers:
	private readonly Type		typeOPCITEMDEF;
	private readonly int		sizeOPCITEMDEF;
	private readonly Type		typeOPCITEMRESULT;
	private readonly int		sizeOPCITEMRESULT;
	}	// class OpcGroup



// ----------------- class OpcEnumItemAttributes ------------------------

public class OpcEnumItemAttributes
	{
	internal OpcEnumItemAttributes( IEnumOPCItemAttributes ifEnump )
		{
		ifEnum = ifEnump;
		}
	~OpcEnumItemAttributes()
		{ Dispose(); }

	public void Dispose()
		{
		if( ! (ifEnum == null) )
			{
			int	rc = Marshal.ReleaseComObject( ifEnum );
			ifEnum = null;
			}
		}

	public void Next( int enumcountmax, out OPCItemAttributes[] attributes )
		{
		attributes = null;

		IntPtr	ptrAtt;
		int		count;
		ifEnum.Next( enumcountmax, out ptrAtt, out count );
		int	runatt = (int) ptrAtt;
		if( (runatt == 0) || (count <= 0) || (count > enumcountmax) )
			return;

		attributes = new OPCItemAttributes[ count ];
		IntPtr	ptrString;
		
		for( int i = 0; i < count; i++ )
			{
			attributes[i] = new OPCItemAttributes();

			ptrString = (IntPtr) Marshal.ReadInt32( (IntPtr) runatt );
			attributes[i].AccessPath = Marshal.PtrToStringUni( ptrString );
			Marshal.FreeCoTaskMem( ptrString );

			ptrString = (IntPtr) Marshal.ReadInt32( (IntPtr)(runatt + 4) );
			attributes[i].ItemID = Marshal.PtrToStringUni( ptrString );
			Marshal.FreeCoTaskMem( ptrString );

			attributes[i].Active			= (Marshal.ReadInt32( (IntPtr)(runatt + 8) )) != 0;
			attributes[i].HandleClient		= Marshal.ReadInt32( (IntPtr)(runatt + 12) );
			attributes[i].HandleServer		= Marshal.ReadInt32( (IntPtr)(runatt + 16) );
			attributes[i].AccessRights		= (OPCACCESSRIGHTS) Marshal.ReadInt32( (IntPtr)(runatt + 20) );
			attributes[i].RequestedDataType	= (VarEnum) Marshal.ReadInt16( (IntPtr)(runatt + 32) );
			attributes[i].CanonicalDataType	= (VarEnum) Marshal.ReadInt16( (IntPtr)(runatt + 34) );

			attributes[i].EUType			= (OPCEUTYPE) Marshal.ReadInt32( (IntPtr)(runatt + 36) );
			attributes[i].EUInfo			= Marshal.GetObjectForNativeVariant( (IntPtr)(runatt + 40) );
			DUMMY_VARIANT.VariantClear( (IntPtr) (runatt + 40) );

			int ptrblob	= Marshal.ReadInt32( (IntPtr)(runatt + 28) );
			if( (ptrblob != 0) )
				{
				int blobsize	= Marshal.ReadInt32( (IntPtr)(runatt + 24) );
				if( blobsize > 0 )
					{
					attributes[i].Blob = new byte[ blobsize ];
					Marshal.Copy( (IntPtr) ptrblob, attributes[i].Blob, 0, blobsize );
					}
				Marshal.FreeCoTaskMem( (IntPtr) ptrblob );
				}

			runatt += 56;
			}
		
		Marshal.FreeCoTaskMem( ptrAtt );
		}
		
	public void Skip( int celt )
		{ ifEnum.Skip( celt ); }

	public void Reset()
		{ ifEnum.Reset(); }

	private	IEnumOPCItemAttributes		ifEnum;

	}	// class OpcEnumItemAttributes




}	// namespace OPC.Data
