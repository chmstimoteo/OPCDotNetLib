/*=====================================================================
  File:      OPC_Data_Srv.cs

  Summary:   OPC DA server interfaces wrapper class

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
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;

using OPC.Common;
using OPC.Data.Interface;

namespace OPC.Data
{



// ------------- managed side only structs ----------------------
public class OPCProperty					// QueryAvailableProperties
	{
	public int		PropertyID;
	public string	Description;
	public VarEnum	DataType;
	
	public override string ToString()
		{
		return "ID:" + PropertyID + " '" + Description + "' T:" + DUMMY_VARIANT.VarEnumToString( DataType );
		}
	}

public class OPCPropertyData				// GetItemProperties
	{
	public int		PropertyID;
	public int		Error;
	public object	Data;

	public override string ToString()
		{
		if( Error == HRESULTS.S_OK )
			return "ID:" + PropertyID + " Data:" + Data.ToString();
		else
			return "ID:" + PropertyID + " Error:" + Error.ToString();
		}
	}

public class OPCPropertyItem				// LookupItemIDs
	{
	public int		PropertyID;
	public int		Error;
	public string	newItemID;

	public override string ToString()
		{
		if( Error == HRESULTS.S_OK )
			return "ID:" + PropertyID + " newID:" + newItemID;
		else
			return "ID:" + PropertyID + " Error:" + Error.ToString();
		}
	}










// ----------------- event argument+handler ------------------------

						// IOPCShutdown
public class ShutdownRequestEventArgs : EventArgs
	{
	public ShutdownRequestEventArgs( string shutdownReasonp )
		{
		shutdownReason = shutdownReasonp;
		}

	public string				shutdownReason;
	}
public delegate void ShutdownRequestEventHandler( object sender, ShutdownRequestEventArgs e );








// --------------------------- OpcServer ------------------------

	[ComVisible(true)]
public class OpcServer : IOPCShutdown
	{
	public OpcServer()
		{
		}
	~OpcServer()
		{ Disconnect(); }
		

	public void Connect( string	clsidOPCserver )
		{
		Disconnect();

		Type	typeofOPCserver = Type.GetTypeFromProgID( clsidOPCserver );
		if( typeofOPCserver == null )
			Marshal.ThrowExceptionForHR( HRESULTS.OPC_E_NOTFOUND );

		OPCserverObj = Activator.CreateInstance( typeofOPCserver );
		ifServer = (IOPCServer) OPCserverObj;
		if( ifServer == null )
			Marshal.ThrowExceptionForHR( HRESULTS.CONNECT_E_NOCONNECTION );

		// connect all interfaces
		ifCommon		= (IOPCCommon)						OPCserverObj;
		ifBrowse		= (IOPCBrowseServerAddressSpace)	ifServer;
		ifItmProps		= (IOPCItemProperties)				ifServer;
		cpointcontainer	= (UCOMIConnectionPointContainer)	OPCserverObj;
		AdviseIOPCShutdown();
		}

	public void Disconnect()
		{
		if( ! (shutdowncpoint == null) )
			{
			if( shutdowncookie != 0 )
				{
				shutdowncpoint.Unadvise( shutdowncookie );
				shutdowncookie = 0;
				}
			int	rc = Marshal.ReleaseComObject( shutdowncpoint );
			shutdowncpoint = null;
			}

		cpointcontainer = null;
		ifBrowse		= null;
		ifItmProps		= null;
		ifCommon		= null;
		ifServer		= null;
		if( ! (OPCserverObj == null) )
			{
			int rc = Marshal.ReleaseComObject( OPCserverObj );
			OPCserverObj = null;
			}
		}

	public void GetStatus( out SERVERSTATUS serverStatus )
		{
		ifServer.GetStatus( out serverStatus );
		}

	public string GetErrorString( int errorCode, int localeID )
		{
		string	errorres;
		ifServer.GetErrorString( errorCode, localeID, out errorres );
		return errorres;
		}


	public OpcGroup AddGroup(	string groupName, bool setActive, int requestedUpdateRate )
		{
		return AddGroup( groupName, setActive, requestedUpdateRate, null, null, 0 );
		}
	public OpcGroup AddGroup(	string groupName, bool setActive, int requestedUpdateRate,
								int[] biasTime, float[] percentDeadband, int localeID )
		{
		if( ifServer == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );
		
		OpcGroup	grp = new OpcGroup( ref ifServer, false, groupName, setActive, requestedUpdateRate );
		grp.internalAdd( biasTime, percentDeadband, localeID );
		return grp;
		}


	// --------------------------------- IOPCServerPublicGroups (indirect) -----------------
	public OpcGroup GetPublicGroup(	string groupName )
		{
		if( ifServer == null )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );
		
		OpcGroup	grp = new OpcGroup( ref ifServer, true, groupName, false, 1000 );
		grp.internalAdd( null, null, 0 );
		return grp;
		}





	// --------------------------------- IOPCCommon -------------------------
	public void SetLocaleID( int lcid )
		{
		ifCommon.SetLocaleID( lcid );
		}
	public void GetLocaleID( out int lcid )
		{
		ifCommon.GetLocaleID( out lcid );
		}
	public void QueryAvailableLocaleIDs( out int[] lcids )
		{
		lcids	= null;
		int		count;
		IntPtr	ptrIds;
		int	hresult = ifCommon.QueryAvailableLocaleIDs( out count, out ptrIds );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );
		if( ((int) ptrIds) == 0 )
			return;
		if( count < 1 )
			{ Marshal.FreeCoTaskMem( ptrIds ); return; }

		lcids = new int[ count ];
		Marshal.Copy( ptrIds, lcids, 0, count );
		Marshal.FreeCoTaskMem( ptrIds );
		}

	public void SetClientName( string name )
		{
		ifCommon.SetClientName( name );
		}








	// ------------------------ IOPCBrowseServerAddressSpace ---------------

	public OPCNAMESPACETYPE QueryOrganization()
		{
		OPCNAMESPACETYPE	ns;
		ifBrowse.QueryOrganization( out ns );
		return ns;
		}
		
	public void ChangeBrowsePosition( OPCBROWSEDIRECTION direction, string name )
		{
		ifBrowse.ChangeBrowsePosition( direction, name );
		}

	public void BrowseOPCItemIDs(	OPCBROWSETYPE filterType, string filterCriteria,
									VarEnum dataTypeFilter, OPCACCESSRIGHTS accessRightsFilter,
									out UCOMIEnumString stringEnumerator )
		{
		stringEnumerator = null;
		object	enumtemp;
		ifBrowse.BrowseOPCItemIDs( filterType, filterCriteria, (short) dataTypeFilter, accessRightsFilter, out enumtemp );
		stringEnumerator = (UCOMIEnumString) enumtemp;
		enumtemp = null;
		}

	public string GetItemID( string itemDataID )
		{
		string itemid;
		ifBrowse.GetItemID( itemDataID, out itemid );
		return itemid;
		}

	public void BrowseAccessPaths( string itemID, out UCOMIEnumString stringEnumerator )
		{
		stringEnumerator = null;
		object	enumtemp;
		ifBrowse.BrowseAccessPaths( itemID, out enumtemp );
		stringEnumerator = (UCOMIEnumString) enumtemp;
		enumtemp = null;
		}

	// extra helper
	public void Browse( OPCBROWSETYPE typ, out ArrayList lst )
		{
		lst = null;
		UCOMIEnumString	enumerator;
		BrowseOPCItemIDs( typ, "", VarEnum.VT_EMPTY, 0, out enumerator );
		if( enumerator == null )
			return;

		lst = new ArrayList( 500 );
		int			cft;
		string[]	strF = new string[100];
		int			hresult;
		do
			{
			cft = 0;
			hresult = enumerator.Next( 100, strF, out cft );
			if( cft > 0 )
				{
				for( int i = 0; i < cft; i++ )
					lst.Add( strF[i] );
				}
			}
		while( hresult == HRESULTS.S_OK );

		int	rc = Marshal.ReleaseComObject( enumerator );
		enumerator = null;
		lst.TrimToSize();
		}








	// ------------------------ IOPCItemProperties ---------------

	public void QueryAvailableProperties( string itemID, out OPCProperty[] opcProperties )
		{
		opcProperties = null;

		int	count = 0;
		IntPtr	ptrID;
		IntPtr	ptrDesc;
		IntPtr	ptrTyp;
		ifItmProps.QueryAvailableProperties( itemID, out count, out ptrID, out ptrDesc, out ptrTyp );
		if( (count == 0) || (count > 10000) )
			return;

		int	runID	= (int) ptrID;
		int	runDesc	= (int) ptrDesc;
		int	runTyp	= (int) ptrTyp;
		if( (runID == 0) || (runDesc == 0) || (runTyp == 0) )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		opcProperties = new OPCProperty[ count ];

		IntPtr ptrString;
		for( int i = 0; i < count; i++ )
			{
			opcProperties[i] = new OPCProperty();
			
			opcProperties[i].PropertyID = Marshal.ReadInt32( (IntPtr) runID );
			runID += 4;

			ptrString = (IntPtr) Marshal.ReadInt32( (IntPtr) runDesc );
			runDesc += 4;
			opcProperties[i].Description = Marshal.PtrToStringUni( ptrString );
			Marshal.FreeCoTaskMem( ptrString );

			opcProperties[i].DataType = (VarEnum) Marshal.ReadInt16( (IntPtr) runTyp );
			runTyp += 2;
			}

		Marshal.FreeCoTaskMem( ptrID );
		Marshal.FreeCoTaskMem( ptrDesc );
		Marshal.FreeCoTaskMem( ptrTyp );
		}



	public bool GetItemProperties( string itemID, int[] propertyIDs, out OPCPropertyData[] propertiesData )
		{
		propertiesData = null;
		int	count = propertyIDs.Length;
		if( count < 1 )
			return false;

		IntPtr	ptrDat;
		IntPtr	ptrErr;
		int	hresult = ifItmProps.GetItemProperties( itemID, count, propertyIDs, out ptrDat, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		int	runDat = (int) ptrDat;
		int	runErr = (int) ptrErr;
		if( (runDat == 0) || (runErr == 0) )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		propertiesData = new OPCPropertyData[ count ];

		for( int i = 0; i < count; i++ )
			{
			propertiesData[i] = new OPCPropertyData();
			propertiesData[i].PropertyID = propertyIDs[i];

			propertiesData[i].Error = Marshal.ReadInt32( (IntPtr) runErr );
			runErr += 4;

			if( propertiesData[i].Error == HRESULTS.S_OK )
				{
				propertiesData[i].Data = Marshal.GetObjectForNativeVariant( (IntPtr) runDat );
				DUMMY_VARIANT.VariantClear( (IntPtr) runDat );
				}
			else
				propertiesData[i].Data = null;
				
			runDat += DUMMY_VARIANT.ConstSize;
			}

		Marshal.FreeCoTaskMem( ptrDat );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}


	public bool LookupItemIDs( string itemID, int[] propertyIDs, out OPCPropertyItem[] propertyItems )
		{
		propertyItems = null;
		int	count = propertyIDs.Length;
		if( count < 1 )
			return false;

		IntPtr	ptrErr;
		IntPtr	ptrIds;
		int	hresult = ifItmProps.LookupItemIDs( itemID, count, propertyIDs, out ptrIds, out ptrErr );
		if( HRESULTS.Failed( hresult ) )
			Marshal.ThrowExceptionForHR( hresult );

		int	runIds = (int) ptrIds;
		int	runErr = (int) ptrErr;
		if( (runIds == 0) || (runErr == 0) )
			Marshal.ThrowExceptionForHR( HRESULTS.E_ABORT );

		propertyItems = new OPCPropertyItem[ count ];

		IntPtr ptrString;
		for( int i = 0; i < count; i++ )
			{
			propertyItems[i] = new OPCPropertyItem();
			propertyItems[i].PropertyID = propertyIDs[i];

			propertyItems[i].Error = Marshal.ReadInt32( (IntPtr) runErr );
			runErr += 4;

			if( propertyItems[i].Error == HRESULTS.S_OK )
				{
				ptrString = (IntPtr) Marshal.ReadInt32( (IntPtr) runIds );
				propertyItems[i].newItemID = Marshal.PtrToStringUni( ptrString );
				Marshal.FreeCoTaskMem( ptrString );
				}
			else
				propertyItems[i].newItemID = null;
				
			runIds += 4;
			}

		Marshal.FreeCoTaskMem( ptrIds );
		Marshal.FreeCoTaskMem( ptrErr );
		return hresult == HRESULTS.S_OK;
		}




	// ------------------------ IOPCShutdown --------------- COMMON CALLBACK
	void IOPCShutdown.ShutdownRequest( string shutdownReason )
		{
		ShutdownRequestEventArgs	e = new ShutdownRequestEventArgs( shutdownReason );
		if( ShutdownRequested != null )
			ShutdownRequested( this, e );
		}

	// -------------------------- event ---------------------
	public event ShutdownRequestEventHandler		ShutdownRequested;





	// -------------------------- private ---------------------

	private void AdviseIOPCShutdown()
		{
		Type	sinktype = typeof( IOPCShutdown );
		Guid	sinkguid = sinktype.GUID;
		
		cpointcontainer.FindConnectionPoint( ref sinkguid, out shutdowncpoint );
		if( shutdowncpoint == null )
			return;

		shutdowncpoint.Advise( this, out shutdowncookie );
		}

	private object							OPCserverObj	= null;
	private IOPCServer						ifServer		= null;
	private IOPCCommon						ifCommon		= null;
	private IOPCBrowseServerAddressSpace	ifBrowse		= null;
	private IOPCItemProperties				ifItmProps		= null;
	
	private	UCOMIConnectionPointContainer	cpointcontainer	= null;
	private	UCOMIConnectionPoint			shutdowncpoint	= null;
	private int								shutdowncookie	= 0;

	}	// class OpcServer











}	// namespace OPC.Data
