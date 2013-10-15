#region Namespaces
using System;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Mechanical;
#endregion

namespace MoveDuctJoin
{
  #region DuctSelectionFilter
  /// <summary>
  /// Allow selection of curve elements only.
  /// </summary>
  class DuctSelectionFilter : ISelectionFilter
  {
    public bool AllowElement( Element e )
    {
      return e is Duct;
    }

    public bool AllowReference( Reference r, XYZ p )
    {
      return true;
    }
  }
  #endregion // DuctSelectionFilter

  #region CmdDisconnect
  /// <summary>
  /// External command to move a duct connector 
  /// away from its original position along the
  /// duct centre line, disconnecting from the
  /// neighbour element.
  /// </summary>
  [Transaction( TransactionMode.Manual )]
  public class CmdDisconnect : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = uidoc.Document;
      Selection sel = uidoc.Selection;
      Duct duct = null;
      XYZ pFrom = null;
      XYZ pTo = null;

      try
      {
        Reference r = sel.PickObject(
          ObjectType.Element,
          new DuctSelectionFilter(),
          "Please pick a duct at the "
          + "connection to move." );

        duct = doc.GetElement( r.ElementId ) as Duct;
        pFrom = r.GlobalPoint;

        r = sel.PickObject(
          ObjectType.Element,
          new DuctSelectionFilter(),
          "Please pick a target point on the "
          + "duct to move the connection to." );

        pTo = r.GlobalPoint;
      }
      catch( Autodesk.Revit.Exceptions
        .OperationCanceledException )
      {
        return Result.Cancelled;
      }

      // Determine connector closest to picked point

      ConnectorSet connectors 
        = duct.ConnectorManager.Connectors;

      Connector con = null;
      double d, dmin = double.MaxValue;

      foreach( Connector c in connectors )
      {
        d = pFrom.DistanceTo( c.Origin );

        if( d < dmin )
        {
          dmin = d;
          con = c;
        }
      }

      // Determine target point to move it to

      Transform cs = con.CoordinateSystem;

      Debug.Assert( 
        con.Origin.IsAlmostEqualTo( cs.Origin ),
        "expected same origin" );

      Line line = Line.CreateUnbound( cs.Origin, cs.BasisZ );

      IntersectionResult ir = line.Project( pTo );

      pTo = ir.XYZPoint;

      Debug.Assert( line.Distance( pTo ) < 1e-9,
        "expected projected point on line" );

      // Modify document within a transaction

      using( Transaction tx = new Transaction( doc ) )
      {
        tx.Start( "Move Duct Connector" );
        con.Origin = pTo;
        tx.Commit();
      }
      return Result.Succeeded;
    }
  }
  #endregion // CmdDisconnect

  #region CmdReconnect
  /// <summary>
  /// External command to move a duct connector 
  /// away from its original position along the
  /// duct centre line, disconnecting from the
  /// neighbour element.
  /// </summary>
  [Transaction( TransactionMode.Manual )]
  public class CmdReconnect : IExternalCommand
  {
    /// <summary>
    /// Return the connector 
    /// connected to the one given.
    /// </summary>
    static Connector GetConnectedConnector( 
      Connector con )
    {
      Connector neighbour = null;

      int ownerId = con.Owner.Id.IntegerValue;

      ConnectorSet refs = con.AllRefs;

      foreach( Connector c in refs )
      {
        // Ignore non-End connectors and  
        // connectors on the same element

        if( c.ConnectorType == ConnectorType.End
          && !ownerId.Equals(
            c.Owner.Id.IntegerValue ) )
        {
          neighbour = c;
          break;
        }
      }
      return neighbour;
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = uidoc.Document;
      Selection sel = uidoc.Selection;
      Duct duct = null;
      XYZ pFrom = null;
      XYZ pTo = null;

      try
      {
        Reference r = sel.PickObject(
          ObjectType.Element,
          new DuctSelectionFilter(),
          "Please pick a duct at the "
          + "connection to move." );

        duct = doc.GetElement( r.ElementId ) as Duct;
        pFrom = r.GlobalPoint;

        r = sel.PickObject(
          ObjectType.Element,
          new DuctSelectionFilter(),
          "Please pick a target point on the "
          + "duct to move the connection to." );

        pTo = r.GlobalPoint;
      }
      catch( Autodesk.Revit.Exceptions
        .OperationCanceledException )
      {
        return Result.Cancelled;
      }

      // Determine connector closest to picked point

      ConnectorSet connectors
        = duct.ConnectorManager.Connectors;

      Connector con = null;
      double d, dmin = double.MaxValue;

      foreach( Connector c in connectors )
      {
        d = pFrom.DistanceTo( c.Origin );

        if( d < dmin )
        {
          dmin = d;
          con = c;
        }
      }

      // Determine target point to move it to

      Transform cs = con.CoordinateSystem;

      Debug.Assert(
        con.Origin.IsAlmostEqualTo( cs.Origin ),
        "expected same origin" );

      Line line = Line.CreateUnbound( cs.Origin, cs.BasisZ );

      IntersectionResult ir = line.Project( pTo );

      pTo = ir.XYZPoint;

      Debug.Assert( line.Distance( pTo ) < 1e-9,
        "expected projected point on line" );

      // Determine translation vector

      XYZ v = pTo - pFrom;

      // Determine neighbouring fitting

      Connector neighbour 
        = GetConnectedConnector( con );

      // Modify document within a transaction

      using( Transaction tx = new Transaction( doc ) )
      {
        tx.Start( "Move Fitting" );

        ElementTransformUtils.MoveElement( 
          doc, neighbour.Owner.Id, v ); 

        tx.Commit();
      }
      return Result.Succeeded;
    }
  }
  #endregion // CmdReconnect
}
