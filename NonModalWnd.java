/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.openoffice.uniyar.fresheye;

import java.util.Hashtable;
import java.util.Enumeration;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.XComponentContext;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.uno.XInterface;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.WindowAttribute;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.XWindow;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFramesSupplier;
import com.sun.star.frame.XFrames;
import com.sun.star.awt.PosSize;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XControl;
import com.sun.star.awt.Rectangle;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XController;
import com.sun.star.view.XControlAccess;

/**
 *
 * @author Administrator
 */
public class NonModalWnd {

    private XComponentContext Context;
    private XMultiComponentFactory CompServMgr;
    private XWindowPeer ContWinPeer, SubPeer;
    private XToolkit xToolkit;
    private XWindow SubWindow;
    private XFrame SubFrame;
    private XFramesSupplier xTreeRoot;
    private XDialog xDialog;
    private XMultiServiceFactory DialogServMgr;
    private XInterface DialogModel;
    // list of dialog and it's controls names and XInterfaces
    private Hashtable<String,Object> controls = new Hashtable<String,Object>();

    public NonModalWnd( XComponentContext Context, XWindowPeer ContWinPeer, Rectangle Bounds ) {
        this.Context = Context;
        this.CompServMgr = Context.getServiceManager();
        this.ContWinPeer = ContWinPeer;
        XInterface oToolkit = (XInterface) this.contextInstance( "com.sun.star.awt.Toolkit" );
        this.xToolkit = (XToolkit) UnoRuntime.queryInterface( XToolkit.class, oToolkit );

        WindowDescriptor aWinDesc = new WindowDescriptor();
        aWinDesc.Type = WindowClass.TOP;
        aWinDesc.WindowServiceName = "";
        aWinDesc.ParentIndex = -1;
        aWinDesc.Parent = ContWinPeer;
        //aWinDesc.Parent = this.xToolkit.getDesktopWindow();
        aWinDesc.Bounds = Bounds;
        aWinDesc.WindowAttributes = WindowAttribute.FULLSIZE;
        try {
            this.SubPeer = this.xToolkit.createWindow( aWinDesc );
        } catch( Exception ex ) {
            System.err.println( "Could not create the subpeer window : " + ex );
            ex.printStackTrace( System.out );
            System.exit( 0 );
        }
        this.SubWindow = (XWindow) UnoRuntime.queryInterface( XWindow.class, this.SubPeer );

        XInterface oFrameTask = (XInterface) this.contextInstance( "com.sun.star.frame.Frame" );
        this.SubFrame = (XFrame) UnoRuntime.queryInterface( XFrame.class, oFrameTask );
        // create subwindow as a parent of document window
        this.SubFrame.initialize( this.SubWindow );

        XInterface xDesktop = (XInterface) this.contextInstance( "com.sun.star.frame.Desktop" );
        this.xTreeRoot = (XFramesSupplier) UnoRuntime.queryInterface( XFramesSupplier.class, xDesktop );
        XFrames xChildContainer = this.xTreeRoot.getFrames();
        xChildContainer.append( this.SubFrame );
        //this.SubWindow.setPosSize( Attrib.posx, Attrib.posy, Attrib.width, Attrib.height, PosSize.POSSIZE );
        this.SubWindow.setPosSize( 0, 0, 1, 1, PosSize.POSSIZE );
    }

    private java.lang.Object contextInstance( String InstName ) {
        Object Instance = null;
        try {
            Instance = this.CompServMgr.createInstanceWithContext( InstName, this.Context );
        } catch( Exception E ) {
            System.err.println( "Couldn't create the instance of " + InstName + ": " + E );
            E.printStackTrace();
            System.exit( 0 );
        }
        return Instance;
    }

    public void createDialog( Hashtable<String,Object> Props ) {
        XInterface Dialog = (XInterface) this.contextInstance( "com.sun.star.awt.UnoControlDialog" );
        this.xDialog = (XDialog) UnoRuntime.queryInterface( XDialog.class, Dialog );
        this.DialogModel = (XInterface) this.contextInstance( "com.sun.star.awt.UnoControlDialogModel" );
        this.DialogServMgr = (XMultiServiceFactory) UnoRuntime.queryInterface( XMultiServiceFactory.class, this.DialogModel );
        XPropertySet oDialogModelProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, this.DialogModel );
        String DialogName = this.setControlProps( oDialogModelProps, Props );
        if ( DialogName.equals( "" ) ) {
            DialogName = "Dialog";
        }
        this.controls.put( DialogName, DialogModel );
        XControlModel oDialogControlModel = (XControlModel) UnoRuntime.queryInterface( XControlModel.class, this.DialogModel );
        XControl oDialogControl = (XControl) UnoRuntime.queryInterface( XControl.class, Dialog );
        oDialogControl.setModel( oDialogControlModel );
        // insert dialog into subwindow
        oDialogControl.createPeer( this.xToolkit, this.SubPeer );
    }

    // returns string value of "Name" property, if found, otherwise empty string
    private String setControlProps( XPropertySet ControlProps, Hashtable<String,Object> Props ) {
        String NameVal = "";
        String propKey = "unknown";
        Object propVal;
        XPropertySetInfo PropInfo = ControlProps.getPropertySetInfo();
        com.sun.star.beans.Property[] Propset = PropInfo.getProperties();
        Enumeration e = Props.keys();
        while ( e.hasMoreElements() ) {
            try {
                propKey = (String) e.nextElement();
                propVal = Props.get( propKey );
                if ( propKey.equals( "Name" ) ) {
                    if ( propVal.equals( "" ) ) {
                        throw new Exception();
                    }
                    NameVal = (String) propVal;
                }
                if ( !propKey.equals( "_TYPE_" ) ) {
                    ControlProps.setPropertyValue( propKey, propVal );
                }
            } catch( Exception E ) {
                System.err.println( "Couldn't set the value of control property " + propKey + ": " + E  );
                E.printStackTrace();
                System.exit( 0 );
            }
        }
        return NameVal;
    }

    public void addControl( String ControlName, Hashtable<String,Object> Props ) {
        String ControlType = "unknown";
        try {
            ControlType = (String) Props.get( "_TYPE_" );
            XInterface oControlModel = (XInterface) this.DialogServMgr.createInstance( "com.sun.star.awt.UnoControl" + ControlType + "Model" );
            XPropertySet oControlProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, oControlModel );
            // set "Name" property
            oControlProps.setPropertyValue( "Name", ControlName );
            // set the rest of the properties
            String nameResult = this.setControlProps( oControlProps, Props );
            if ( !nameResult.equals( "" ) && !nameResult.equals( ControlName ) ) {
                System.err.println( "Control property \"name\":" + ControlName +  " does not match : " + nameResult );
                throw new Exception();
            }
            this.controls.put( ControlName, oControlModel );
            // insert the controls models into the dialog model
            XNameContainer xDialogNameContainer = (XNameContainer) UnoRuntime.queryInterface( XNameContainer.class, this.DialogModel );
            xDialogNameContainer.insertByName( ControlName, oControlModel );
        } catch( Exception E ) {
            System.err.println( "Couldn't create the instance of " + ControlType + ": " + E );
            E.printStackTrace();
            System.exit( 0 );
        }
    }

    public void addControls( Hashtable<String,Object> PropsList ) {
        Enumeration e = PropsList.keys();
        while ( e.hasMoreElements() ) {
            String propsKey = (String) e.nextElement();
            Hashtable<String,Object> Prop = (Hashtable<String,Object>) PropsList.get( propsKey );
            this.addControl( propsKey, Prop);
        }
    }

    public void setFocus( String ControlName ) {
        try {
            XInterface Model = (XInterface) this.controls.get( ControlName );
            XControlModel xObjControlModel = (XControlModel) UnoRuntime.queryInterface( XControlModel.class, Model );
            XController xController = this.SubFrame.getController();
            if ( xController != null ) {
                XControlAccess xCtrlAcc = (XControlAccess) UnoRuntime.queryInterface( XControlAccess.class, xController );
                XControl xObjControl = xCtrlAcc.getControl( xObjControlModel );
                XControl xControl = (XControl) UnoRuntime.queryInterface( XControl.class, xObjControl );
                // the focus can be set to an XWindow only
                XWindow xControlWindow = (XWindow) UnoRuntime.queryInterface( XWindow.class, xControl );
                // grab the focus
                xControlWindow.setFocus();
            }
        } catch( Exception E ) {
            System.err.println( "Couldn't set focus to " + ControlName + ": " + E );
            E.printStackTrace();
            System.exit( 0 );
        }
    }

    // returns XPropertySet of the previousely added control with name ControlName
    public XPropertySet getControlPropsByName( String ControlName ) {
        XInterface Model = (XInterface) this.controls.get( ControlName );
        return (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, Model );
    }

    public void setControlPropByName( String ControlName, String propKey, Object propVal ) {
        XPropertySet ControlProps = this.getControlPropsByName( ControlName );
        try {
            ControlProps.setPropertyValue( propKey, propVal );
        } catch ( Exception E ) {
            System.err.println( "Could not set the value of property " + propKey + ": " + E );
            E.printStackTrace();
            System.exit( 0 );
        }
    }

    public short executeDialog() {
        return this.xDialog.execute();
    }

    public void invalidate( int mode ) {
       this.SubPeer.invalidate( (short) mode );
    }

    public void Show() {
        this.SubWindow.setVisible( true );
    }

    public void Hide() {
        this.SubWindow.setVisible( false );
    }

    public void Destroy() {
        this.xTreeRoot.getFrames().remove( this.SubFrame );
        this.SubWindow.dispose();
    }

}
