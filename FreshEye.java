package org.openoffice.uniyar.fresheye;

import java.util.Hashtable;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.awt.Rectangle;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XTextDocument;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.frame.XDesktop;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.awt.InvalidateStyle;
import com.sun.star.frame.XFrame;
import com.sun.star.text.XText;

public final class FreshEye extends WeakBase
   implements com.sun.star.lang.XServiceInfo,
              com.sun.star.frame.XDispatchProvider,
              com.sun.star.lang.XInitialization,
              com.sun.star.frame.XDispatch {
    private static final String m_implementationName = FreshEye.class.getName();
    private static final String[] m_serviceNames = {
        "com.sun.star.frame.ProtocolHandler" };

    /*** begin of FreshEye main objects ***/
    private XTextDocument xTextDocument;
    private XComponentContext xComponentContext;
    private XFrame xFrame;
    private XWindow oContainerWindow;
    private XWindowPeer oContWindowPeer;

    // sensitivity threshold (default, min, max)
    private Long Threshold_val = new Long( 600 );
    private Long Threshold_min = new Long( 400 );
    private Long Threshold_max = new Long( 1100 );

    // progress calculation
    private Integer Progress_max = 0;

    // window definitions
    private Rectangle modBounds = new Rectangle( 0, 0, 210, 60 );
    private Hashtable<String,Object> modProps = new Hashtable<String,Object>();
    private Hashtable<String,Object> modControls = new Hashtable<String,Object>();
    private Rectangle nonmodBounds = new Rectangle( 0, 0, 210, 10 );
    private Hashtable<String,Object> nonmodProps = new Hashtable<String,Object>();
    private Hashtable<String,Object> nonmodControls = new Hashtable<String,Object>();

/*** end of FreshEye main objects ***/

    // initializes window properties
    private void initializeWindowProps() {
        modProps.put( "Name", "ModDialog" );
        modProps.put( "Closeable", true );
        modProps.put( "Moveable", true );
        modProps.put( "Sizeable", true );
        modProps.put( "Title", "Свежий взгляд" );
        modProps.put( "BackgroundColor", new Integer( 0xEEEEEE ) );
        modProps.put( "PositionX", new Integer( 0 ) );
        modProps.put( "PositionY", new Integer( 0 ) );
        modProps.put( "Width", new Integer( 210 ) );
        modProps.put( "Height", new Integer( 60 ) );

        Hashtable<String,Object> labelProps = new Hashtable<String,Object>();
        labelProps.put( "_TYPE_", "FixedText" ); // "magic" value (see class NonModalWnd)
        labelProps.put( "Label", "Введите значение порога схожести (min = " + Threshold_min + ", max = " + Threshold_max + ")" );
        labelProps.put( "Border", new Short( (short) 0 ) );
        labelProps.put( "PositionX", new Integer( 10 ) );
        labelProps.put( "PositionY", new Integer( 10 ) );
        labelProps.put( "Width", new Integer( 160 ) );
        labelProps.put( "Height", new Integer( 20 ) );
        modControls.put( "Label1",  labelProps );

        Hashtable<String,Object> numfieldProps = new Hashtable<String,Object>();
        numfieldProps.put( "_TYPE_", "NumericField" );
        numfieldProps.put( "Value", new Double( (double) Threshold_val ) );
        numfieldProps.put( "ValueMin", new Double( Threshold_min ) );
        numfieldProps.put( "ValueMax", new Double( Threshold_max ) );
        numfieldProps.put( "DecimalAccuracy", new Short( (short)0 ) );
        numfieldProps.put( "Border", new Short( (short)1 ) );
        numfieldProps.put( "PositionX", new Integer( 170 ) );
        numfieldProps.put( "PositionY", new Integer( 7 ) );
        numfieldProps.put( "Width", new Integer( 20 ) );
        numfieldProps.put( "Height", new Integer( 15 ) );
        numfieldProps.put( "Tabstop", true );
        modControls.put( "Threshold",  numfieldProps );

        Hashtable<String,Object> buttonOkProps = new Hashtable<String,Object>();
        buttonOkProps.put( "_TYPE_", "Button" );
        buttonOkProps.put( "Label", "Проверка" );
        buttonOkProps.put( "DefaultButton", true );
        buttonOkProps.put( "PushButtonType", new Short( (short)1 ) );
        buttonOkProps.put( "PositionX", new Integer( 125 ) );
        buttonOkProps.put( "PositionY", new Integer( 35 ) );
        buttonOkProps.put( "Width", new Integer( 40 ) );
        buttonOkProps.put( "Height", new Integer( 15 ) );
        buttonOkProps.put( "Tabstop", true );
        modControls.put( "ButtonOk",  buttonOkProps );

        Hashtable<String,Object> buttonCancelProps = new Hashtable<String,Object>();
        buttonCancelProps.put( "_TYPE_", "Button" );
        buttonCancelProps.put( "Label", "Отмена" );
        buttonCancelProps.put( "PushButtonType", new Short( (short)2 ) );
        buttonCancelProps.put( "PositionX", new Integer( 170 ) );
        buttonCancelProps.put( "PositionY", new Integer( 35 ) );
        buttonCancelProps.put( "Width", new Integer( 30 ) );
        buttonCancelProps.put( "Height", new Integer( 15 ) );
        buttonCancelProps.put( "Tabstop", true );
        modControls.put( "ButtonCancel",  buttonCancelProps );

        nonmodProps.put( "Name", "NonModDialog" );
        nonmodProps.put( "Closeable", false );
        nonmodProps.put( "Moveable", true );
        nonmodProps.put( "Sizeable", true );
        nonmodProps.put( "Title", "Свежий взгляд - пожалуйста подождите..." );
        nonmodProps.put( "BackgroundColor", new Integer( 0xEEEEEE ) );
        nonmodProps.put( "PositionX", new Integer( 0 ) );
        nonmodProps.put( "PositionY", new Integer( 0 ) );
        nonmodProps.put( "Width", new Integer( nonmodBounds.Width ) );
        nonmodProps.put( "Height", new Integer( nonmodBounds.Height ) );

        Hashtable<String,Object> progressProps = new Hashtable<String,Object>();
        progressProps.put( "_TYPE_", "ProgressBar" );
        progressProps.put( "Border", new Short( (short)0 ) );
        progressProps.put( "FillColor", new Integer( 0x8000 ) );
        progressProps.put( "PositionX", new Integer( 0 ) );
        progressProps.put( "PositionY", new Integer( 0 ) );
        progressProps.put( "Width", new Integer( nonmodBounds.Width - 1 ) );
        progressProps.put( "Height", new Integer( nonmodBounds.Height - 1 ) );
        progressProps.put( "ProgressValueMin", new Integer( 0 ) );
        progressProps.put( "ProgressValueMax", new Integer( 0 ) );
        progressProps.put( "ProgressValue", 0 );
        nonmodControls.put( "ProgressBar1" , progressProps );
    }

    public FreshEye( XComponentContext xCompContext ) {
        try {
            this.xComponentContext = xCompContext;
            XMultiComponentFactory xMCF = this.xComponentContext.getServiceManager();
            Object desktop = xMCF.createInstanceWithContext( "com.sun.star.frame.Desktop", this.xComponentContext);
            XDesktop xDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class, desktop);
            XComponent xComponent = xDesktop.getCurrentComponent();
            this.xTextDocument = (XTextDocument) UnoRuntime.queryInterface( XTextDocument.class, xComponent );
        } catch (java.lang.Exception e) {
                e.printStackTrace(System.err);
        }
     };

    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ) {
        XSingleComponentFactory xFactory = null;

        if ( sImplementationName.equals( m_implementationName ) )
            xFactory = Factory.createComponentFactory(FreshEye.class, m_serviceNames);
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo( XRegistryKey xRegistryKey ) {
        return Factory.writeRegistryServiceInfo(m_implementationName,
                                                m_serviceNames,
                                                xRegistryKey);
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
         return m_implementationName;
    }

    public boolean supportsService( String sService ) {
        int len = m_serviceNames.length;

        for( int i=0; i < len; i++) {
            if (sService.equals(m_serviceNames[i]))
                return true;
        }
        return false;
    }

    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch queryDispatch( com.sun.star.util.URL aURL,
                                                       String sTargetFrameName,
                                                       int iSearchFlags ) {
        if ( aURL.Protocol.compareTo("org.openoffice.uniyar.fresheye.fresheye:") == 0 ) {
            if ( aURL.Path.compareTo("Cyrillic Eye") == 0 )
                return this;
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch[] queryDispatches(
         com.sun.star.frame.DispatchDescriptor[] seqDescriptors ) {
        int nCount = seqDescriptors.length;
        com.sun.star.frame.XDispatch[] seqDispatcher =
            new com.sun.star.frame.XDispatch[seqDescriptors.length];

        for( int i=0; i < nCount; ++i ) {
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL,
                                             seqDescriptors[i].FrameName,
                                             seqDescriptors[i].SearchFlags );
        }
        return seqDispatcher;
    }

    // com.sun.star.lang.XInitialization:
    public void initialize( Object[] object ) throws com.sun.star.uno.Exception  {
        if ( object.length > 0 ) {
            this.xFrame = (com.sun.star.frame.XFrame)UnoRuntime.queryInterface(
                com.sun.star.frame.XFrame.class, object[0]);
        }
    }

    // com.sun.star.frame.XDispatch:
    public void dispatch( com.sun.star.util.URL aURL,
                           com.sun.star.beans.PropertyValue[] aArguments ) {
         if ( aURL.Protocol.compareTo("org.openoffice.uniyar.fresheye.fresheye:") == 0 ) {
            if ( aURL.Path.compareTo("Cyrillic Eye") == 0 ) {
                this.initializeWindowProps();
                this.oContainerWindow = this.xFrame.getContainerWindow();
                this.oContWindowPeer = (XWindowPeer) UnoRuntime.queryInterface( XWindowPeer.class, this.oContainerWindow );
                NonModalWnd InputWnd = new NonModalWnd( this.xComponentContext, this.oContWindowPeer, this.modBounds );
                InputWnd.invalidate( InvalidateStyle.UPDATE );
                InputWnd.createDialog( this.modProps );
                InputWnd.addControls( this.modControls );
                InputWnd.Show();
                InputWnd.setFocus( "Threshold" );
                Short InputResult = InputWnd.executeDialog();
                try {
                    Double res = (Double) InputWnd.getControlPropsByName( "Threshold" ).getPropertyValue( "Value" );
                    Threshold_val = res.longValue();
                } catch ( Exception E ) {
                    E.printStackTrace( System.err );
                }
                InputWnd.Hide();
                InputWnd.Destroy();
                if ( InputResult != 0 ) {
                    NonModalWnd ProcessWnd = new NonModalWnd( this.xComponentContext, oContWindowPeer, this.nonmodBounds );
                    ProcessWnd.invalidate( InvalidateStyle.UPDATE );
                    ProcessWnd.createDialog( this.nonmodProps );
                    ProcessWnd.addControls( this.nonmodControls );
                    ProcessWnd.Show();
                    // get text from the document
                    XText xText = this.xTextDocument.getText();
                    // count paragraphs
                    TextProcessor TP = new TextProcessor( xText );
                    try {
                        Thread t = new Thread( TP );
                        t.start();
                        t.join();
                        this.Progress_max = TP.getResult();
                        ProcessWnd.setControlPropByName( "ProgressBar1", "ProgressValueMax", this.Progress_max );
                        // hightlight fresheye
                        TP = new FreshEyeProcessor( xText, Threshold_val.intValue() );
                        t = new Thread( TP );
                        t.start();
                        Short InvStyle = InvalidateStyle.UPDATE |
                            InvalidateStyle.CHILDREN |
                            InvalidateStyle.TRANSPARENT |
                            InvalidateStyle.NOCLIPCHILDREN;
                        while ( t.isAlive() ) {
                            ProcessWnd.setControlPropByName( "ProgressBar1", "ProgressValue", TP.ParagraphsCount );
                            ProcessWnd.invalidate( InvStyle );
                            this.oContWindowPeer.invalidate( InvStyle );
                            t.join( 1000 );
                        }
                    } catch ( InterruptedException E ) {
                    }
                    ProcessWnd.Hide();
                    ProcessWnd.Destroy();
                }
                return;
            }
        }
    }

    public void addStatusListener( com.sun.star.frame.XStatusListener xControl,
                                    com.sun.star.util.URL aURL ) {
        // add your own code here
    }

    public void removeStatusListener( com.sun.star.frame.XStatusListener xControl,
                                       com.sun.star.util.URL aURL ) {
        // add your own code here
    }

}
