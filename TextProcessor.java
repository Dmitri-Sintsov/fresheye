/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package org.openoffice.uniyar.fresheye;

import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.table.XCell;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XEnumeration;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.lang.XServiceInfo;

/**
 *
 * @author Administrator
 */
public class TextProcessor implements Runnable {

    public XText xText;
    public Integer ParagraphsCount = 0;

    public TextProcessor( XText xText ) {
        this.xText = xText;
    }

    private void GetNestedCells( Object TextTable ) {
        XCell Cell;
        boolean CellHasElements;
        XEnumeration Enum;
        Object Elem = null;
        XServiceInfo ElemServiceInfo;
        XTextRange Anchor;
        String[] CellNames = ( (XTextTable)UnoRuntime.queryInterface( XTextTable.class, TextTable ) ).getCellNames();
        for ( int i = 0; i < CellNames.length; i++ ) {
            Cell = ( (XTextTable)UnoRuntime.queryInterface( XTextTable.class, TextTable ) ).getCellByName( CellNames[i] );
            CellHasElements = ( (XEnumerationAccess)UnoRuntime.queryInterface(XEnumerationAccess.class, Cell ) ).hasElements();
            if ( CellHasElements ) {
                Enum = ( (XEnumerationAccess)UnoRuntime.queryInterface( XEnumerationAccess.class, Cell ) ).createEnumeration();
                while ( Enum.hasMoreElements() ) {
                    try {
                        Elem = Enum.nextElement();
                    } catch ( Exception E ) {
                        System.err.println( "Could not get next element of XText: " + E );
                        E.printStackTrace( System.out );
                        System.exit( 0 );
                    }
                    ElemServiceInfo = (XServiceInfo)UnoRuntime.queryInterface( XServiceInfo.class, Elem );
                    if ( ElemServiceInfo.supportsService("com.sun.star.text.TextTable") ) {
                        GetNestedCells(Elem);
                    }
                    if ( ElemServiceInfo.supportsService("com.sun.star.text.TextContent") ) {
                        Anchor = ( (XTextContent)UnoRuntime.queryInterface( XTextContent.class, Elem ) ).getAnchor();
                        ProcessParagraph(Anchor);
                    }
                }
            } else {
                ProcessParagraph( (XText)UnoRuntime.queryInterface( XText.class, Cell ) );
            }
        }
    }

    public void ProcessParagraph( XTextRange Para ) {
        this.ParagraphsCount++;
    }

    // returns the count of paragraphs processed
    public void run() {
        Object Elem = null;
        XEnumeration Enum =
            ( (XEnumerationAccess)UnoRuntime.queryInterface( XEnumerationAccess.class, xText ) ).createEnumeration();
        while ( Enum.hasMoreElements() ) {
            try {
                Elem = Enum.nextElement();
            } catch ( Exception E ) {
                System.err.println( "Could not get next element of XText: " + E );
                E.printStackTrace( System.out );
                System.exit( 0 );
            }
            XServiceInfo ElemServiceInfo = (XServiceInfo)UnoRuntime.queryInterface( XServiceInfo.class, Elem );
            if ( ElemServiceInfo.supportsService("com.sun.star.text.TextTable") ) {
                GetNestedCells(Elem);
            }
            if ( ElemServiceInfo.supportsService("com.sun.star.text.TextContent") ) {
                XTextRange Anchor = ( (XTextContent)UnoRuntime.queryInterface( XTextContent.class, Elem ) ).getAnchor();
                ProcessParagraph( Anchor );
            }
        }
    }

    public int getResult() {
        return this.ParagraphsCount;
    }

}
