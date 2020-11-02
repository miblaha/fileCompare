package com.redhat.eap.qe.docs.Compare;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.*;
import com.sun.star.io.IOException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import ooo.connector.BootstrapSocketConnector;


/**
 * Use OpenOffice to compare documents
 */
public class Test {
    public static void main(String[] args) throws Exception {

       // String ooHome = System.getProperty("ooHome");
        String newFile = "file:///home/miblaha/Downloads/EAP-CD20-bulk97/19.odt";
        String oldFile = "file:///home/miblaha/Downloads/EAP-CD20-bulk97/20.odt";
        String comparedFile = "file:///home/miblaha/Downloads/EAP-CD20-bulk97/21.odt";
        // String ooHome = "/home/miblaha/libreoffice4/program";
        String ooHome = "/opt/libreoffice6.4/program";

        XMultiComponentFactory xMCF = null;

        Object oDesktop = null;
        try {
            XComponentContext xContext = BootstrapSocketConnector.bootstrap(ooHome);
            //XComponentContext xContext = Bootstrap.bootstrap();
            xMCF = xContext.getServiceManager();
            oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
            UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);
        } catch (BootstrapException | Exception e) {
            e.printStackTrace();
        }

        // Query for the XPropertySet interface
        XPropertySet properestMultiComponentFactory = UnoRuntime.queryInterface(XPropertySet.class, xMCF);

        // Get the default context from the office server.
        Object objectDefaultContext = null;
        try {
            objectDefaultContext = properestMultiComponentFactory.getPropertyValue("DefaultContext");
        } catch (UnknownPropertyException | WrappedTargetException e) {
            e.printStackTrace();
        }

        // Query for the interface XComponentContext.
        XComponentContext xcomponentcontext = UnoRuntime.queryInterface(XComponentContext.class, objectDefaultContext);

        XComponentLoader xcomponentloader = UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);

        // Preparing properties for loading the document
        PropertyValue[] propertyvalue = new PropertyValue[1];
        // Setting the flag for hidding the open document
        propertyvalue[0] = new PropertyValue();
        propertyvalue[0].Name = "Hidden";
        propertyvalue[0].Value = Boolean.TRUE;
        //TODO: Hardcoding opening word documents -- this will need to change.
        //propertyvalue[ 1 ] = new PropertyValue();
        //propertyvalue[ 1 ].Name = "FilterName";
        //propertyvalue[ 1 ].Value = "HTML (StarWriter)";

        // Loading the wanted document
        Object objectDocumentToStore = null;
        try {
            objectDocumentToStore = xcomponentloader.loadComponentFromURL(newFile, "_blank", 0, propertyvalue);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Getting an object that will offer a simple way to store a document to a URL.
        XStorable xstorable = UnoRuntime.queryInterface(XStorable.class, objectDocumentToStore);

        // Preparing properties for comparing the document
        propertyvalue = new PropertyValue[1];
        // Setting the flag for overwriting
        propertyvalue[0] = new PropertyValue();
        propertyvalue[0].Name = "URL";
        propertyvalue[0].Value = oldFile;
        // Setting the filter name
        //propertyvalue[ 1 ] = new PropertyValue();
        //propertyvalue[ 1 ].Name = "FilterName";
        //propertyvalue[ 1 ].Value = context.get("convertFilterName");


        XFrame frame;
        frame = UnoRuntime.queryInterface(XFrame.class, oDesktop);
        Object dispatchHelperObj = null;
        try {
            dispatchHelperObj = xMCF.createInstanceWithContext("com.sun.star.frame.DispatchHelper", xcomponentcontext);
        } catch (Exception e) {
            e.printStackTrace();
        }
        XDispatchHelper dispatchHelper = UnoRuntime.queryInterface(XDispatchHelper.class, dispatchHelperObj);
        XDispatchProvider dispatchProvider = UnoRuntime.queryInterface(XDispatchProvider.class, frame);
        dispatchHelper.executeDispatch(dispatchProvider, ".uno:CompareDocuments", "", 0, propertyvalue);

        // Preparing properties for storing the document
        propertyvalue = new PropertyValue[1];
        // Setting the flag for overwriting
        propertyvalue[0] = new PropertyValue();
        propertyvalue[0].Name = "Overwrite";
        propertyvalue[0].Value = Boolean.TRUE;
        // Setting the filter name
        //propertyvalue[ 1 ] = new PropertyValue();
        //propertyvalue[ 1 ].Name = "FilterName";
        //propertyvalue[ 1 ].Value = context.get("convertFilterName");

        // Storing and converting the document
        try {
            xstorable.storeToURL(comparedFile, propertyvalue);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Getting the method dispose() for closing the document

        XComponent xcomponent = UnoRuntime.queryInterface(XComponent.class, xstorable);

        // Closing the converted document
        xcomponent.dispose();
        System.exit(18);

    }
}
