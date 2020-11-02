package com.redhat.eap.qe.docs.Compare;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 * Use OpenOffice to compare documents
 */
public class Compare {
    public static void main(String[] args) throws Exception {


        String newFile = "file:///home/miblaha/Downloads/EAP-CD20-bulk97/19.odt";
        String oldFile = "file:///home/miblaha/Downloads/EAP-CD20-bulk97/20.odt";
        String comparedFile = "file:///home/miblaha/Downloads/EAP-CD20-bulk97/22.odt";
/*
        ooHome = "/home/miblaha/openoffice4/program/";

        String OPENOFFICE_HOST = "";

        String newFile = System.getProperty("newFile");
        String oldFile = System.getProperty("oldFile");
        String comparedFile = System.getProperty("comparedFile");

        XMultiComponentFactory xMCF = null;

        // Create OOo server with additional -nofirststartwizard and -headless options
        List oooOptions = OOoServer.getDefaultOOoOptions();
        oooOptions.add("-nofirststartwizard");
        oooOptions.add("-headless");
        OOoServer oooServer = new OOoServer(ooHome, oooOptions); */

        Object oDesktop = null;
        XComponentLoader xcomponentloader = null;
        XComponentContext xRemoteContext = null;
        // BootstrapSocketConnector bootstrapSocketConnector = new BootstrapSocketConnector(oooServer);
        // XComponentContext xContext = bootstrapSocketConnector.connect();

        XComponentContext xLocalContext = null;
        try {
            // create default local component context
            xLocalContext = Bootstrap.createInitialComponentContext(null);
            // initial serviceManager
            XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
            // create a urlresolver
            Object urlResolver = xLocalServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", xLocalContext);
            // query for the XUnoUrlResolver interface
            XUnoUrlResolver xUrlResolver = UnoRuntime.queryInterface(XUnoUrlResolver.class, urlResolver);

            Object initialObject = xUrlResolver.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");

            XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, initialObject);

            Object context = xPropertySet.getPropertyValue("DefaultContext");

            xRemoteContext = UnoRuntime.queryInterface(XComponentContext.class, context);

            XMultiComponentFactory mxRemoteServiceManager = xRemoteContext.getServiceManager();

            oDesktop = mxRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", null);

            UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);


            // Query for the XPropertySet interface
            XPropertySet properestMultiComponentFactory = UnoRuntime.queryInterface(XPropertySet.class, mxRemoteServiceManager);

            // Get the default context from the office server.
            Object objectDefaultContext = null;

            objectDefaultContext = properestMultiComponentFactory.getPropertyValue("DefaultContext");


            // Query for the interface XComponentContext.
            XComponentContext xcomponentcontext = UnoRuntime.queryInterface(XComponentContext.class, objectDefaultContext);

            xcomponentloader = UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);

            //PropertyValue[] pPropValues = new PropertyValue[0];

            // XComponent xComponent = xcomponentloader.loadComponentFromURL("private:factory/swriter", "_blank", 0, pPropValues);


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
                dispatchHelperObj = mxRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", xcomponentcontext);
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
            System.exit(0);
        } catch (java.lang.Exception e) {
            e.printStackTrace();
        }
    }
}
