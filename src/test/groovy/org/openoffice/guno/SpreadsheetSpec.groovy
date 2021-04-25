/*
 * *************************************************************
 * 
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 * 
 *************************************************************/

package org.openoffice.guno

import com.sun.star.container.XIndexAccess
import com.sun.star.container.XNamed
import com.sun.star.frame.XComponentLoader
import com.sun.star.frame.XController
import com.sun.star.frame.XModel
import com.sun.star.lang.XComponent
import com.sun.star.lang.XMultiComponentFactory
import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheetView
import com.sun.star.sheet.XSpreadsheets
import com.sun.star.uno.RuntimeException
import com.sun.star.uno.UnoRuntime
import com.sun.star.uno.XComponentContext
import spock.lang.Narrative
import spock.lang.Shared
import spock.lang.Specification
import spock.lang.Title

import ooo.connector.BootstrapSocketConnector


/**
 *
 * @author Carl Marcum - CodeBuilders.net
 */
@Narrative(""" Spreadsheet specific tests for the SpreadsheetExtension.
Each test will start using the same running office and the office will shutdown afterward. """)
@Title("Unit tests for the SpreadsheetExtension")
class SpreadsheetSpec extends Specification {

    @Shared
    XComponentContext mxRemoteContext
    @Shared
    XMultiComponentFactory mxRemoteServiceManager
    @Shared
    XSpreadsheetDocument xSpreadsheetDocument
    @Shared
    XComponent xComponent


    // fixture methods (setup, cleanup, setupSpec, cleanupSpec)

    // put our expensive operations here like connections
    def setupSpec() {

        // connect to the office and get a component context
        if (mxRemoteContext == null) {
            try {
                // bootstrap by file path
                String oooExeFolder = "/opt/openoffice4/program/"
                mxRemoteContext = BootstrapSocketConnector.bootstrap(oooExeFolder);
                System.out.println("Connected to a running office ...")

                // mxRemoteServiceManager = mxRemoteContext.getServiceManager()

            } catch (Exception e) {
                System.err.println("ERROR: can't get a component context from a running office ...")
                e.printStackTrace()
                System.exit(1)
            }
        }

        // replaces initDocument()
        XComponentLoader aLoader = mxRemoteContext.componentLoader

        xComponent = aLoader.loadComponentFromURL(
                "private:factory/scalc", "_default", 0, new com.sun.star.beans.PropertyValue[0])

        xSpreadsheetDocument = xComponent.getSpreadsheetDocument(mxRemoteContext)

    }

    def cleanupSpec() {

        // close it all down
        // Check supported functionality of the document (model or controller).
        com.sun.star.frame.XModel xModel = UnoRuntime.queryInterface(
                com.sun.star.frame.XModel.class, xSpreadsheetDocument)

        if (xModel != null) {
            // It is a full featured office document.
            // Try to use close mechanism instead of a hard dispose().
            // But maybe such service is not available on this model.
            com.sun.star.util.XCloseable xCloseable = UnoRuntime.queryInterface(
                    com.sun.star.util.XCloseable.class, xModel)

            if (xCloseable != null) {
                try {
                    // use close(boolean DeliverOwnership)
                    // The boolean parameter DeliverOwnership tells objects vetoing the close process that they may
                    // assume ownership if they object the closure by throwing a CloseVetoException
                    // Here we give up ownership. To be on the safe side, catch possible veto exception anyway.
                    xCloseable.close(true);
                } catch (com.sun.star.util.CloseVetoException exCloseVeto) {

                }
            }
            // If close is not supported by this model - try to dispose it.
            // But if the model disagree with a reset request for the modify state
            // we shouldn't do so. Otherwhise some strange things can happen.
            else {
                try {
                    com.sun.star.lang.XComponent xDisposeable = UnoRuntime.queryInterface(
                            com.sun.star.lang.XComponent.class, xModel)
                    xDisposeable.dispose()
                } catch (com.sun.star.beans.PropertyVetoException exModifyVeto) {

                }
            }

        }
    }


    // feature methods
    def "get sheet by name"() {

        given: "add a new spreadsheet MySheet"
        XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets()
        xSpreadsheets.insertNewByName("MySheet", (short) 0)


        when: "we request the sheet by name"
        XSpreadsheet xSpreadsheet = xSpreadsheetDocument.getSheetByName("MySheet")

        then: "we have a spreadsheet"
        xSpreadsheet != null

        cleanup: "remove the spreadsheet"
        xSpreadsheets.removeByName("MySheet")

    }

    def "get sheet by index"() {

        when: "we request spreadsheet by index 1"
        XSpreadsheet xSpreadsheet = xSpreadsheetDocument.getSheetByIndex(1)

        then: "we have a spreadsheet"
        xSpreadsheet != null

    }

    def "get active sheet"() {

        given: "a spreadsheet document with Sheet2 active"
        XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets()
        XSpreadsheet xSheet2 = null
        XIndexAccess xSheetsIA = UnoRuntime.queryInterface(XIndexAccess.class, xSpreadsheets)
        xSheet2 = UnoRuntime.queryInterface(XSpreadsheet.class, xSheetsIA.getByIndex(1))

        XModel xModel = UnoRuntime.queryInterface(XModel.class, xSpreadsheetDocument)
        XController xController = xModel.getCurrentController()
        XSpreadsheetView xSpreadsheetView = UnoRuntime.queryInterface(XSpreadsheetView, xController)
        xSpreadsheetView.setActiveSheet(xSheet2)

        when: "we request the active sheet"
        XSpreadsheet xSheet = xSpreadsheetDocument.getActiveSheet()

        and: "we get it's name"
        XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, xSheet)
        String name = xNamed.getName()

        then: "we have a sheet named Sheet2"
        name == "Sheet2"

        cleanup: "set Sheet1 active"
        XSpreadsheet xSheet1 = null
        xSheet1 = UnoRuntime.queryInterface(XSpreadsheet.class, xSheetsIA.getByIndex(0))

    }

    def "set active sheet"() {

        given: "a spreadsheet document with Sheet1 active"

        XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets()
        XSpreadsheet xSheet1 = null
        XIndexAccess xSheetsIA = UnoRuntime.queryInterface(XIndexAccess.class, xSpreadsheets)
        xSheet1 = UnoRuntime.queryInterface(XSpreadsheet.class, xSheetsIA.getByIndex(0))

        XModel xModel = UnoRuntime.queryInterface(XModel.class, xSpreadsheetDocument)
        XController xController = xModel.getCurrentController()
        XSpreadsheetView xSpreadsheetView = UnoRuntime.queryInterface(XSpreadsheetView, xController)
        xSpreadsheetView.setActiveSheet(xSheet1)

        and: "we get a reference to Sheet2 to use it"
        XSpreadsheet xSheet2 = null
        xSheet2 = UnoRuntime.queryInterface(XSpreadsheet.class, xSheetsIA.getByIndex(1))

        when: "we set the active sheet to Sheet2"
        xSpreadsheetDocument.setActiveSheet(xSheet2)

        and: "get the active sheet's name"
        XSpreadsheet xSheet = null
        xSheet =  xSpreadsheetView.getActiveSheet()
        XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, xSheet)
        String name = xNamed.getName()

        then: "the active sheet is Sheet2"
        name == "Sheet2"

        cleanup: "set Sheet1 active"
        xSpreadsheetView.setActiveSheet(xSheet1)

    }


    // helper methods

    @Deprecated
    // Connect to a running office that is accepting connections.
    void connect() {
        if (mxRemoteContext == null && mxRemoteServiceManager == null) {
            try {
                // added to bootstrap by file path
                String oooExeFolder = "/opt/openoffice4/program/"

                // First step: get the remote office component context
                // mxRemoteContext = com.sun.star.comp.helper.Bootstrap.bootstrap()

                // we use the bootstrapconnector for tests
                mxRemoteContext = BootstrapSocketConnector.bootstrap(oooExeFolder);
                // mxRemoteContext = BootstrapPipeConnector.bootstrap(oooExeFolder);
                System.out.println("Connected to a running office ...")

                mxRemoteServiceManager = mxRemoteContext.getServiceManager()

            } catch (Exception e) {
                System.err.println("ERROR: can't get a component context from a running office ...")
                e.printStackTrace()
                System.exit(1)
            }
        }
    }

    @Deprecated
    /** Creates an empty spreadsheet document.
     @return The XSpreadsheetDocument interface of the document.  */
    XSpreadsheetDocument initDocument()
            throws RuntimeException, Exception {
        XComponentLoader aLoader = UnoRuntime.queryInterface(
                XComponentLoader.class,
                mxRemoteServiceManager.createInstanceWithContext(
                        "com.sun.star.frame.Desktop", mxRemoteContext))

        // changed to use class var xComponent
        xComponent = aLoader.loadComponentFromURL(
                "private:factory/scalc", "_default", 0, new com.sun.star.beans.PropertyValue[0])

        XSpreadsheetDocument xSpreadsheetDocument = UnoRuntime.queryInterface(
                XSpreadsheetDocument.class, xComponent)

        return xSpreadsheetDocument
    }


}

