using Atlas.Data;
using System;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.ucb;
using unoidl.com.sun.star.uno;

namespace Atlas.Interops.LibreOffice
{
    public class CalcReader : IDocReader
    {
        XComponentContext xContext = Bootstrap.bootstrap();

        public DocAttributes PullAttributes(string pathToFile)
        {
            XMultiComponentFactory xMCF = xContext.getServiceManager();
            object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
            XComponentLoader xLoader = (XComponentLoader)oDesktop;
            PropertyValue[] emptyArgs = new PropertyValue[0];

            string strSheet = @"private:factory/scalc";

            XComponent xComp = xLoader.loadComponentFromURL(strSheet, "_blank", 0, emptyArgs);
            XSpreadsheetDocument doc = (XSpreadsheetDocument)xComp;
            return new DocAttributes();
        }
    }
}
