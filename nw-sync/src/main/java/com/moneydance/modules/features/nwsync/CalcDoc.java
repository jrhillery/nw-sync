/*
 * Created on Dec 2, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static com.sun.star.util.NumberFormat.DATE;
import static com.sun.star.util.NumberFormat.UNDEFINED;
import static java.time.temporal.ChronoUnit.DAYS;

import java.io.InputStream;
import java.lang.reflect.Method;
import java.net.URL;
import java.net.URLClassLoader;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import com.moneydance.modules.features.nwsync.CellHandler.DateCellHandler;
import com.moneydance.modules.features.nwsync.CellHandler.FloatCellHandler;
import com.moneydance.modules.features.nwsync.OdsAccessor.OdsException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.frame.XDesktop2;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.table.CellContentType;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.Date;
import com.sun.star.util.XNumberFormats;
import com.sun.star.util.XNumberFormatsSupplier;

/**
 * Class to hold attributes of a spreadsheet document.
 */
public class CalcDoc {

	private XSpreadsheetDocument spreadSheetDoc;
	private String urlString;
	private XNumberFormats numberFormats;
	private LocalDate zeroDate;

	private XSpreadsheet firstSheet = null;
	private List<CellHandler> changes = new ArrayList<>();

	private static Properties nwSyncProps = null;
	private static boolean classPathUpdated = false;

	private static final String propertiesFileName = "nw-sync.properties";

	/**
	 * Sole constructor.
	 *
	 * @param spreadSheetDoc
	 */
	public CalcDoc(XSpreadsheetDocument spreadSheetDoc) throws OdsException {
		this.spreadSheetDoc = spreadSheetDoc;
		this.urlString = queryInterface(XModel.class, spreadSheetDoc).getURL();
		this.numberFormats = queryInterface(XNumberFormatsSupplier.class, spreadSheetDoc)
				.getNumberFormats();
		XPropertySet docProps = queryInterface(XPropertySet.class, spreadSheetDoc);
		if (docProps == null)
			// Unable to obtain properties for %s.
			throw new OdsException(null, "NWSYNC30", this.urlString);

		Date nullDate;
		try {
			nullDate = (Date) docProps.getPropertyValue("NullDate");
		} catch (Exception e) {
			// Exception obtaining NullDate for %s.
			throw new OdsException(e, "NWSYNC31", this.urlString);
		}
		if (nullDate == null)
			// Unable to obtain NullDate for %s.
			throw new OdsException(null, "NWSYNC32", this.urlString);

		this.zeroDate = LocalDate.of(nullDate.Year, nullDate.Month, nullDate.Day);

	} // end (XSpreadsheetDocument) constructor

	/**
	 * @return a list of currently open spreadsheet documents
	 */
	public static List<XSpreadsheetDocument> getSpreadsheetDocs() throws OdsException {
		List<XSpreadsheetDocument> docList = new ArrayList<>();
		XDesktop2 libreOfficeDesktop = getOfficeDesktop();
		XEnumeration compItr = libreOfficeDesktop.getComponents().createEnumeration();

		if (!compItr.hasMoreElements()) {
			// no components so we probably started the desktop => terminate it
			libreOfficeDesktop.terminate();
		} else {
			do {
				XServiceInfo comp = next(XServiceInfo.class, compItr);

				if (comp.supportsService("com.sun.star.sheet.SpreadsheetDocument")) {
					docList.add(queryInterface(XSpreadsheetDocument.class, comp));
				}
			} while (compItr.hasMoreElements());
		}

		return docList;
	} // end getSpreadsheetDocs()

	/**
	 * Modify the system class loader's class path to add the required
	 * LibreOffice jar files. Need to use reflection since the
	 * URLClassLoader.addURL(URL) method is protected.
	 *
	 * @param officePath location of the installed LibreOffice jar files
	 */
	private static void addOfficeApiToClassPath(Path officePath) throws OdsException {
		URLClassLoader sysClassLoader = (URLClassLoader) ClassLoader.getSystemClassLoader();
		Method addURL;
		try {
			addURL = URLClassLoader.class.getDeclaredMethod("addURL", new Class[] { URL.class });
		} catch (Exception e) {
			// Exception obtaining class loader addURL method.
			throw new OdsException(e, "NWSYNC49");
		}
		addURL.setAccessible(true);
		for (String apiJar : new String[] { "juh.jar", "jurt.jar", "ridl.jar", "unoil.jar",
				"unoloader.jar" }) {
			URL apiUrl;
			try {
				apiUrl = officePath.resolve(apiJar).toUri().toURL();
			} catch (Exception e) {
				// Exception obtaining URL to jar %s in path %s.
				throw new OdsException(e, "NWSYNC50", apiJar, officePath);
			}
			try {
				addURL.invoke(sysClassLoader, apiUrl);
			} catch (Exception e) {
				// Exception adding %s to class path.
				throw new OdsException(e, "NWSYNC51", apiUrl);
			}
		} // end for

	} // end addOfficeApiToClassPath(Path)

	/**
	 * @return a LibreOffice desktop interface
	 */
	private static XDesktop2 getOfficeDesktop() throws OdsException {
		if (!classPathUpdated) {
			String officeInstallPath = getNwSyncProps().getProperty("office.install.path");
			if (officeInstallPath == null)
				// Unable to obtain office.install.path from %s on the class path.
				throw new OdsException(null, "NWSYNC54", propertiesFileName);

			addOfficeApiToClassPath(Paths.get(officeInstallPath, "program/classes"));
			classPathUpdated = true;
		}
		XComponentContext remoteContext;
		try {
			remoteContext = Bootstrap.bootstrap();
		} catch (Exception e) {
			// Exception obtaining office context.
			throw new OdsException(e, "NWSYNC33");
		}
		if (remoteContext == null)
			// Unable to obtain office context.
			throw new OdsException(null, "NWSYNC34");

		XMultiComponentFactory remoteServiceMgr = remoteContext.getServiceManager();
		if (remoteServiceMgr == null)
			// Unable to obtain office service manager.
			throw new OdsException(null, "NWSYNC35");

		XDesktop2 libreOfficeDesktop;
		try {
			libreOfficeDesktop = queryInterface(XDesktop2.class, remoteServiceMgr
				.createInstanceWithContext("com.sun.star.frame.Desktop", remoteContext));
		} catch (Exception e) {
			// Exception obtaining office desktop.
			throw new OdsException(e, "NWSYNC36");
		}
		if (libreOfficeDesktop == null)
			// Unable to obtain office desktop.
			throw new OdsException(null, "NWSYNC37");

		return libreOfficeDesktop;
	} // end getOfficeDesktop()

	/**
	 * Close our connection to the office process.
	 */
	public static void closeOfficeConnection() {
		try {
			// get the bridge factory from the local service manager
			XBridgeFactory bridgeFactory = queryInterface(XBridgeFactory.class,
				Bootstrap.createSimpleServiceManager()
					.createInstance("com.sun.star.bridge.BridgeFactory"));

			if (bridgeFactory != null) {
				for (XBridge bridge : bridgeFactory.getExistingBridges()) {
					// dispose of this bridge after closing its connection
					queryInterface(XComponent.class, bridge).dispose();
				}
			}
		} catch (Throwable e) {
			System.err.println("Exception disposing office process connection bridge:");
			e.printStackTrace(System.err);
		}

	} // end closeOfficeConnection()

	/**
	 * @return a row iterator for the first sheet in the spreadsheet document
	 */
	public XEnumeration getFirstSheetRowIterator() throws OdsException {
		if (this.firstSheet == null) {
			XIndexAccess sheetIndex = queryInterface(XIndexAccess.class,
					this.spreadSheetDoc.getSheets());
			if (sheetIndex == null)
				// Unable to index sheets in %s.
				throw new OdsException(null, "NWSYNC38", this.urlString);

			try {
				this.firstSheet = queryInterface(XSpreadsheet.class, sheetIndex.getByIndex(0));
			} catch (Exception e) {
				// Exception obtaining first sheet in %s.
				throw new OdsException(e, "NWSYNC39", this.urlString);
			}
			if (this.firstSheet == null)
				// Unable to obtain first sheet in %s.
				throw new OdsException(null, "NWSYNC40", this.urlString);
		}

		// get a cursor so we don't iterator over all the empty rows at the bottom
		XUsedAreaCursor cur = queryInterface(XUsedAreaCursor.class,
				this.firstSheet.createCursor());
		if (cur == null)
			// Unable to get cursor in %s.
			throw new OdsException(null, "NWSYNC41", this.urlString);

		cur.gotoStartOfUsedArea(false); // set the range to a single cell
		cur.gotoEndOfUsedArea(true); // expand range to include all used area
		XColumnRowRange colRowRange = queryInterface(XColumnRowRange.class, cur);
		if (colRowRange == null)
			// Unable to get column row range in %s.
			throw new OdsException(null, "NWSYNC42", this.urlString);

		XEnumerationAccess rowAccess = queryInterface(XEnumerationAccess.class,
				colRowRange.getRows());
		if (rowAccess == null)
			// Unable to get row enumeration access in %s.
			throw new OdsException(null, "NWSYNC43", this.urlString);

		return rowAccess.createEnumeration();
	} // end getFirstSheetRowIterator()

	/**
	 * @param dateNum date value in spreadsheet cell
	 * @return LocalDate instance corresponding to dateNum
	 */
	public LocalDate getLocalDate(double dateNum) {

		return this.zeroDate.plusDays((long) dateNum);
	} // end getLocalDate(double)

	/**
	 * @param localDate date the cell should represent
	 * @return date number value for the spreadsheet cell
	 */
	public long getDateNumber(LocalDate localDate) {

		return this.zeroDate.until(localDate, DAYS);
	} // end getDateNumber(LocalDate)

	public void addChange(CellHandler cHandler) {
		if (cHandler != null) {
			this.changes.add(cHandler);
		}

	} // end addChange(CellHandler)

	/**
	 * Commit any changes to the spreadsheet document.
	 */
	public void commitChanges() {
		for (CellHandler cHandler : this.changes) {
			cHandler.applyUpdate();
		}
		forgetChanges();

	} // end commitChanges()

	/**
	 * Clear out any pending changes.
	 */
	public void forgetChanges() {
		this.changes.clear();

	} // end forgetChanges()

	/**
	 * @return true when the spreadsheet has uncommitted changes in memory
	 */
	public boolean isModified() {

		return !this.changes.isEmpty();
	} // end isModified()

	/**
	 * @param zInterface
	 * @param iterator
	 * @return the next zInterface from the iterator
	 */
	public static <T> T next(Class<T> zInterface, XEnumeration iterator) throws OdsException {
		T rslt;
		try {
			rslt = queryInterface(zInterface, iterator.nextElement());
		} catch (Exception e) {
			// Exception iterating to next %s.
			throw new OdsException(e, "NWSYNC44", zInterface.getSimpleName());
		}
		if (rslt == null)
			// Unable to obtain next %s.
			throw new OdsException(null, "NWSYNC45", zInterface.getSimpleName());

		return rslt;
	} // end next(Class<T>, XEnumeration)

	/**
	 * @return a string representation of this CalcDoc
	 */
	public String toString() {

		return this.urlString;
	} // end toString()

	/**
	 * @param cell
	 * @param contentType
	 * @return true when the supplied cell has the specified content type
	 */
	public static boolean isValueType(XCell cell, CellContentType contentType) {

		return cell != null && contentType.equals(cell.getType());
	} // end isValueType(XCell, CellContentType)

	/**
	 * @param cell
	 * @return the cell's number format properties
	 */
	public XPropertySet getNumberFormatProps(XCell cell) throws OdsException {
		XPropertySet cellNumberFormatProps = null;
		XPropertySet cellProps = queryInterface(XPropertySet.class, cell);

		if (cellProps != null) {
			try {
				cellNumberFormatProps = this.numberFormats
					.getByKey((Integer) cellProps.getPropertyValue("NumberFormat"));
			} catch (Exception e) {
				// Exception obtaining cell number format.
				throw new OdsException(e, "NWSYNC46");
			}
		}

		return cellNumberFormatProps;
	} // end getNumberFormatProps(XCell)

	/**
	 * @param cell
	 * @return the cell's number format type
	 */
	private short getNumberFormatType(XCell cell) throws OdsException {
		XPropertySet cellNumberFormatProps = getNumberFormatProps(cell);
		if (cellNumberFormatProps == null)
			return UNDEFINED;

		try {
			return (Short) cellNumberFormatProps.getPropertyValue("Type");
		} catch (Exception e) {
			// Exception obtaining number format type.
			throw new OdsException(e, "NWSYNC47");
		}
	} // end getNumberFormatType(XCell)

	/**
	 * @param row office Row instance
	 * @param index
	 * @return CellHandler instance for the cell at zero-based index in the supplied row
	 */
	public CellHandler getCellHandlerByIndex(XCellRange row, int index) throws OdsException {
		XCell cell = getCellByIndex(row, index);

		if (isValueType(cell, VALUE)) {
			short numberFormatType = getNumberFormatType(cell);

			return (numberFormatType & DATE) != 0
				? new DateCellHandler(cell, this)
				: new FloatCellHandler(cell, this);
		}

		return null;
	} // end getCellHandlerByIndex(XCellRange, int)

	/**
	 * @param row office Row instance
	 * @param index
	 * @return cell at zero-based index in the supplied row
	 */
	public static XCell getCellByIndex(XCellRange row, int index) throws OdsException {
		try {

			return row.getCellByPosition(index, 0);
		} catch (Exception e) {
			// Exception obtaining cell in row.
			throw new OdsException(e, "NWSYNC48");
		}
	} // end getCellByIndex(XCellRange, int)

	/**
	 * @return our properties
	 */
	public static Properties getNwSyncProps() throws OdsException {
		if (nwSyncProps == null) {
			InputStream propsStream = CalcDoc.class.getClassLoader()
				.getResourceAsStream(propertiesFileName);
			if (propsStream == null)
				// Unable to find %s on the class path.
				throw new OdsException(null, "NWSYNC52", propertiesFileName);

			nwSyncProps = new Properties();
			try {
				nwSyncProps.load(propsStream);
			} catch (Exception e) {
				nwSyncProps = null;

				// Exception loading %s.
				throw new OdsException(e, "NWSYNC53", propertiesFileName);
			} finally {
				try {
					propsStream.close();
				} catch (Exception e) { /* ignore */ }
			}
		}

		return nwSyncProps;
	} // end getNwSyncProps()

} // end class CalcDoc
