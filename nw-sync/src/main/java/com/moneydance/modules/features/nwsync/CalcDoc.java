/*
 * Created on Dec 2, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static com.sun.star.util.NumberFormat.CURRENCY;
import static com.sun.star.util.NumberFormat.DATE;
import static java.time.temporal.ChronoUnit.DAYS;

import java.time.LocalDate;

import com.moneydance.modules.features.nwsync.CellHandler.CurrencyCellHandler;
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
	 * @return a Libre office desktop interface
	 */
	public static XDesktop2 getOfficeDesktop() throws OdsException {
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
	 * @return the first sheet in the spreadsheet document
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
		} catch (Exception e) {
			System.err.println("Exception disposing office process connection bridge:");
			e.printStackTrace(System.err);
		}

	} // end closeOfficeConnection()

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
		CellContentType type = cell.getType();

		return cell != null && contentType.equals(type);
	} // end isValueType(XCell, CellContentType)

	/**
	 * @param cell
	 * @param numberFormatType
	 * @return true if the cell has the specified number format type
	 */
	private boolean isFormat(XCell cell, short numberFormatType) throws OdsException {
		XPropertySet cellProps = queryInterface(XPropertySet.class, cell);
		if (cellProps == null)
			return false;

		XPropertySet cellNumberFormatProps;
		try {
			cellNumberFormatProps = this.numberFormats
					.getByKey((Integer) cellProps.getPropertyValue("NumberFormat"));
		} catch (Exception e) {
			// Exception obtaining cell number format.
			throw new OdsException(e, "NWSYNC46");
		}
		if (cellNumberFormatProps == null)
			return false;

		try {
			return numberFormatType == (Short) cellNumberFormatProps.getPropertyValue("Type");
		} catch (Exception e) {
			// Exception obtaining number format type.
			throw new OdsException(e, "NWSYNC47");
		}
	} // end isFormat(XCell, short)

	/**
	 * @param row office Row instance
	 * @param index
	 * @return CellHandler instance for the cell at zero-based index in the supplied row
	 */
	public CellHandler getCellHandlerByIndex(XCellRange row, int index) throws OdsException {
		XCell cell = getCellByIndex(row, index);

		if (isValueType(cell, VALUE)) {
			if (isFormat(cell, DATE))
				return new DateCellHandler(cell, this);

			if (isFormat(cell, CURRENCY))
				return new CurrencyCellHandler(cell, this);

			return new FloatCellHandler(cell, this);
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

} // end class CalcDoc
