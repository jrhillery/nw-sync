/*
 * Created on Dec 2, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static com.sun.star.util.NumberFormat.DATE;
import static com.sun.star.util.NumberFormat.PERCENT;
import static com.sun.star.util.NumberFormat.UNDEFINED;
import static java.time.temporal.ChronoUnit.DAYS;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import com.leastlogic.moneydance.util.MduException;
import com.moneydance.modules.features.nwsync.CellHandler.DateCellHandler;
import com.moneydance.modules.features.nwsync.CellHandler.FloatCellHandler;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.frame.XModel;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.table.CellContentType;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.util.Date;
import com.sun.star.util.XNumberFormats;
import com.sun.star.util.XNumberFormatsSupplier;

/**
 * Class to hold attributes of a spreadsheet document.
 */
public class CalcDoc {

	private final XSpreadsheetDocument spreadsheetDoc;
	private final MessageBundleProvider msgProvider;
	private final String urlString;
	private final XNumberFormats numberFormats;
	private final LocalDate zeroDate;

	private final List<CellHandler> changes = new ArrayList<>();

	/**
	 * Sole constructor.
	 *
	 * @param spreadsheetDoc Spreadsheet document
	 * @param msgProvider    Message bundle provider
	 */
	public CalcDoc(XSpreadsheetDocument spreadsheetDoc,
			MessageBundleProvider msgProvider) throws MduException {
		this.spreadsheetDoc = spreadsheetDoc;
		this.msgProvider = msgProvider;
		this.urlString = queryInterface(XModel.class, spreadsheetDoc).getURL();
		this.numberFormats = queryInterface(XNumberFormatsSupplier.class, spreadsheetDoc)
			.getNumberFormats();
		XPropertySet docProps = queryInterface(XPropertySet.class, spreadsheetDoc);
		if (docProps == null)
			// Unable to obtain properties for %s.
			throw asException(null, "NWSYNC30", this.urlString);

		Date nullDate;
		try {
			nullDate = (Date) docProps.getPropertyValue("NullDate");
		} catch (Exception e) {
			// Exception obtaining NullDate for %s.
			throw asException(e, "NWSYNC31", this.urlString);
		}
		if (nullDate == null)
			// Unable to obtain NullDate for %s.
			throw asException(null, "NWSYNC32", this.urlString);

		this.zeroDate = LocalDate.of(nullDate.Year, nullDate.Month, nullDate.Day);

	} // end (XSpreadsheetDocument) constructor

	/**
	 * @return A row iterator for the first sheet in the spreadsheet document
	 */
	public XEnumeration getFirstSheetRowIterator() throws MduException {
		XIndexAccess sheetIndex = getSheets();
		if (sheetIndex == null)
			// Unable to index sheets in %s.
			throw asException(null, "NWSYNC38", this.urlString);
		XSpreadsheet firstSheet;

		try {
			firstSheet = queryInterface(XSpreadsheet.class, sheetIndex.getByIndex(0));
		} catch (Exception e) {
			// Exception obtaining first sheet in %s.
			throw asException(e, "NWSYNC39", this.urlString);
		}
		if (firstSheet == null)
			// Unable to obtain first sheet in %s.
			throw asException(null, "NWSYNC40", this.urlString);

		// get a cursor, so we don't iterator over all the empty rows at the bottom
		XUsedAreaCursor cur = queryInterface(XUsedAreaCursor.class, firstSheet.createCursor());
		if (cur == null)
			// Unable to get cursor in %s.
			throw asException(null, "NWSYNC41", this.urlString);

		cur.gotoStartOfUsedArea(false); // set the range to a single cell
		cur.gotoEndOfUsedArea(true); // expand range to include all used area
		XColumnRowRange colRowRange = queryInterface(XColumnRowRange.class, cur);
		if (colRowRange == null)
			// Unable to get column row range in %s.
			throw asException(null, "NWSYNC42", this.urlString);

		XEnumerationAccess rowAccess = queryInterface(XEnumerationAccess.class,
				colRowRange.getRows());
		if (rowAccess == null)
			// Unable to get row enumeration access in %s.
			throw asException(null, "NWSYNC43", this.urlString);

		return rowAccess.createEnumeration();
	} // end getFirstSheetRowIterator()

	/**
	 * @return The index access of the sheets in our spreadsheet document
	 */
	public XIndexAccess getSheets() {

		return queryInterface(XIndexAccess.class, this.spreadsheetDoc.getSheets());
	} // end getSheets()

	/**
	 * @param dateNum Date value in spreadsheet cell
	 * @return LocalDate instance corresponding to dateNum
	 */
	public LocalDate getLocalDate(double dateNum) {

		return this.zeroDate.plusDays((long) dateNum);
	} // end getLocalDate(double)

	/**
	 * @param localDate Date the cell should represent
	 * @return Date number value for the spreadsheet cell
	 */
	public long getDateNumber(LocalDate localDate) {

		return this.zeroDate.until(localDate, DAYS);
	} // end getDateNumber(LocalDate)

	/**
	 * Add a cell handler to our list of changes.
	 *
	 * @param cHandler The cell handler to add
	 */
	public void addChange(CellHandler cHandler) {
		if (cHandler != null) {
			this.changes.add(cHandler);
		}

	} // end addChange(CellHandler)

	/**
	 * Commit any changes to the spreadsheet document.
	 */
	public void commitChanges() {
		this.changes.forEach(CellHandler::applyUpdate);

	} // end commitChanges()

	/**
	 * Clear out any pending changes.
	 */
	public void forgetChanges() {
		this.changes.clear();

	} // end forgetChanges()

	/**
	 * @return True when the spreadsheet has uncommitted changes in memory
	 */
	public boolean isModified() {

		return !this.changes.isEmpty();
	} // end isModified()

	/**
	 * @return A string representation of this CalcDoc
	 */
	public String toString() {

		return this.urlString;
	} // end toString()

	/**
	 * @param cell        The cell to check
	 * @param contentType The content type of interest
	 * @return True when the supplied cell has the specified content type
	 */
	public static boolean isContentType(XCell cell, CellContentType contentType) {

		return cell != null && contentType.equals(cell.getType());
	} // end isContentType(XCell, CellContentType)

	/**
	 * @param cell The cell to read
	 * @return The cell's number format properties
	 */
	public XPropertySet getNumberFormatProps(XCell cell) throws MduException {
		XPropertySet cellNumberFormatProps = null;
		XPropertySet cellProps = queryInterface(XPropertySet.class, cell);

		if (cellProps != null) {
			try {
				cellNumberFormatProps = this.numberFormats
					.getByKey((Integer) cellProps.getPropertyValue("NumberFormat"));
			} catch (Exception e) {
				// Exception obtaining cell number format.
				throw asException(e, "NWSYNC46");
			}
		}

		return cellNumberFormatProps;
	} // end getNumberFormatProps(XCell)

	/**
	 * @param cell The cell to read
	 * @return The cell's number format type
	 */
	private short getNumberFormatType(XCell cell) throws MduException {
		XPropertySet cellNumberFormatProps = getNumberFormatProps(cell);
		if (cellNumberFormatProps == null)
			return UNDEFINED;

		try {
			return (Short) cellNumberFormatProps.getPropertyValue("Type");
		} catch (Exception e) {
			// Exception obtaining number format type.
			throw asException(e, "NWSYNC47");
		}
	} // end getNumberFormatType(XCell)

	/**
	 * @param row   Office Row instance
	 * @param index Zero-based index to use
	 * @return CellHandler instance for the cell at zero-based index in the supplied row
	 */
	public CellHandler getCellHandlerByIndex(XCellRange row, int index) throws MduException {
		XCell cell = getCellByIndex(row, index);

		if (isContentType(cell, VALUE)) {
			short numberFormatType = getNumberFormatType(cell);

			return (numberFormatType & PERCENT) != 0
				? null
				: (numberFormatType & DATE) != 0
				? new DateCellHandler(cell, this)
				: new FloatCellHandler(cell, this);
		}

		return null;
	} // end getCellHandlerByIndex(XCellRange, int)

	/**
	 * @param row   Office Row instance
	 * @param index Zero-based index to use
	 * @return Cell at zero-based index in the supplied row
	 */
	public XCell getCellByIndex(XCellRange row, int index) throws MduException {
		try {

			return row.getCellByPosition(index, 0);
		} catch (Exception e) {
			// Exception obtaining cell in row.
			throw asException(e, "NWSYNC48");
		}
	} // end getCellByIndex(XCellRange, int)

	/**
	 * @param cause  Exception that caused this (null if none)
	 * @param key    The resource bundle key (or message)
	 * @param params Optional parameters for the detail message
	 * @return An exception with the supplied data
	 */
	private MduException asException(Throwable cause, String key, Object... params) {

		return new MduException(cause, this.msgProvider.retrieveMessage(key), params);
	} // end asException(Throwable, String, Object...)

} // end class CalcDoc
