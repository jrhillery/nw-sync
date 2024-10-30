/*
 * Created on Dec 2, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.table.CellContentType.VALUE;
import static com.sun.star.uno.UnoRuntime.queryInterface;
import static com.sun.star.util.NumberFormat.CURRENCY;
import static com.sun.star.util.NumberFormat.DATE;
import static com.sun.star.util.NumberFormat.PERCENT;
import static com.sun.star.util.NumberFormat.UNDEFINED;
import static java.time.temporal.ChronoUnit.DAYS;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import com.leastlogic.moneydance.util.MdLog;
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
	private final String urlString;
	private final XNumberFormats numberFormats;
	private final LocalDate zeroDate;

	private final List<CellHandler> changes = new ArrayList<>();

	/**
	 * Sole constructor.
	 *
	 * @param spreadsheetDoc Spreadsheet document
	 */
	public CalcDoc(XSpreadsheetDocument spreadsheetDoc) throws MduException {
		this.spreadsheetDoc = spreadsheetDoc;
		this.urlString = queryInterface(XModel.class, spreadsheetDoc).getURL();
		this.numberFormats = queryInterface(XNumberFormatsSupplier.class, spreadsheetDoc)
			.getNumberFormats();
		XPropertySet docProps = queryInterface(XPropertySet.class, spreadsheetDoc);
		if (docProps == null)
			throw new MduException(null, "Unable to obtain properties for %s", this.urlString);

		Date nullDate;
		try {
			nullDate = (Date) docProps.getPropertyValue("NullDate");
		} catch (Exception e) {
			throw new MduException(e, "Exception obtaining NullDate for %s", this.urlString);
		}
		if (nullDate == null)
			throw new MduException(null, "Unable to obtain NullDate for %s", this.urlString);

		this.zeroDate = LocalDate.of(nullDate.Year, nullDate.Month, nullDate.Day);

	} // end (XSpreadsheetDocument) constructor

	/**
	 * @return A row iterator for the first sheet in the spreadsheet document
	 */
	public XEnumeration getFirstSheetRowIterator() throws MduException {
		XIndexAccess sheetIndex = getSheets();
		if (sheetIndex == null)
			throw new MduException(null, "Unable to index sheets in %s", this.urlString);
		XSpreadsheet firstSheet;

		try {
			firstSheet = queryInterface(XSpreadsheet.class, sheetIndex.getByIndex(0));
		} catch (Exception e) {
			throw new MduException(e, "Exception obtaining first sheet in %s", this.urlString);
		}
		if (firstSheet == null)
			throw new MduException(null, "Unable to obtain first sheet in %s", this.urlString);

		// get a cursor, so we don't iterator over all the empty rows at the bottom
		XUsedAreaCursor cur = queryInterface(XUsedAreaCursor.class, firstSheet.createCursor());
		if (cur == null)
			throw new MduException(null, "Unable to get cursor in %s", this.urlString);

		cur.gotoStartOfUsedArea(false); // set the range to a single cell
		cur.gotoEndOfUsedArea(true); // expand range to include all used area
		XColumnRowRange colRowRange = queryInterface(XColumnRowRange.class, cur);
		if (colRowRange == null)
			throw new MduException(null, "Unable to get column row range in %s", this.urlString);

		XEnumerationAccess rowAccess = queryInterface(XEnumerationAccess.class,
				colRowRange.getRows());
		if (rowAccess == null)
			throw new MduException(null, "Unable to get row enumeration access in %s", this.urlString);

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
	public XPropertySet getNumberFormatProps(XCell cell) {
		XPropertySet cellNumberFormatProps = null;
		XPropertySet cellProps = queryInterface(XPropertySet.class, cell);

		if (cellProps != null) {
			try {
				cellNumberFormatProps = this.numberFormats
					.getByKey((Integer) cellProps.getPropertyValue("NumberFormat"));
			} catch (Exception e) {
				MdLog.all("Problem obtaining cell number format", e);
			}
		}

		return cellNumberFormatProps;
	} // end getNumberFormatProps(XCell)

	/**
	 * @param cell The cell to read
	 * @return The cell's number format type
	 */
	private short getNumberFormatType(XCell cell) {
		XPropertySet cellNumberFormatProps = getNumberFormatProps(cell);

		if (cellNumberFormatProps == null)
			return UNDEFINED;

		try {
			return (Short) cellNumberFormatProps.getPropertyValue("Type");
		} catch (Exception e) {
			MdLog.all("Problem obtaining type of cell number format", e);

			return UNDEFINED;
		}
	} // end getNumberFormatType(XCell)

	/**
	 * @param row   Office Row instance
	 * @param index Zero-based index to use
	 * @return CellHandler instance for the cell at zero-based index in the supplied row
	 */
	public CellHandler getCellHandlerByIndex(XCellRange row, int index) {
		XCell cell = getCellByIndex(row, index);

		if (isContentType(cell, VALUE)) {
			short numberFormatType = getNumberFormatType(cell);

			return (numberFormatType & PERCENT) != 0
				? null
				: (numberFormatType & DATE) != 0
				? new DateCellHandler(cell, this)
				: new FloatCellHandler(cell, this, (numberFormatType & CURRENCY) != 0);
		}

		return null;
	} // end getCellHandlerByIndex(XCellRange, int)

	/**
	 * @param row   Office Row instance
	 * @param index Zero-based index to use
	 * @return Cell at zero-based index in the supplied row
	 */
	public XCell getCellByIndex(XCellRange row, int index) {
		try {

			return row.getCellByPosition(index, 0);
		} catch (Exception e) {
			MdLog.all("Problem obtaining cell %d in row: %s".formatted(index, row), e);

			return null;
		}
	} // end getCellByIndex(XCellRange, int)

} // end class CalcDoc
