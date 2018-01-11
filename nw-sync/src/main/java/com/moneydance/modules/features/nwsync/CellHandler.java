/*
 * Created on Nov 18, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.uno.UnoRuntime.queryInterface;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.time.LocalDate;

import com.johns.moneydance.util.MdUtil;
import com.moneydance.modules.features.nwsync.OdsAccessor.OdsException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.table.CellAddress;
import com.sun.star.table.XCell;
import com.sun.star.text.XText;

/**
 * Abstract read and write access to numeric spreadsheet cells.
 */
public abstract class CellHandler {

	/**
	 * Provide read and write access to floating point spreadsheet cells.
	 */
	public static class FloatCellHandler extends CellHandler {
		public FloatCellHandler(XCell cell, CalcDoc calcDoc) {
			super(cell, calcDoc);

		} // end (XCell, CalcDoc) constructor

		public Double getValue() {
			return this.cell.getValue();
		} // end getValue()

		public void setValue(Number value) {
			if (value != null) {
				this.cell.setValue(value.doubleValue());
			}

		} // end setValue(Number)

		public NumberFormat getNumberFormat() throws OdsException {
			if (this.numberFormat == null) {
				this.numberFormat = createDecimalFormat();
			}
			return this.numberFormat;
		} // end getNumberFormat()

	} // end class FloatCellHandler

	/**
	 * Provide read and write access to date spreadsheet cells.
	 */
	public static class DateCellHandler extends CellHandler {
		public DateCellHandler(XCell cell, CalcDoc calcDoc) {
			super(cell, calcDoc);

		} // end (XCell, CalcDoc) constructor

		/**
		 * @return the numeric date value of this cell in decimal form YYYYMMDD
		 */
		public Integer getValue() {
			return MdUtil.convLocalToDateInt(getDateValue());
		} // end getValue()

		/**
		 * @return the local date value of this cell
		 */
		public LocalDate getDateValue() {

			return this.calcDoc.getLocalDate(this.cell.getValue());
		} // end getDateValue()

		/**
		 * @param value the numeric date value (in decimal form YYYYMMDD) to save in this cell
		 */
		public void setValue(Number value) {
			if (value != null) {
				LocalDate localDate = MdUtil.convDateIntToLocal(value.intValue());
				long dateNum = this.calcDoc.getDateNumber(localDate);
				this.cell.setValue(dateNum);
			}

		} // end setValue(Number)

		public NumberFormat getNumberFormat() {
			if (this.numberFormat == null) {
				this.numberFormat = (DecimalFormat) NumberFormat.getIntegerInstance();
			}
			return this.numberFormat;
		} // end getNumberFormat()

	} // end class DateCellHandler

	protected XCell cell;
	protected CalcDoc calcDoc;
	protected DecimalFormat numberFormat = null;
	private Number newValue = null;

	/**
	 * Sole constructor.
	 *
	 * @param cell office Cell instance to be handled
	 * @param calcDoc local spreadsheet document containing this cell
	 */
	public CellHandler(XCell cell, CalcDoc calcDoc) {
		this.cell = cell;
		this.calcDoc = calcDoc;

	} // end (XCell, CalcDoc) constructor

	/**
	 * @return the numeric value of this cell
	 */
	public abstract Number getValue();

	/**
	 * @param value the value to save in this cell
	 */
	public abstract void setValue(Number value);

	/**
	 * @return a formatter for values in this cell
	 */
	public abstract NumberFormat getNumberFormat() throws OdsException;

	/**
	 * @return the formula string of this cell
	 */
	public String getFormula() {

		return this.cell.getFormula();
	} // end getFormula()

	/**
	 * @return the text displayed in this cell
	 */
	public String getDisplayText() {

		return asDisplayText(this.cell);
	} // end getDisplayText()

	/**
	 * @param newValue new value to save for later application
	 */
	public void setNewValue(Number newValue) {
		this.newValue = newValue;
		this.calcDoc.addChange(this);

	} // end setNewValue(Number)

	/**
	 * Apply any previously set new value.
	 */
	public void applyUpdate() {
		setValue(this.newValue);

	} // end applyUpdate()

	/**
	 * @return a string representation of this CellHandler
	 */
	public String toString() {
		StringBuilder sb = new StringBuilder(getDisplayText());
		XCellAddressable addressable = queryInterface(XCellAddressable.class, this.cell);
		CellAddress cellAdr = addressable == null ? null : addressable.getCellAddress();

		if (cellAdr != null) {
			sb.append('[').append(cellAdr.Column)
			  .append(", ").append(cellAdr.Row)
			  .append(", ").append(cellAdr.Sheet).append(']');
		}

		return sb.toString();
	} // end toString()

	/**
	 * @return a decimal format for this cell
	 */
	protected DecimalFormat createDecimalFormat() throws OdsException {
		XPropertySet numberFormatProps = this.calcDoc.getNumberFormatProps(this.cell);

		if (numberFormatProps != null) {
			String fmtString;
			try {
				fmtString = (String) numberFormatProps.getPropertyValue("FormatString");
			} catch (Exception e) {
				// Exception obtaining number format string in cell %s.
				throw new OdsException(e, "NWSYNC55", this);
			}
			if (fmtString != null && !fmtString.equals("General")) {
				// isolate the positive subpattern
				String[] patternParts = fmtString.split(";");
				// transform any currency symbols
				String pattern = patternParts[0].replace("[$$-409]", "$");
				// System.err.println("Spreadsheet format string [" + fmtString
				// + "]\t Number format pattern [" + pattern + ']');
				DecimalFormat df = new DecimalFormat();
				df.applyLocalizedPattern(pattern);

				return df;
			}
		}

		return (DecimalFormat) NumberFormat.getNumberInstance();
	} // end createDecimalFormat()

	/**
	 * @param cell
	 * @return the text displayed in cell
	 */
	public static String asDisplayText(XCell cell) {
		XText cellText = queryInterface(XText.class, cell);

		return cellText == null ? null : cellText.getString();
	} // end asDisplayText(XCell)

} // end class CellHandler
