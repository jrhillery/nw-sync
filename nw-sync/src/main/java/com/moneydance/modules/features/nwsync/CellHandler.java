/*
 * Created on Nov 18, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.uno.UnoRuntime.queryInterface;

import java.text.NumberFormat;
import java.time.LocalDate;

import com.sun.star.sheet.XCellAddressable;
import com.sun.star.table.CellAddress;
import com.sun.star.table.XCell;
import com.sun.star.text.XText;

/**
 * Abstract read and write access to numeric spreadsheet cells.
 */
public abstract class CellHandler {

	/**
	 * Provide read and write access to currency spreadsheet cells.
	 */
	public static class CurrencyCellHandler extends CellHandler {
		public CurrencyCellHandler(XCell cell, CalcDoc calcDoc) {
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

		public NumberFormat getNumberFormat() {
			return NumberFormat.getCurrencyInstance();
		} // end getNumberFormat()

	} // end class CurrencyCellHandler

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

		public NumberFormat getNumberFormat() {
			return NumberFormat.getNumberInstance();
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
			LocalDate date = this.calcDoc.getLocalDate(this.cell.getValue());
			int dateInt = date.getYear() * 10000 + date.getMonthValue() * 100
					+ date.getDayOfMonth();

			return dateInt;
		} // end getValue()

		/**
		 * @param value the numeric date value (in decimal form YYYYMMDD) to save in this cell
		 */
		public void setValue(Number value) {
			if (value != null) {
				LocalDate localDate = convDateIntToLocal(value.intValue());
				long dateNum = this.calcDoc.getDateNumber(localDate);
				this.cell.setValue(dateNum);
			}

		} // end setValue(Number)

		public NumberFormat getNumberFormat() {
			return NumberFormat.getIntegerInstance();
		} // end getNumberFormat()

	} // end class DateCellHandler

	protected XCell cell;
	protected CalcDoc calcDoc;
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
	public abstract NumberFormat getNumberFormat();

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
	 * @param cell
	 * @return the text displayed in cell
	 */
	public static String asDisplayText(XCell cell) {
		XText cellText = queryInterface(XText.class, cell);

		return cellText == null ? null : cellText.getString();
	} // end asDisplayText(XCell)

	/**
	 * @param dateInt the numeric date value in decimal form YYYYMMDD
	 * @return the corresponding local date
	 */
	public static LocalDate convDateIntToLocal(int dateInt) {
		int year = dateInt / 10000;
		int month = (dateInt % 10000) / 100;
		int dayOfMonth = dateInt % 100;

		return LocalDate.of(year, month, dayOfMonth);
	} // end convDateIntToLocal(int)

} // end class CellHandler
