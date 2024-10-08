/*
 * Created on Nov 18, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.sun.star.uno.UnoRuntime.queryInterface;

import java.time.LocalDate;

import com.leastlogic.moneydance.util.MdUtil;
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
		private final boolean isCurrency;

		public FloatCellHandler(XCell cell, CalcDoc calcDoc, boolean isCurrency) {
			super(cell, calcDoc);
			this.isCurrency = isCurrency;

		} // end constructor

		/**
		 * @return The numeric value of this cell as a Double
		 */
		public Double getValue() {
			return this.cell.getValue();
		} // end getValue()

		/**
		 * @param value The double value to save in this cell
		 */
		public void setValue(Number value) {
			if (value != null) {
				this.cell.setValue(value.doubleValue());
			}

		} // end setValue(Number)

		/**
		 * @return true when this cell's number format type is currency
		 */
		public boolean isCurrency() {
			return this.isCurrency;
		} // end isCurrency()

	} // end class FloatCellHandler

	/**
	 * Provide read and write access to date spreadsheet cells.
	 */
	public static class DateCellHandler extends CellHandler {
		public DateCellHandler(XCell cell, CalcDoc calcDoc) {
			super(cell, calcDoc);

		} // end (XCell, CalcDoc) constructor

		/**
		 * @return The numeric date value of this cell in decimal form YYYYMMDD
		 */
		public Integer getValue() {
			return MdUtil.convLocalToDateInt(getDateValue());
		} // end getValue()

		/**
		 * @return The local date value of this cell
		 */
		public LocalDate getDateValue() {

			return this.calcDoc.getLocalDate(this.cell.getValue());
		} // end getDateValue()

		/**
		 * @param value The numeric date value (in decimal form YYYYMMDD) to save in this cell
		 */
		public void setValue(Number value) {
			if (value != null) {
				LocalDate localDate = MdUtil.convDateIntToLocal(value.intValue());
				long dateNum = this.calcDoc.getDateNumber(localDate);
				this.cell.setValue(dateNum);
			}

		} // end setValue(Number)

		/**
		 * @return false since this is a date cell
		 */
		public boolean isCurrency() {
			return false;
		} // end isCurrency()

	} // end class DateCellHandler

	protected XCell cell;
	protected CalcDoc calcDoc;
	private Number newValue = null;

	/**
	 * Sole constructor.
	 *
	 * @param cell    Office Cell instance to be handled
	 * @param calcDoc Local spreadsheet document containing this cell
	 */
	public CellHandler(XCell cell, CalcDoc calcDoc) {
		this.cell = cell;
		this.calcDoc = calcDoc;

	} // end (XCell, CalcDoc) constructor

	/**
	 * @return The numeric value of this cell
	 */
	public abstract Number getValue();

	/**
	 * @param value The value to save in this cell
	 */
	public abstract void setValue(Number value);

	/**
	 * @return true when this cell's number format type is currency
	 */
	public abstract boolean isCurrency();

	/**
	 * @param newValue New value to save for later application
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
	 * @return A string representation of this CellHandler
	 */
	public String toString() {
		StringBuilder sb = new StringBuilder(asDisplayText(this.cell));
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
	 * @param cell The cell to read
	 * @return The text displayed in cell
	 */
	public static String asDisplayText(XCell cell) {
		XText cellText = queryInterface(XText.class, cell);

		return cellText == null ? "" : cellText.getString();
	} // end asDisplayText(XCell)

} // end class CellHandler
