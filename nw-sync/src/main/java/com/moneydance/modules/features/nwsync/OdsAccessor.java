/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.infinitekind.moneydance.model.Account.AccountType.CREDIT_CARD;
import static org.odftoolkit.odfdom.dom.attribute.office.OfficeValueTypeAttribute.Value.STRING;

import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import java.util.ResourceBundle;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;
import org.odftoolkit.simple.table.Table;

import com.infinitekind.moneydance.model.Account;
import com.infinitekind.moneydance.model.AccountBook;
import com.infinitekind.moneydance.model.AcctFilter;
import com.infinitekind.moneydance.model.CurrencyTable;
import com.infinitekind.moneydance.model.CurrencyType;

/**
 * Provides read/write access to an ods (OpenOffice/LibreOffice) spreadsheet
 * document.
 */
public class OdsAccessor {

	private File spreadSheetFile;
	private Account root;
	private CurrencyTable securities;
	private StringWriter msgBuffer;
	private PrintWriter msgWriter;
	private ResourceBundle msgBundle;

	private ArrayList<CellHandler> changes = new ArrayList<CellHandler>();
	private SpreadsheetDocument spreadSheetDoc = null;
	private Table firstSheet = null;
	private Row dateRow = null;
	private int latestColumn = 0;
	private int latestDate = 0;
	private int numPricesSet = 0;
	private int numBalancesSet = 0;

	private static final String baseMessageBundleName = "com.moneydance.modules.features.nwsync.NwSyncMessages";
	private static final double [] centMult = {1, 10, 100, 1000, 10000};

	/**
	 * Sole constructor.
	 *
	 * @param spreadSheetFile ods spreadsheet document
	 * @param accountBook Moneydance account book
	 */
	public OdsAccessor(File spreadSheetFile, AccountBook accountBook) {
		this.spreadSheetFile = spreadSheetFile;
		this.root = accountBook.getRootAccount();
		this.securities = accountBook.getCurrencies();
		this.msgBuffer = new StringWriter();
		this.msgWriter = new PrintWriter(this.msgBuffer);
		try {
			this.msgBundle = ResourceBundle.getBundle(baseMessageBundleName);
		} catch (Exception e) {
			System.err.format("Unable to load message bundle %s. %s%n", baseMessageBundleName, e);
			this.msgBundle = new ResourceBundle() {
				protected Object handleGetObject(String key) {
					// just use the key since we have no message bundle
					return key;
				}

				public Enumeration<String> getKeys() {
					return null;
				}
			};
		}

	} // end (File, AccountBook) constructor

	/**
	 * Synchronize data between a spreadsheet document and Moneydance.
	 */
	public void syncNwData() throws OdsException {
		Iterator<Row> rowItr = getFirstSheet().getRowIterator();
		findDateRow(rowItr);
		findLatestDate();

		while (rowItr.hasNext()) {
			Row r = rowItr.next();
			Cell key = r.getCellByIndex(0);

			if (CellHandler.isValueType(key, STRING)) {
				CellHandler val = CellHandler.getCellHandlerByIndex(r, this.latestColumn);

				if (val != null && val.getFormula() == null) {
					String keyVal = key.getStringValue();
					CurrencyType security = this.securities.getCurrencyByTickerSymbol(keyVal);

					if (security != null) {
						// found this row's ticker symbol in Moneydance securities
						Number price = val.getValue();

						if (price instanceof Double && price.doubleValue() > 0) {
							// set the new security price
							setPrice(security, price.doubleValue());
						}
					} else {
						Account account = getAccount(keyVal);

						if (account != null) {
							// found this row's account in Moneydance
							setBalanceIfDiff(val, account, keyVal);
						} else {
							System.err.println("Ignoring row " + keyVal);
						}
					}
				}
			}
		} // end while

	} // end syncNwData()

	/**
	 * @param keyVal
	 * @return the Moneydance account corresponding to keyVal
	 */
	private Account getAccount(String keyVal) {
		final String[] actNames = keyVal.split(":");
		Account account = this.root.getAccountByName(actNames[0]);

		if (account != null && actNames.length > 1) {
			List<Account> subs = account.getSubAccounts(new AcctFilter() {

				public boolean matches(Account acct) {
					String accountName = acct.getAccountName();

					return accountName.equalsIgnoreCase(actNames[1]);
				} // end matches(Account)

				public String format(Account acct) {

					return acct.getFullAccountName();
				} // end format(Account)
			});
			account = subs == null || subs.isEmpty() ? null : subs.get(0);
		}

		return account;
	} // end getAccount(String)

	/**
	 * Set the Moneydance security price for the latest date found in the
	 * spreadsheet.
	 *
	 * @param security
	 * @param price
	 */
	private void setPrice(CurrencyType security, double price) {
		// Get the old price rounded to the tenth place past the decimal point
		BigDecimal bd = BigDecimal.valueOf(1 / security.getUserRateByDateInt(this.latestDate));
		double oldPrice = bd.setScale(10, RoundingMode.HALF_EVEN).doubleValue();

		if (price != oldPrice) {
			double rate = 1 / price;
			security.setSnapshotInt(this.latestDate, rate);
			security.setUserRate(rate); // this is usually the newest snapshot
			security.syncItem();
			++this.numPricesSet;

			// Change %s price from $%.2f to $%.2f (%+.2f%%).%n
			writeFormatted("NWSYNC10", security.getName(), oldPrice, price,
					(price / oldPrice - 1) * 100);
		}

	} // end setPrice(CurrencyType, double)

	/**
	 * Set the spreadsheet account balance if it differs from Moneydance.
	 *
	 * @param val
	 * @param account
	 * @param keyVal
	 */
	private void setBalanceIfDiff(CellHandler val, Account account, String keyVal) {
		int decimalPlaces = account.getCurrencyType().getDecimalPlaces();
		double balance = account.getUserCurrentBalance() / centMult[decimalPlaces];

		if (account.getAccountType() == CREDIT_CARD) {
			balance = -balance;
		}
		Number oldBalance = val.getValue();

		if ((oldBalance instanceof Double) && balance != oldBalance.doubleValue()) {
			val.setNewValue(balance);
			this.changes.add(val);
			++this.numBalancesSet;

			// Change %s balance from %s to %s.%n
			NumberFormat nf = val.getNumberFormat();
			nf.setMaximumFractionDigits(decimalPlaces);
			writeFormatted("NWSYNC11", keyVal, nf.format(oldBalance.doubleValue()),
					nf.format(balance));
		}

	} // end setBalanceIfDiff(CellHandler, Account, String)

	/**
	 * Capture row with 'Date' in first column.
	 *
	 * @param rowIterator
	 */
	private void findDateRow(Iterator<Row> rowIterator) throws OdsException {
		while (rowIterator.hasNext()) {
			Row r = rowIterator.next();
			Cell c = r.getCellByIndex(0);

			if (CellHandler.isValueType(c, STRING)
					&& "Date".equalsIgnoreCase(c.getStringValue())) {
				this.dateRow = r;

				return;
			}
		} // end while

		// Unable to find row with 'Date' in first column in %s.
		throw new OdsException(null, "NWSYNC01", this.spreadSheetFile);
	} // end findDateRow(Iterator<Row>)

	/**
	 * Capture index of the rightmost date in the supplied row. Also capture this
	 * date's YYYYMMDD value.
	 */
	private void findLatestDate() throws OdsException {
		int cellIndex = 0;
		CellHandler c = null, latestCell;

		do {
			latestCell = c;
			c = CellHandler.getCellHandlerByIndex(this.dateRow, ++cellIndex);
		} while (c instanceof CellHandler.DateCellHandler);

		if (cellIndex == 1)
			// Unable to find any dates in the row with 'Date' in first column in %s.
			throw new OdsException(null, "NWSYNC02", this.spreadSheetFile);

		this.latestColumn = cellIndex - 1;

		// capture date value in decimal form YYYYMMDD
		this.latestDate = latestCell.getValue().intValue();

		// Found rightmost date: %s.%n
		writeFormatted("NWSYNC13", latestCell.getDisplayText());

	} // end findLatestDate()

	/**
	 * Commit any changes to the spreadsheet document.
	 *
	 * @return null
	 */
	public OdsAccessor commitChanges() throws OdsException {
		if (isModified()) {
			for (CellHandler cHandler : this.changes) {
				cHandler.applyUpdate();
			}
			SpreadsheetDocument doc = getSpreadSheetDoc();
			try {
				doc.save(this.spreadSheetFile);
			} catch (Exception e) {
				// Exception saving changes to spreadsheet document %s. %s
				throw new OdsException(e, "NWSYNC00", this.spreadSheetFile, e);
			}
			this.changes.clear();

			// write a summary line
			writeSummary();
		}
		return null;
	} // end commitChanges()

	/**
	 * Write a summary line.
	 */
	private void writeSummary() {
		String date = this.dateRow.getCellByIndex(this.latestColumn).getDisplayText();

		// For %s changed %d security price%s and %d account balance%s.%n
		writeFormatted("NWSYNC12", date, this.numPricesSet, this.numPricesSet == 1 ? "" : "s",
				this.numBalancesSet, this.numBalancesSet == 1 ? "" : "s");

	} // end writeSummary()

	/**
	 * @return True when the spreadsheet has uncommitted changes in memory.
	 */
	public boolean isModified() {

		return !this.changes.isEmpty();
	} // end isModified()

	/**
	 * @return a string representing accumulated messages
	 */
	public String getMessages() {

		return this.msgBuffer.toString();
	} // end getMessages()

	/**
	 * @return the spreadsheet document, loading it if necessary
	 */
	private SpreadsheetDocument getSpreadSheetDoc() throws OdsException {
		if (this.spreadSheetDoc == null) {
			try {
				this.spreadSheetDoc = SpreadsheetDocument.loadDocument(this.spreadSheetFile);
			} catch (Exception e) {
				// Exception loading spreadsheet document %s. %s
				throw new OdsException(e, "NWSYNC04", this.spreadSheetFile, e);
			}

			if (this.spreadSheetDoc == null)
				// Unable to load spreadsheet document %s.
				throw new OdsException(null, "NWSYNC05", this.spreadSheetFile);
		}
		return this.spreadSheetDoc;
	} // end getSpreadSheetDoc()

	/**
	 * @return the first sheet in the spreadsheet document
	 */
	private Table getFirstSheet() throws OdsException {
		if (this.firstSheet == null) {
			this.firstSheet = getSpreadSheetDoc().getSheetByIndex(0);

			if (this.firstSheet == null)
				// Unable to load first sheet in %s.
				throw new OdsException(null, "NWSYNC03", this.spreadSheetFile);
		}
		return this.firstSheet;
	} // end getFirstSheet()

	/**
	 * Inner class to house exceptions.
	 */
	public class OdsException extends Exception {

		private static final long serialVersionUID = -9019657124886615063L;

		/**
		 * @param cause Exception that caused this (null if none)
		 * @param key The resource bundle key (or message)
		 * @param params Optional parameters for the detail message
		 */
		public OdsException(Throwable cause, String key, Object... params) {
			super(String.format(retrieveMessage(key), params), cause);

		} // end (Throwable, String, Object...) constructor

	} // end class OdsException

	/**
	 * @param key The resource bundle key (or message)
	 * @return message for this key
	 */
	private String retrieveMessage(String key) {
		try {
			return this.msgBundle.getString(key);
		} catch (Exception e) {
			// just use the key when not found
			return key;
		}
	} // end retrieveMessage(String)

	/**
	 * @param key The resource bundle key (or message)
	 * @param params Optional array of parameters for the message
	 */
	private void writeFormatted(String key, Object... params) {
		this.msgWriter.format(retrieveMessage(key), params);

	} // end writeFormatted(String, Object...)

} // end class OdsAccessor
