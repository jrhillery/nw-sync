/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import static com.infinitekind.moneydance.model.Account.AccountType.CREDIT_CARD;
import static com.sun.star.table.CellContentType.FORMULA;
import static com.sun.star.table.CellContentType.TEXT;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.ResourceBundle;

import com.infinitekind.moneydance.model.Account;
import com.infinitekind.moneydance.model.AccountBook;
import com.infinitekind.moneydance.model.AcctFilter;
import com.infinitekind.moneydance.model.CurrencyTable;
import com.infinitekind.moneydance.model.CurrencyType;
import com.sun.star.container.XEnumeration;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;

/**
 * Provides read/write access to an ods (OpenOffice/LibreOffice) spreadsheet
 * document.
 */
public class OdsAccessor {
	private Account root;
	private CurrencyTable securities;
	private StringWriter msgBuffer;
	private PrintWriter msgWriter;

	private List<CellHandler> changes = new ArrayList<>();
	private CalcDoc calcDoc = null;
	private XCellRange dateRow = null;
	private int latestColumn = 0;
	private int latestDate = 0;
	private int numPricesSet = 0;
	private int numBalancesSet = 0;

	private static ResourceBundle msgBundle = null;

	private static final String baseMessageBundleName = "com.moneydance.modules.features.nwsync.NwSyncMessages";
	private static final double [] centMult = {1, 10, 100, 1000, 10000};

	/**
	 * Sole constructor.
	 *
	 * @param accountBook Moneydance account book
	 */
	public OdsAccessor(AccountBook accountBook) {
		this.root = accountBook.getRootAccount();
		this.securities = accountBook.getCurrencies();
		this.msgBuffer = new StringWriter();
		this.msgWriter = new PrintWriter(this.msgBuffer);

	} // end (AccountBook) constructor

	/**
	 * Synchronize data between a spreadsheet document and Moneydance.
	 */
	public void syncNwData() throws OdsException {
		CalcDoc calcDoc = getCalcDoc();
		if (calcDoc == null)
			return; // nothing to synchronize

		XEnumeration rowItr = calcDoc.getFirstSheetRowIterator();

		if (!findDateRow(rowItr))
			return; // can't synchronize without a date row

		if (!findLatestDate())
			return; // can't synchronize without knowing the latest date

		while (rowItr.hasMoreElements()) {
			XCellRange row = CalcDoc.next(XCellRange.class, rowItr);
			XCell key = CalcDoc.getCellByIndex(row, 0);

			if (CalcDoc.isValueType(key, TEXT) || CalcDoc.isValueType(key, FORMULA)) {
				String keyVal = CellHandler.asDisplayText(key);
				CellHandler val = calcDoc.getCellHandlerByIndex(row, this.latestColumn);

				if (val != null) {
					CurrencyType security = this.securities.getCurrencyByTickerSymbol(keyVal);

					if (security != null) {
						// found this row's ticker symbol in Moneydance securities
						// set the new security price
						setPrice(val, security);
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
	 * Set the spreadsheet security price for the latest date found in the
	 * spreadsheet.
	 *
	 * @param val
	 * @param security
	 */
	private void setPrice(CellHandler val, CurrencyType security) {
		// Get the price rounded to the tenth place past the decimal point
		BigDecimal bd = BigDecimal.valueOf(1 / security.getUserRateByDateInt(this.latestDate));
		double price = bd.setScale(10, RoundingMode.HALF_EVEN).doubleValue();

		Number oldPrice = val.getValue();

		if ((oldPrice instanceof Double) && price != oldPrice.doubleValue()) {
			// Change %s price from %s to %s (%+.2f%%).%n
			writeFormatted("NWSYNC10", security.getName(), val.getDisplayText(),
				val.getNumberFormat().format(price),
				(price / oldPrice.doubleValue() - 1) * 100);

			val.setNewValue(price);
			this.changes.add(val);
			++this.numPricesSet;
		}

	} // end setPrice(CellHandler, CurrencyType)

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
			// Change %s balance from %s to %s.%n
			writeFormatted("NWSYNC11", keyVal, val.getDisplayText(),
				val.getNumberFormat().format(balance));

			val.setNewValue(balance);
			this.changes.add(val);
			++this.numBalancesSet;
		}

	} // end setBalanceIfDiff(CellHandler, Account, String)

	/**
	 * Capture row with 'Date' in first column.
	 *
	 * @param rowIterator
	 * @return true when found
	 */
	private boolean findDateRow(XEnumeration rowIterator) throws OdsException {
		while (rowIterator.hasMoreElements()) {
			XCellRange row = CalcDoc.next(XCellRange.class, rowIterator);
			XCell c = CalcDoc.getCellByIndex(row, 0);

			if (CalcDoc.isValueType(c, TEXT)
					&& "Date".equalsIgnoreCase(CellHandler.asDisplayText(c))) {
				this.dateRow = row;

				return true;
			}
		} // end while

		// Unable to find row with 'Date' in first column in %s.%n
		writeFormatted("NWSYNC01", getCalcDoc());

		return false;
	} // end findDateRow(XEnumeration)

	/**
	 * Capture index of the rightmost date in the supplied row. Also capture this
	 * date's YYYYMMDD value.
	 *
	 * @return true when found
	 */
	private boolean findLatestDate() throws OdsException {
		int cellIndex = 0;
		CellHandler c = null, latestCell;

		do {
			latestCell = c;
			c = getCalcDoc().getCellHandlerByIndex(this.dateRow, ++cellIndex);
		} while (c instanceof CellHandler.DateCellHandler);

		if (cellIndex == 1) {
			// Unable to find any dates in the row with 'Date' in first column in %s.%n
			writeFormatted("NWSYNC02", getCalcDoc());

			return false;
		}

		this.latestColumn = cellIndex - 1;

		// capture date value in decimal form YYYYMMDD
		this.latestDate = latestCell.getValue().intValue();

		// Found date [%s] in %s.%n
		writeFormatted("NWSYNC13", latestCell.getDisplayText(), getCalcDoc());

		return true;
	} // end findLatestDate()

	/**
	 * Commit any changes to the spreadsheet document.
	 */
	public void commitChanges() {
		if (isModified()) {
			for (CellHandler cHandler : this.changes) {
				cHandler.applyUpdate();
			}
			this.changes.clear();

			// write a summary line
			writeSummary();
		}

	} // end commitChanges()

	/**
	 * Write a summary line.
	 */
	private void writeSummary() {
		String date;
		try {
			date = CellHandler
					.asDisplayText(CalcDoc.getCellByIndex(this.dateRow, this.latestColumn));
		} catch (Exception e) {
			date = "unknown date";
		}

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
	 * @return the currently open spreadsheet document
	 */
	private CalcDoc getCalcDoc() throws OdsException {
		if (this.calcDoc == null) {
			List<XSpreadsheetDocument> docList = CalcDoc.getSpreadsheetDocs();

			switch (docList.size()) {
			case 0:
				// No open spreadsheet documents found.%n
				writeFormatted("NWSYNC05");
				break;

			case 1:
				// found one => use it
				this.calcDoc = new CalcDoc(docList.get(0));
				break;

			default:
				// Found %d open spreadsheet documents. Can only work with one.%n
				writeFormatted("NWSYNC04", docList.size());
				break;
			}
		}

		return this.calcDoc;
	} // end getCalcDoc()

	/**
	 * Release any resources we acquired. This includes closing the connection to
	 * the office process.
	 *
	 * @return null
	 */
	public OdsAccessor releaseResources() {
		CalcDoc.closeOfficeConnection();

		return null;
	} // end releaseResources()

	private static ResourceBundle getMsgBundle() {
		if (msgBundle == null) {
			try {
				msgBundle = ResourceBundle.getBundle(baseMessageBundleName);
			} catch (Exception e) {
				System.err.format("Unable to load message bundle %s. %s%n", baseMessageBundleName, e);
				msgBundle = new ResourceBundle() {
					protected Object handleGetObject(String key) {
						// just use the key since we have no message bundle
						return key;
					}

					public Enumeration<String> getKeys() {
						return null;
					}
				};
			} // end catch
		}

		return msgBundle;
	} // end getMsgBundle()

	/**
	 * Inner class to house exceptions.
	 */
	public static class OdsException extends Exception {

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
	private static String retrieveMessage(String key) {
		try {
			return getMsgBundle().getString(key);
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
