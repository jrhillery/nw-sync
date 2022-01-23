/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import com.moneydance.apps.md.controller.FeatureModule;

/**
 * Module used to synchronize John's NW spreadsheet document with Moneydance.
 */
public class Main extends FeatureModule {
	private MessageWindow messageWindow = null;
	private OdsAccessor odsAcc = null;
	private final Object synchObj = new Object();

	/**
	 * Register this module to be invoked via the Extensions menu.
	 *
	 * @see com.moneydance.apps.md.controller.FeatureModule#init()
	 */
	public void init() {
		getContext().registerFeature(this, "do:nw:sync", null, getName());

	} // end init()

	/**
	 * This is called when this extension is invoked.
	 *
	 * @see com.moneydance.apps.md.controller.FeatureModule#invoke(java.lang.String)
	 */
	public void invoke(String uri) {
		System.err.println(getName() + " invoked with uri [" + uri + ']');
		showConsole();

		if (this.odsAcc == null) {
			this.odsAcc = new OdsAccessor(this.messageWindow,
				getContext().getCurrentAccountBook());
		}
		try {
			synchronized (this.synchObj) {
				this.messageWindow.clearText();
				this.odsAcc.forgetChanges();
				this.odsAcc.syncNwData();
			}
			this.messageWindow.enableCommitButton(this.odsAcc.isModified());
		} catch (Throwable e) {
			handleException(e);
		}

	} // end invoke(String)

	/**
	 * This is called when the commit button is selected.
	 */
	void commitChanges() {
		try {
			synchronized (this.synchObj) {
				this.odsAcc.commitChanges();
			}
			this.messageWindow.enableCommitButton(this.odsAcc.isModified());
		} catch (Throwable e) {
			handleException(e);
		}

	} // end commitChanges()

	private void handleException(Throwable e) {
		this.messageWindow.addText(e.toString());
		this.messageWindow.enableCommitButton(false);
		e.printStackTrace(System.err);

	} // end handleException(Throwable)

	public void cleanup() {
		closeConsole();

	} // end cleanup()

	public String getName() {

		return "NW Sync";
	} // end getName()

	/**
	 * Show our console window.
	 */
	private synchronized void showConsole() {
		if (this.messageWindow == null) {
			this.messageWindow = new MessageWindow(this);
			this.messageWindow.setVisible(true);
		} else {
			this.messageWindow.setVisible(true);
			this.messageWindow.toFront();
			this.messageWindow.requestFocus();
		}

	} // end showConsole()

	/**
	 * Close our console window and release resources.
	 */
	synchronized void closeConsole() {
		if (this.messageWindow != null)
			this.messageWindow = this.messageWindow.goAway();

		if (this.odsAcc != null)
			this.odsAcc = this.odsAcc.releaseResources();

	} // end closeConsole()

} // end class Main
