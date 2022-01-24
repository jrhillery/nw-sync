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
			this.odsAcc = new OdsAccessor(this.messageWindow.getLocale(),
				getContext().getCurrentAccountBook());
		}
		try {
			synchronized (this) {
				this.messageWindow.clearText();

				// SwingWorker instances are not reusable, so make a new one
				NwSyncWorker worker = new NwSyncWorker(this.messageWindow, this.odsAcc);
				worker.execute();
			}
		} catch (Throwable e) {
			handleException(e);
		}

	} // end invoke(String)

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

	} // end closeConsole()

} // end class Main
