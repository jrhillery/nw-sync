/*
 * Created on Nov 10, 2017
 */
package com.moneydance.modules.features.nwsync;

import com.moneydance.apps.md.controller.FeatureModule;

/**
 * Module used to synchronize John's NW spreadsheet document with Moneydance.
 */
public class Main extends FeatureModule {
	private NwSyncConsole syncConsole = null;
	private NwSyncWorker syncWorker = null;

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
		System.err.format("%s invoked with uri [%s].%n", getName(), uri);
		showConsole();

		try {
			if (this.syncWorker != null) {
				this.syncWorker.stopExecute(getName());
			}
			this.syncConsole.clearText();

			// SwingWorker instances are not reusable, so make a new one
			this.syncWorker = new NwSyncWorker(this.syncConsole, getContext());
			this.syncWorker.execute();
		} catch (Throwable e) {
			handleException(e);
		}

	} // end invoke(String)

	private void handleException(Throwable e) {
		this.syncConsole.addText(e.toString());
		this.syncConsole.enableCommitButton(false);
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
		if (this.syncConsole == null) {
			this.syncConsole = new NwSyncConsole(this);
			this.syncConsole.setVisible(true);
		} else {
			this.syncConsole.setVisible(true);
			this.syncConsole.toFront();
			this.syncConsole.requestFocus();
		}

	} // end showConsole()

	/**
	 * Close our console window and release resources.
	 */
	synchronized void closeConsole() {
		if (this.syncConsole != null)
			this.syncConsole = this.syncConsole.goAway();

	} // end closeConsole()

} // end class Main
