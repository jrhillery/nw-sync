/*
 * Created on Jan 20, 2018
 */
package com.moneydance.modules.features.nwsync;

/**
 * Message bundle provider interface.
 */
public interface MessageBundleProvider {

	/**
	 * @param key The resource bundle key (or message)
	 * @return message for this key
	 */
	String retrieveMessage(String key);

} // end interface MessageBundleProvider
