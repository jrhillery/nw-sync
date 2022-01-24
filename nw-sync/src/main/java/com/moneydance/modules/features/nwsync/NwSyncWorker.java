package com.moneydance.modules.features.nwsync;

import javax.swing.SwingWorker;
import java.util.List;

public class NwSyncWorker extends SwingWorker<Boolean, String> {
   private final MessageWindow syncWindow;
   private final OdsAccessor odsAcc;

   public NwSyncWorker(MessageWindow syncWindow, OdsAccessor odsAcc) {
      super();
      this.syncWindow = syncWindow;
      this.odsAcc = odsAcc;
   } // end constructor

   public void registerClosableResource(AutoCloseable closable) {
      this.syncWindow.setCloseableResource(closable);
   } // end registerClosableResource(AutoCloseable)

   /**
    * Long-running routine to synchronize Moneydance with a spreadsheet.
    * Runs on worker thread.
    *
    * @return true when changes have been detected
    */
   protected Boolean doInBackground() {
      try {
         this.odsAcc.forgetChanges();
         this.odsAcc.syncNwData(this);

         return this.odsAcc.isModified();
      } catch (Throwable e) {
         display(e.toString());
         e.printStackTrace(System.err);

         return false;
      }
   } // end doInBackground()

   /**
    * Enable the commit button if we have changes.
    * Runs on event dispatch thread after the doInBackground method is finished.
    */
   protected void done() {
      try {
         this.syncWindow.enableCommitButton(get());
      } catch (Exception e) {
         this.syncWindow.addText(e.toString());
         e.printStackTrace(System.err);
      }
   } // end done()

   /**
    * Runs on worker thread.
    *
    * @param msgs Messages to display
    */
   public void display(String... msgs) {
      publish(msgs);
   } // end display(String...)

   /**
    * Runs on event dispatch thread.
    *
    * @param chunks Messages to process
    */
   protected void process(List<String> chunks) {
      for (String msg: chunks) {
         this.syncWindow.addText(msg);
      }
   } // end process(List<String>)

} // end class NwSyncWorker
