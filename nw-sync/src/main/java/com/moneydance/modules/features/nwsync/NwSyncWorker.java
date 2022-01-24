package com.moneydance.modules.features.nwsync;

import com.infinitekind.moneydance.model.AccountBook;
import com.moneydance.apps.md.controller.FeatureModuleContext;

import javax.swing.SwingWorker;
import java.util.List;

public class NwSyncWorker extends SwingWorker<Boolean, String> {
   private final MessageWindow syncWindow;
   private final AccountBook accountBook;

   public NwSyncWorker(MessageWindow syncWindow, FeatureModuleContext fmContext) {
      super();
      this.syncWindow = syncWindow;
      this.accountBook = fmContext.getCurrentAccountBook();
   } // end constructor

   /**
    * Long-running routine to synchronize Moneydance with a spreadsheet.
    * Runs on worker thread.
    *
    * @return true when changes have been detected
    */
   protected Boolean doInBackground() {
      try {
         OdsAccessor odsAcc = new OdsAccessor(this,
               this.syncWindow.getLocale(), this.accountBook);
         this.syncWindow.setStaged(odsAcc);
         this.syncWindow.setCloseableResource(odsAcc);
         odsAcc.syncNwData();

         return odsAcc.isModified();
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
