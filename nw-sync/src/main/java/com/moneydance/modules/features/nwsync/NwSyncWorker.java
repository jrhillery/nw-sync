package com.moneydance.modules.features.nwsync;

import com.moneydance.apps.md.controller.FeatureModuleContext;

import javax.swing.SwingWorker;
import java.util.List;
import java.util.concurrent.CancellationException;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutionException;

import static javax.swing.SwingWorker.StateValue.DONE;

public class NwSyncWorker extends SwingWorker<Boolean, String> implements AutoCloseable {
   private final NwSyncConsole syncConsole;
   private final String extensionName;
   private final OdsAccessor odsAcc;
   private final CountDownLatch finishedLatch = new CountDownLatch(1);

   /**
    * Sole constructor.
    *
    * @param syncConsole   Our NW sync console
    * @param extensionName This extension's name
    * @param fmContext     Moneydance context
    */
   public NwSyncWorker(NwSyncConsole syncConsole, String extensionName,
                       FeatureModuleContext fmContext) {
      super();
      this.syncConsole = syncConsole;
      this.extensionName = extensionName;
      this.odsAcc = new OdsAccessor(this,
         syncConsole.getLocale(), fmContext.getCurrentAccountBook());
      syncConsole.setStaged(this.odsAcc);
      syncConsole.addCloseableResource(this);
   } // end constructor

   /**
    * Long-running routine to synchronize Moneydance with a spreadsheet.
    * Runs on worker thread.
    *
    * @return true when changes have been detected
    */
   protected Boolean doInBackground() {
      try {
         this.odsAcc.syncNwData();

         return this.odsAcc.isModified();
      } catch (Throwable e) {
         display(e.toString());
         e.printStackTrace(System.err);

         return false;
      } finally {
         this.finishedLatch.countDown();
      }
   } // end doInBackground()

   /**
    * Enable the commit button if we have changes.
    * Runs on event dispatch thread after the doInBackground method is finished.
    */
   protected void done() {
      try {
         this.syncConsole.enableCommitButton(get());
      } catch (CancellationException e) {
         // ignore
      } catch (Exception e) {
         this.syncConsole.addText(e.toString());
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
      if (!isCancelled()) {
         for (String msg: chunks) {
            this.syncConsole.addText(msg);
         }
      }
   } // end process(List<String>)

   /**
    * Stop a running execution.
    *
    * @return null
    */
   public NwSyncWorker stopExecute() {
      close();

      // we no longer need closing
      this.syncConsole.removeCloseableResource(this);

      return null;
   } // end stopExecute(String)

   /**
    * Closes this resource, relinquishing any underlying resources.
    * Cancel this worker, wait for it to complete, discard its results and close odsAcc.
    */
   public void close() {
      try (this.odsAcc) { // make sure we close odsAcc
         if (getState() != DONE) {
            System.err.format(this.syncConsole.getLocale(),
               "Cancelling running %s invocation.%n", this.extensionName);
            cancel(false);

            // wait for prior worker to complete
            try {
               this.finishedLatch.await();
            } catch (InterruptedException e) {
               // ignore
            }

            // discard results and some exceptions
            try {
               get();
            } catch (CancellationException | InterruptedException | ExecutionException e) {
               // ignore
            }
         }
      } // end try-with-resources
   } // end close()

} // end class NwSyncWorker
