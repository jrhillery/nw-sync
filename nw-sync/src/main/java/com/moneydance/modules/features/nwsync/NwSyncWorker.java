package com.moneydance.modules.features.nwsync;

import com.moneydance.apps.md.controller.FeatureModuleContext;

import javax.swing.SwingWorker;
import java.util.List;
import java.util.concurrent.CancellationException;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutionException;

import static javax.swing.SwingWorker.StateValue.DONE;

public class NwSyncWorker extends SwingWorker<Boolean, String> {
   private final NwSyncConsole syncConsole;
   private final OdsAccessor odsAcc;
   private final CountDownLatch finishedLatch = new CountDownLatch(1);

   /**
    * Sole constructor.
    *
    * @param syncConsole Our NW sync console
    * @param fmContext  Moneydance context
    */
   public NwSyncWorker(NwSyncConsole syncConsole, FeatureModuleContext fmContext) {
      super();
      this.syncConsole = syncConsole;
      this.odsAcc = new OdsAccessor(this,
         syncConsole.getLocale(), fmContext.getCurrentAccountBook());
   } // end constructor

   /**
    * Long-running routine to synchronize Moneydance with a spreadsheet.
    * Runs on worker thread.
    *
    * @return true when changes have been detected
    */
   protected Boolean doInBackground() {
      try {
         this.syncConsole.setStaged(this.odsAcc);
         this.syncConsole.setCloseableResource(this.odsAcc);
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
    * @param name This extension's name
    */
   public void stopExecute(String name) {
      if (getState() != DONE) {
         System.err.format(this.syncConsole.getLocale(),
            "Cancelling prior %s invocation.%n", name);
         cancel(false);

         // wait for prior worker to complete
         try {
            this.finishedLatch.await();
         } catch (InterruptedException e) {
            // ignore
         }

         // discard results and select exceptions
         try {
            get();
         } catch (CancellationException | InterruptedException | ExecutionException e) {
            // ignore
         }
      }
      this.odsAcc.close();
   } // end stopExecute(String)

} // end class NwSyncWorker
