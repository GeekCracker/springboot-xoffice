package com.hg.xs;

import java.util.logging.Level;
import java.util.logging.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * WordApp
 * 
 * @author wanghg
 */
public class WordApp {
	private static Logger logger = Logger.getLogger(WordApp.class.getName());

	public static void toHtml(String src, String tar) {
		ComThread.InitSTA();
		ActiveXComponent activexcomponent = new ActiveXComponent("Word.Application");
		try {
			activexcomponent.setProperty("Visible", new Variant(false));
			Dispatch dispatch = activexcomponent.getProperty("Documents").toDispatch();
			Dispatch dispatch1 = Dispatch.invoke(dispatch, "Open", 1,
					new Object[] { src, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
			Dispatch.invoke(dispatch1, "SaveAs", 1, new Object[] { tar, new Variant(8) }, new int[1]);
			Variant variant = new Variant(false);
			Dispatch.call(dispatch1, "Close", variant);
		} catch (Exception exception) {
			exception.printStackTrace();
		} finally {
			activexcomponent.invoke("Quit", new Variant[0]);
			ComThread.Release();
			ComThread.quitMainSTA();
		}
	}

	public static void toPdf(String src, String tar) {
		ComThread.InitSTA();
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", new Variant(false));
			app.setProperty("AutomationSecurity", new Variant(3));
			Dispatch docs = app.getProperty("Documents").toDispatch();
			doc = Dispatch.call(docs, "Open", new Object[] { src, Boolean.FALSE, Boolean.TRUE }).toDispatch();
			Dispatch.call(doc, "SaveAs", new Object[] { tar, Integer.valueOf(17) });
		} catch (Throwable t) {
			logger.log(Level.SEVERE, t.getMessage(), t);
		} finally {
			if (doc != null) {
				try {
					Dispatch.call(doc, "Close", new Object[] { Boolean.FALSE });
				} catch (Throwable t) {
					logger.log(Level.SEVERE, t.getMessage(), t);
				}
			}
			if (app != null) {
				try {
					app.invoke("Quit", new Variant[] { new Variant(false) });
					app.safeRelease();
					ComThread.Release();
				} catch (Throwable t) {
					logger.log(Level.SEVERE, t.getMessage(), t);
				}
			}

		}
	}
}
