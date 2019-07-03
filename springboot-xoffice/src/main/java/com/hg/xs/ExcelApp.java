package com.hg.xs;

import java.util.logging.Level;
import java.util.logging.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * ExcelApp
 * 
 * @author wanghg
 */
public class ExcelApp {
	private static Logger logger = Logger.getLogger(WordApp.class.getName());

	public static void toHtml(String src, String tar) {
		ComThread.InitSTA();
		// 启动excel
		ActiveXComponent app = new ActiveXComponent("Excel.Application");
		try {
			// 设置excel不可见
			app.setProperty("Visible", new Variant(false));
			Dispatch excels = app.getProperty("Workbooks").toDispatch();
			// 打开excel文件
			Dispatch excel = Dispatch.invoke(excels, "Open", Dispatch.Method,
					new Object[] { src, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
			// 作为html格式保存到临时文件
			Dispatch.invoke(excel, "SaveAs", Dispatch.Method, new Object[] { tar, new Variant(44) }, new int[1]);
			Variant f = new Variant(false);
			Dispatch.call(excel, "Close", f);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			app.invoke("Quit", new Variant[] {});
			app.safeRelease();
			ComThread.Release();
		}
	}
	public static void toPdf(String src, String tar) {
		ComThread.InitSTA();
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			app = new ActiveXComponent("Excel.Application");
			app.setProperty("Visible", new Variant(false));
			app.setProperty("AutomationSecurity", new Variant(3));
			Dispatch docs = app.getProperty("Workbooks").toDispatch();
			doc = Dispatch.call(docs, "Open", new Object[] { src, Boolean.FALSE, Boolean.TRUE }).toDispatch();
			Dispatch.call(doc, "ExportAsFixedFormat", new Object[] { Integer.valueOf(0), tar });
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
					app.invoke("Quit");
					app.safeRelease();
					ComThread.Release();
				} catch (Throwable t) {
					logger.log(Level.SEVERE, t.getMessage(), t);
				}
			}
		}
	}
}
