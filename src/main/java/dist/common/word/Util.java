package dist.common.word;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * Created by chenyp on 2016/1/29.
 */
public class Util {

    /*private static ActiveXComponent WORDAPP=new ActiveXComponent("Word.Application");
    static {
        WORDAPP=new ActiveXComponent("Word.Application");
        WORDAPP.setProperty("Visible", new Variant(false));
    }*/

    public static boolean WordToHtml(String sourceFile,String targetFile){
        boolean result=false;
        ComThread.InitSTA();
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        try
        {
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();
            Dispatch doc = Dispatch.invoke(docs, "Open", Dispatch.Method,new Object[] { sourceFile, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
            Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] {targetFile, new Variant(Const.WORD_TO_HTML) }, new int[1]);
            Dispatch.call(doc, "Close", new Variant(false));
            result=true;
        }
        catch (Exception e)
        {
            e.printStackTrace();
            result=false;
        }
        finally
        {
            app.invoke("Quit", new Variant[]{});
            ComThread.Release();
            ComThread.quitMainSTA();
        }
        return result;
    }

    public static boolean HTMLToWord(String sourceFile,String targetFile) {
        boolean result=false;
        ComThread.InitSTA();
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        try {
            app.setProperty("Visible", new Variant(false));
            Dispatch htmls = app.getProperty("Documents").toDispatch();
            Dispatch html = Dispatch.invoke(htmls, "Open", Dispatch.Method, new Object[] { sourceFile, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
            Dispatch.invoke(html, "SaveAs", Dispatch.Method, new Object[] { targetFile, new Variant(Const.HTML_TO_WORD) }, new int[1]);
            Variant f = new Variant(false);
            Dispatch.call(html, "Close", f);
            result=true;
        } catch (Exception e) {
            e.printStackTrace();
            result=false;
        } finally {
            app.invoke("Quit", new Variant[] {});
            ComThread.Release();
            ComThread.quitMainSTA();
        }
        return result;
    }
}
