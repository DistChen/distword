import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import dist.common.word.Util;
import org.junit.Test;

/**
 * Created by chenyp on 2016/1/29.
 */
public class OfficeToHtml {
    public static final int WORD_HTML = 8;
    public static final int PPT_HTML = 12;
    public static final int EXCEL_HTML = 44;
    public static final int HTML_TO_WORD = 1;

    String wordFile="D:/demo/demo.doc";
    String wordHtmlFile="D:/demo/demo.html";

    String excelFile="D:/demo/demo.xlsx";
    String excelHtmlFile="D:/demo/excel.html";

    String pptFile="D:/demo/demo.ppt";
    String pptHtmlFile="D:/demo/ppt.html";

    @Test
    public void WordToHtml(){
        Util.WordToHtml(wordFile,wordHtmlFile);
    }


    @Test
    public void htmlToWord() {
        Util.HTMLToWord(wordHtmlFile,"d:/demo/htmlword.doc");
        //Util.ColseWordAPP();
    }



    @Test
    public void excelToHtml() {
        ActiveXComponent app = new ActiveXComponent("Excel.Application"); // 启动 Excel
        try {
            app.setProperty("Visible", new Variant(false));
            Dispatch excels = app.getProperty("Workbooks").toDispatch();
            Dispatch excel = Dispatch.invoke(excels, "Open", Dispatch.Method, new Object[]{excelFile, new Variant(false), new Variant(true)}, new int[1]).toDispatch();
            Dispatch.invoke(excel, "SaveAs", Dispatch.Method, new Object[]{excelHtmlFile, new Variant(EXCEL_HTML)}, new int[1]);
            Variant f = new Variant(false);
            Dispatch.call(excel, "Close", f);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            app.invoke("Quit", new Variant[]{});
        }
    }

    @Test
    public void pptToHtml() {
        ComThread.InitSTA();
        ActiveXComponent app = new ActiveXComponent("PowerPoint.Application");
        try {
            //app.setProperty("Visible", new Variant(false));
            Dispatch ppts = app.getProperty("Presentations").toDispatch();
            Dispatch ppt = Dispatch.invoke(ppts, "Open", Dispatch.Method, new Object[]{pptFile, new Variant(false), new Variant(true)}, new int[1]).toDispatch();
            Dispatch.invoke(ppt, "SaveAs", Dispatch.Method, new Object[]{pptHtmlFile, new Variant(PPT_HTML)}, new int[1]);
            Variant f = new Variant(false);
            Dispatch.call(ppt, "Close", f);
        } catch (Exception exception) {
            exception.printStackTrace();
        } finally {
            app.invoke("Quit", new Variant[0]);
            ComThread.Release();
            ComThread.quitMainSTA();
        }
    }
}
