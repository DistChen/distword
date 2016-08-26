package dist.common.word;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;


import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;



public class SaveWordServlet extends HttpServlet {

    private static final long serialVersionUID = 1L;
   
    protected void doPost(HttpServletRequest request,HttpServletResponse response) throws ServletException, IOException {
        request.setCharacterEncoding("UTF-8");
        String fileName=request.getParameter("htmlid");
        String htmlcontent=request.getParameter("htmlcontent");
        String htmlPath = request.getSession().getServletContext().getRealPath("/")+"word\\HtmlWord";
        File input = new File(htmlPath+"\\"+fileName);
        Document doc = Jsoup.parse(input, "UTF-8");
        Element body = doc.body();
        body.html(htmlcontent);
        //System.out.println(body.html());
        FileOutputStream fos = new FileOutputStream(htmlPath+"\\"+fileName, false);
        fos.write(doc.html().getBytes());
        fos.close();

        Util.HTMLToWord(htmlPath+"\\"+fileName,htmlPath+"\\word.doc");


    }
    
    protected void doGet(HttpServletRequest req, HttpServletResponse resp)
            throws ServletException, IOException {
        doPost(req, resp);
    }

}