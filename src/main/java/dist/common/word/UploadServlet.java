package dist.common.word;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.*;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;


public class UploadServlet extends HttpServlet {

    private static final long serialVersionUID = 1L;
    private int maxPostSize = 100 * 1024 * 1024; 
   
    protected void doPost(HttpServletRequest request,HttpServletResponse response) throws ServletException, IOException {

      DiskFileItemFactory factory = new DiskFileItemFactory();
      factory.setSizeThreshold(4096);  
      ServletFileUpload fileUpload = new ServletFileUpload(factory);
      fileUpload.setSizeMax(maxPostSize);
      //文件上传地址
      String wordPath = request.getSession().getServletContext().getRealPath("/")+"word\\WordSource";
      String htmlPath = request.getSession().getServletContext().getRealPath("/")+"word\\HtmlWord";
      String newFileName = "";
      String srcFileName = "";
      PrintWriter out = response.getWriter();
		try {
			File file = new File(wordPath);
			if (!file.exists()) {
				file.mkdirs();
			}
 			request.setCharacterEncoding("utf-8");
			List items = fileUpload.parseRequest(request);
			Iterator iter = items.iterator();
			while (iter.hasNext()) {
                String guid=UUID.randomUUID().toString();
				FileItem item = (FileItem) iter.next();
				if (!item.isFormField()) {
					srcFileName = item.getName();
					newFileName = guid+"_"+srcFileName;
					File uploadedFile = new File(wordPath + "/" + newFileName);
					item.write(uploadedFile);
                    if(Util.WordToHtml(wordPath + "/" + newFileName,htmlPath+"/"+guid+".html")){
                        out.write(guid+".html");
                    }
				}
			}
        } catch (FileUploadException e) {
            e.printStackTrace();
            out.write("");
        } catch (Exception e) {
            e.printStackTrace();
            out.write("");
        }
    }
    
    protected void doGet(HttpServletRequest req, HttpServletResponse resp)
            throws ServletException, IOException {
        doPost(req, resp);
    }

}