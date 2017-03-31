 在线浏览 office 文档一直是个刚需，除了利用[Office Web Apps](http://www.chenyp.com/2016/03/15/office-online/) 之外，我也探索了另外一种实现：将 office 文档转成 html，然后嵌入到可编辑的 iframe 中，同时解决浏览和编辑的问题，不过编辑特殊的东西肯定是不行的，比如图表之类的。

将 office 转成 html 用到了 jacob，在 Maven 中央仓库中的地址如下：
> [http://search.maven.org/#search%7Cga%7C1%7Cg%3A%22net.sf.jacob-project%22](http://note.youdao.com/)

### 先上demo图，选择文档:
![这里写图片描述](http://img.blog.csdn.net/20160826193807016)

### 点击上传，程序会自动将word转换成 html，并加载到 iframe 中，如下：
![这里写图片描述](http://img.blog.csdn.net/20160826193819220)

### 在转换的过程中，会把图表转换成图片，由于iframe是可编辑的，我们可以对其进行修改并保存，保存后会直接下载下来，如下所示：
![这里写图片描述](http://img.blog.csdn.net/20160826193918354)

### 打开保存下来的文档，可以看到内容已经加进来了，而且之前变成图片的图表也重新还原成了图表类型，仍可以继续编辑：
![这里写图片描述](http://img.blog.csdn.net/20160826193948870)

由于我们可以把 iframe 制作成一个富文本编辑器，所以可以提供更多的编辑功能，如格式化等，或者直接借助已有的富文本编辑器来进行改造。我之前也写过一个简易的富文本编辑器来满足项目的要求，这里就不说 js 制作富文本编辑器的过程和细节了。

这里把一些相应的代码贴上来:
### java：word 转为 html：

```
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
```
### java：html 转回 word:

```
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
```
同样的道理，还可以实现excel、ppt 与 html 的互转。
### javascript：上传后加载文档
```
function display(fileName){
    var iframe=document.getElementById("WordContainer");
    iframe.src="word/HtmlWord/"+fileName;
    iframe.onload = iframe.onreadystatechange = function() {
        if (this.readyState && this.readyState != 'complete') {
            return;
        }
        else {
            var iframeDocument = this.contentDocument || this.contentWindow.document;
            iframeDocument.body.contentEditable=true;
        }
    }
    $("#htmlid").val(fileName);
}
```
### javascript：编辑文档并保存
```
function saveWord(){
    var iframe=document.getElementById("WordContainer");
    var iframeDocument = iframe.contentDocument || iframe.contentWindow.document;
    console.log(iframeDocument.body.innerHTML);
    var htmlid=$("#htmlid")[0].value;
    var htmlcontent=iframeDocument.body.innerHTML;
    $.post("save",
            {
                htmlid:htmlid,
                htmlcontent:htmlcontent
            },
            function(data,status){
                console.log(data);
            }
    );
}
```

我的目的是想开发一款类似“百度 Doc” [http://word.baidu.com/](http://word.baidu.com/) 的东西，这样可以丰富、稳固自己的知识，目前正在抽时间筹备中(苦命的打工仔，天天加班)，欢迎各位拍砖、指导，谢谢。

![这里写图片描述](http://img.blog.csdn.net/20160826194616404)
