package com.test;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;


import org.apache.commons.codec.binary.Base64;
//import java.util.Base64;
import javax.swing.*;

import java.awt.Container;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JButton;
import java.awt.event.*;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;

import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;

public class WordToHtml  {

    public static void main(String[] args) {
        new WordToHtml("word转html");        //创建窗口
    }
    /**{
     * 创建并显示GUI。出于线程安全的考虑，
     * 这个方法在事件调用线程中调用。
     */
    public  WordToHtml(String title) {
        JFrame jf = new JFrame(title);
        Container conn = jf.getContentPane();    //得到窗口的容器
        jf.setBounds(500,500,500,500); //设置窗口的属性 窗口位置以及窗口的大小
        jf.setVisible(true);//设置窗口可见
        jf.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); //设置关闭方式 如果不设置的话 似乎关闭窗口之后不会退出程序

        JFileChooser jfc=new JFileChooser();
        conn.add(jfc);
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
        jfc.showDialog(new JLabel(), "选择");
        File file=jfc.getSelectedFile();
        if(file.isDirectory()){
            System.out.println("文件夹:"+file.getAbsolutePath());
        }else if(file.isFile()){
            System.out.println("文件:"+file.getAbsolutePath());
        }


        String filestr=jfc.getSelectedFile().toString();
        System.out.println("需解析的文件"+filestr);

        String path= filestr.split(jfc.getSelectedFile().getName())[0];

        int dot = jfc.getSelectedFile().getName().lastIndexOf('.');
        String filename= jfc.getSelectedFile().getName().substring(0, dot);
        String gs=jfc.getSelectedFile().getName().substring(dot+1);

        // doc 和docx

        this.docToHtml(filestr,path,filename);
        System.out.println("解析完毕...");
        JLabel j1=new JLabel("文件已解析到:"+path+"\\"+filename+".html");
        conn.add(j1);

    }

    public void docToHtml(String filestr,String path,String filename) {
        try{
            HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream(filestr));
            WordToHtmlConverter wordToHtmlConverter = new ImageConverter(
                    DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument()
            );
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(out);
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer serializer = transformerFactory.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            out.close();
            String result = new String(out.toByteArray());
            FileUtils.writeStringToFile(new File(path, filename+".html"), result, "utf-8");
        }catch (IOException exception){
            System.out.println("需解析的文件"+filestr);

        }catch (ParserConfigurationException exception){
            System.out.println("需解析的文件"+filestr);
        }
        catch (TransformerException exception){
            System.out.println("需解析的文件"+filestr);
        }
    }



    public static class ImageConverter extends WordToHtmlConverter{

        public ImageConverter(Document document) {
            super(document);
        }
        @Override
        protected void processImageWithoutPicturesManager(Element currentBlock, boolean inlined, Picture picture){
            Element imgNode = currentBlock.getOwnerDocument().createElement("img");
            StringBuffer sb = new StringBuffer();
            sb.append(java.util.Base64.getMimeEncoder().encodeToString(picture.getRawContent()));
            sb.insert(0, "data:" + picture.getMimeType() + ";base64,");
            imgNode.setAttribute("src", sb.toString());
            currentBlock.appendChild(imgNode);
        }
    }
}