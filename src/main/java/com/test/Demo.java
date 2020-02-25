package com.test;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.hwpf.usermodel.Picture;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.swing.*;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.awt.event.*;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Demo extends JFrame {
    private JButton selectfile;
    private JTextField filename;
    private JTextArea outmessage;
    private JButton jiexi;
    JPanel jp;
    String outmsg;
    public Demo() {

        selectfile.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                outmessage.setText("");
                JFileChooser jfc=new JFileChooser();
                jp.add(jfc);
                jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
                jfc.showDialog(new JLabel(), "选择");
                File file=jfc.getSelectedFile();
                if(file.isDirectory()){
                    System.out.println("文件夹:"+file.getAbsolutePath());
                }else if(file.isFile()){
                    System.out.println("文件:"+file.getAbsolutePath());
                }
                String filestr=jfc.getSelectedFile().toString();
                filename.setText(filestr);
            }
        });
        jiexi.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                String filestr=filename.getText();
                File file=new File(filestr);
                String path=filestr.split(file.getName())[0];
                String filename=file.getName();

                int dot = filename.lastIndexOf('.');
                outmsg+=filestr+"\n";
                outmsg+=path+"\n";
                Demo demo=new Demo();
                outmsg+=filename.substring(0, dot)+"\n";
                String result=demo.docToHtml(filestr, path, filename.substring(0, dot));
                outmsg+=result+"\n";
                outmsg += "解析完毕....";
                outmessage.setText(outmsg);
            }
        });


    }
    public void init(){
        this.setTitle("doc转换HTML");
        jp = new JPanel();
        this.add(jp);
        filename.setColumns(30);
        filename.setEnabled(false);
        jp.add(filename);
        jp.add(selectfile);
        jp.add(jiexi);
        outmessage.setColumns(50);
        outmessage.setRows(20);
        jp.add(outmessage);
        this.setSize(800, 500);
        this.setLocation(200, 300);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setVisible(true);
    }
    public String docToHtml(String filestr, String path, String filename) {
        String outmsgstr="";
        try{
            HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream(filestr));
            WordToHtmlConverter wordToHtmlConverter = new WordToHtml.ImageConverter(
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
            outmsgstr="文件格式不正确，无法解析。仅支持doc 2003/2007";
            System.out.println("需解析的文件"+filestr);

        }catch (ParserConfigurationException exception){
            outmsgstr="文件格式不正确，无法解析。仅支持doc 2003/2007";
            System.out.println("需解析的文件"+filestr);
        }
        catch (TransformerException exception){
            outmsgstr="文件格式不正确，无法解析。仅支持doc 2003/2007";
            System.out.println("需解析的文件"+filestr);
        }catch (IllegalArgumentException exception){
            outmsgstr="文件格式不正确，无法解析。仅支持doc 2003/2007";
        }
        return outmsgstr;
    }



    public class ImageConverter extends WordToHtmlConverter{

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
    public static void main(String[] args) {
       new Demo().init();
    }
}
