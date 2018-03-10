package excelOperate;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;  
import org.apache.poi.xwpf.usermodel.*;  
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;  
  
import java.io.File;  
import java.io.FileOutputStream;  
import java.math.BigInteger;  
  
  
/** 
 * Created by zhouhs on 2017/1/9. 
 */  
public class WordExportController {  
  
    public static void main(String[] args)throws Exception {  
        //Blank Document  
        XWPFDocument document= new XWPFDocument();  
  
        //Write the Document in file system  
        FileOutputStream out = new FileOutputStream(new File("create_table.docx"));  
  
  
        //��ӱ���  
        XWPFParagraph titleParagraph = document.createParagraph();  
        //���ö������  
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);  
  
        XWPFRun titleParagraphRun = titleParagraph.createRun();  
        titleParagraphRun.setText("Java PoI");  
        titleParagraphRun.setColor("000000");  
        titleParagraphRun.setFontSize(20);  
  
  
        //����  
        XWPFParagraph firstParagraph = document.createParagraph();  
        XWPFRun run = firstParagraph.createRun();  
        run.setText("Java POI ����word�ļ���");  
        run.setColor("696969");  
        run.setFontSize(16);  
  
        //���ö��䱳����ɫ  
        CTShd cTShd = run.getCTR().addNewRPr().addNewShd();  
        cTShd.setVal(STShd.CLEAR);
 //       cTShd.setFill("97FFFF");  
  
  
        //����  
        XWPFParagraph paragraph1 = document.createParagraph();  
        XWPFRun paragraphRun1 = paragraph1.createRun();  
        paragraphRun1.setText("\r");  
  
  
        //������Ϣ���  
        XWPFTable infoTable = document.createTable();  
        //ȥ���߿�  
        infoTable.getCTTbl().getTblPr().unsetTblBorders();  
  
  
        //�п��Զ��ָ�  
        CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();  
        infoTableWidth.setType(STTblWidth.DXA);  
        infoTableWidth.setW(BigInteger.valueOf(9072));  
  
  
        //����һ��  
        XWPFTableRow infoTableRowOne = infoTable.getRow(0);  
        infoTableRowOne.getCell(0).setText("ְλ");  
        infoTableRowOne.addNewTableCell().setText(": Java ��������ʦ");  
  
        //���ڶ���  
        XWPFTableRow infoTableRowTwo = infoTable.createRow();  
        infoTableRowTwo.getCell(0).setText("����");  
        infoTableRowTwo.getCell(1).setText(": seawater");  
  
        //��������  
        XWPFTableRow infoTableRowThree = infoTable.createRow();  
        infoTableRowThree.getCell(0).setText("����");  
        infoTableRowThree.getCell(1).setText(": xxx-xx-xx");  
  
        //��������  
        XWPFTableRow infoTableRowFour = infoTable.createRow();  
        infoTableRowFour.getCell(0).setText("�Ա�");  
        infoTableRowFour.getCell(1).setText(": ��");  
  
        //��������  
        XWPFTableRow infoTableRowFive = infoTable.createRow();  
        infoTableRowFive.getCell(0).setText("�־ӵ�");  
        infoTableRowFive.getCell(1).setText(": xx");  
  
  
        //�������֮��Ӹ�����  
        XWPFParagraph paragraph = document.createParagraph();  
        XWPFRun paragraphRun = paragraph.createRun();  
        paragraphRun.setText("\r");  
  
  
  
        //�����������  
        XWPFTable ComTable = document.createTable();  
  
  
        //�п��Զ��ָ�  
        CTTblWidth comTableWidth = ComTable.getCTTbl().addNewTblPr().addNewTblW();  
        comTableWidth.setType(STTblWidth.DXA);  
        comTableWidth.setW(BigInteger.valueOf(9072));  
  
        //����һ��  
        XWPFTableRow comTableRowOne = ComTable.getRow(0);  
        comTableRowOne.getCell(0).setText("��ʼʱ��");  
        comTableRowOne.addNewTableCell().setText("����ʱ��");  
        comTableRowOne.addNewTableCell().setText("��˾����");  
        comTableRowOne.addNewTableCell().setText("title");  
  
        //���ڶ���  
        XWPFTableRow comTableRowTwo = ComTable.createRow();  
        comTableRowTwo.getCell(0).setText("2016-09-06");  
        comTableRowTwo.getCell(1).setText("����");  
        comTableRowTwo.getCell(2).setText("seawater");  
        comTableRowTwo.getCell(3).setText("Java��������ʦ");  
  
        //��������  
        XWPFTableRow comTableRowThree = ComTable.createRow();  
        comTableRowThree.getCell(0).setText("2016-09-06");  
        comTableRowThree.getCell(1).setText("����");  
        comTableRowThree.getCell(2).setText("seawater");  
        comTableRowThree.getCell(3).setText("Java��������ʦ");  
  
  
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();  
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);  
  
        //���ҳü  
        CTP ctpHeader = CTP.Factory.newInstance();  
        CTR ctrHeader = ctpHeader.addNewR();  
        CTText ctHeader = ctrHeader.addNewT();  
        String headerText = "Java POI create MS word file.";  
        //ctHeader.setStringValue(headerText);  
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);  
        //����Ϊ�Ҷ���  
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);  
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];  
        parsHeader[0] = headerParagraph;  
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);  
  
  
        //���ҳ��  
        CTP ctpFooter = CTP.Factory.newInstance();  
        CTR ctrFooter = ctpFooter.addNewR();  
        CTText ctFooter = ctrFooter.addNewT();  
        String footerText = "http://blog.csdn.net/zhouseawater";  
        //ctFooter.setStringValue(footerText);  
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);  
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);  
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];  
        parsFooter[0] = footerParagraph;  
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);  
  
  
        document.write(out);  
        out.close();  
        System.out.println("create_table document written success.");  
    }  
  
  
}  