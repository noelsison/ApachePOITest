/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package apachepoitest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 *
 * @author Noel
 */
public class ApachePOITest {
    public static void showParagraphElementProperties(List<XWPFRun> rl)
    {
        System.out.println("\nELEMENTS: ");
        int counter = 1;
        for (XWPFRun r : rl) {
            if(r.toString().trim().length() > 0) {
                System.out.println("#" + counter++ + ": " + r.toString());
            }
            else {
                //System.out.println("#" + counter++ + ": <SPACE>");
                break;
            }
            if (r.getColor() != null) {
                System.out.println("COLOR: " + r.getColor());
            }
            if (r.getFontFamily() != null) {
                System.out.println("FONT: " + r.getFontFamily());
            }
            if (r.getFontSize() > 0) {
                System.out.println("FONT SIZE: " + r.getFontSize());
            }
            if (r.getPictureText().length() > 0) {
                System.out.println("PIC TEXT: " + r.getPictureText());
            }
            if (r.getTextPosition() > 0) {
                System.out.println("TEXT POS: " + r.getTextPosition());
            }
            if (r.isBold()) {
                System.out.println("BOLD: " + r.isBold());
            }
            if (r.isItalic()) {
                System.out.println("ITALIC: " + r.isItalic());
            }
            if (r.isStrike()) {
                System.out.println("STRIKETHROUGH: " + r.isStrike());
            }
            if (! r.getUnderline().toString().equals("NONE")) {
                System.out.println("UNDERLINE: " + r.getUnderline().toString());
            }
            if (! r.getSubscript().toString().equals("BASELINE")) {
                System.out.println("Subscript: " + r.getSubscript().toString());
            }
            System.out.println("");
        }
    }
    public static void showParagraphProperties(List<XWPFParagraph> lp)
    {
        int i1 = 1;
        for (XWPFParagraph p : lp) {
            //System.out.println(p.getStyleID() + " " + sl1.getStyle(p.getStyleID()).getCTStyle().xmlText());
            System.out.println("____________________________________");
            if(p.getParagraphText().trim().length() > 0) {
                System.out.println("\n#" + i1++ + " LINE: " + p.getParagraphText());
                System.out.println("ALIGNMENT: " + p.getAlignment().toString());
                //Uncomment to display other properties
                /*
                System.out.println("BORDER BETWEEN: " + p.getBorderBetween().toString());
                System.out.println("BORDER BOTTOM: " + p.getBorderBottom().toString());
                System.out.println("BORDER LEFT: " + p.getBorderLeft().toString());
                System.out.println("BORDER RIGHT: " + p.getBorderRight().toString());
                System.out.println("BORDER TOP: " + p.getBorderTop().toString());
                System.out.println("BODY ELEMENT TYPE: " + p.getElementType().toString());
                System.out.println("FOOTNOTE: " + p.getFootnoteText());
                System.out.println("INDENTATION 1ST LINE: " + p.getIndentationFirstLine());
                System.out.println("INDENTATION HANGING: " + p.getIndentationHanging());
                System.out.println("INDENTATION LEFT: " + p.getIndentationLeft());
                System.out.println("INDENTATION RIGHT: " + p.getIndentationRight());
                System.out.println("NUMBERING FORMAT: " + p.getNumFmt());
                System.out.println("NUMERIC STYLE ILVL: " + p.getNumIlvl());
                System.out.println("SPACING AFTER: " + p.getSpacingAfter());
                System.out.println("SPACING AFTER LINES: " + p.getSpacingAfterLines());
                System.out.println("SPACING BEFORE: " + p.getSpacingBefore());
                System.out.println("SPACING BEFORE LINES: " + p.getSpacingBeforeLines());
                System.out.println("SPACING LINE RULE: " + p.getSpacingLineRule());
                System.out.println("VERTICAL ALIGNMENT: " + p.getVerticalAlignment());
                */
            }   // can also use .searchText to look for a string
            else {
                //System.out.println("\n#" + i1++ + " LINE: <SPACE>");
            }
                
            showParagraphElementProperties(p.getRuns());
        }
    }
    public static void showTableProperties(List<XWPFTable> lt)
    {
        for (XWPFTable t: lt) {
            System.out.println("TABLE: ");
            //System.out.println("COL BAND SIZE: " + t.getColBandSize());
            //System.out.println("ROW BAND SIZE: " + t.getRowBandSize());
            System.out.println("NO. OF ROWS: " + t.getNumberOfRows());
            System.out.println("WIDTH: " + t.getWidth());
        }
    }
    public static void showProperties(XWPFDocument docx) {
        List<XWPFParagraph> lp = docx.getParagraphs();
        List<XWPFTable> lt = docx.getTables();
        showParagraphProperties(lp);
        showTableProperties(lt);
    }
    /**
     * @param args the command line arguments
     */
    public static Boolean checkElementProperty(List<XWPFRun> rl, String key, Object value)
    {
        Boolean exists = false;
        Object o = false;
        for (XWPFRun r : rl) {
            System.out.println("                RUN: " + r.toString());
            switch (key) {
                case "COLOR":
                    o = r.getColor();
                    break;
                case "FONT FAMILY":
                    o = r.getFontFamily();
                    break;
                case "FONT SIZE":
                    o = r.getFontSize();
                    break;
                case "BOLD":
                    o = r.isBold();
                    break;
                default:
                    return false;
            }
            if (o != null) {
                break;
            }
        }
        //System.out.println(o);
        return o.equals(value);
    }
    public static void main(String[] args) {
        
        try {
            XWPFDocument docx1 = new XWPFDocument(new FileInputStream(new File("C:\\Users\\Noel\\Documents\\NetBeansProjects\\ApachePOITest\\resume_only.docx")));
            //showProperties(docx1);
            
            List<XWPFParagraph> lp = docx1.getParagraphs();
            Boolean temp_bool;
            temp_bool = true;
            
            String q1_strings[] = {"Melissa Martin", "555 West Main St.", "Sampaloc, Metro Manila", "Phone: 312-312-3123", "E-mail: TeachMartin@email.com"};
            //FONT MV Boli
            //SIZE 12
            Map to_check = new HashMap();
            to_check.put("FONT FAMILY", "MV Boli");
            to_check.put("FONT SIZE", (int) 12);
            
            for (XWPFParagraph p: lp) {
                for (String s: q1_strings) {
                    if (p.getParagraphText().trim().equals(s)) {
                        System.out.println("DETECTED: " + p.getParagraphText());
                        System.out.println("RUNS: " + p.getRuns().size());
                        for (Object key: to_check.keySet()) {
                            System.out.println("    PROPERTY: " + key);
                            temp_bool = checkElementProperty(p.getRuns(), key.toString(), to_check.get(key));
                            System.out.println("    EXISTS: " + temp_bool);
                        }
                    }
                }
            }
        } catch (IOException ex) {
            Logger.getLogger(ApachePOITest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
