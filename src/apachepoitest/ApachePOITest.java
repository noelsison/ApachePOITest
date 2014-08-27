/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package apachepoitest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
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
    
    public static void main(String[] args) {
        try {
            // Open document to check
            XWPFDocument docx1 = new XWPFDocument(new FileInputStream(new File("C:\\Users\\Noel\\Documents\\NetBeansProjects\\ApachePOITest\\resume_only.docx")));
            
            // Put the following to an XML file that contains strings to check with respective properties to check
            // Question 1 in Level 1
            // Initialize strings to find
            ArrayList<String> sl = new ArrayList();
            String[] tl = {"Melissa Martin", "555 West Main St.", "Sampaloc, Metro Manila", "Phone: 312-312-3123", "E-mail: TeachMartin@email.com"};
            for (String s : tl) {
                sl.add(s);
            }
            
            // Initialize properties these strings should have
            Map properties = new HashMap();
            properties.put("FONT FAMILY", "MV Boli");
            properties.put("FONT SIZE", "12");
            
            // We go through all paragraphs of the document and check for the presence of the strings
            // If they are present, check if the properties given above are present
            // Result is displayed as String = {Property1 = Score1, Property2 = Score2, ...}
            // Scores are determined by the number of elements within the paragraph which follows the given formatting
            Map<String, HashMap> results;
            results = DocumentPropertyChecker.checkRunPropertiesOfParagraphs(docx1.getParagraphs(), sl, properties);
            System.out.println(results.toString());
            
        } catch (IOException ex) {
            Logger.getLogger(ApachePOITest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
