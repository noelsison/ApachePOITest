/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package apachepoitest;

import com.eclipsesource.json.JsonArray;
import com.eclipsesource.json.JsonObject;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 *
 * @author Noel
 */
public class ApachePOITest {
    
    @SuppressWarnings("empty-statement")
    public static void main(String[] args) {
        try {
            // Open document to check
            /*
            Writer fw = new FileWriter("C:\\Users\\Noel\\Documents\\NetBeansProjects\\ApachePOITest\\test1.json"); 
            JsonObject jo = new JsonObject().add( "name", "John" ).add( "age", 23 );
         
            JsonArray ja = new JsonArray().add( "John" ).add( 23 );
            jo.writeTo(fw);
            ja.writeTo(fw);
            fw.close();
            */
            XWPFDocument docx1 = new XWPFDocument(new FileInputStream(new File("C:\\Users\\Noel\\Documents\\NetBeansProjects\\ApachePOITest\\resume_only.docx")));
            
            // Put the following to an XML file that contains strings to check with respective properties to check
            // Question 1 in Level 1
            // Initialize strings to find
            ArrayList<String> sl = new ArrayList();
            String[] tl = {"Melissa Martin", "555 West Main St.", "Sampaloc, Metro Manila", "Phone: 312-312-3123", "E-mail: TeachMartin@email.com"};
            sl.addAll(Arrays.asList(tl));
            
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
            System.out.println("1. " + results.toString());
            
            //2
            tl = new String[]{"Summary", "Educational Background", "Related Work Experience", "Additional Work Experience"};
            sl.addAll(Arrays.asList(tl));
            
            //properties
            properties = new HashMap();
            properties.put("BOLD", "true");
            
            results = DocumentPropertyChecker.checkRunPropertiesOfParagraphs(docx1.getParagraphs(), sl, properties);
            System.out.println("2. " + results.toString());
            
            //3
            tl = new String[]{"Holds Bachelor's Degree in Music and Education with TEFL certification", 
                              "5 years experience in teaching Englsih to Spanish speaking students ages 12 and up",
                              "Exceptional skills in teaching English and Spanish language",
                              "Bachelor of Music; Univeristy of Sto. Tomas 2004",
                              "Bachelor of Science in Education; Univerity of the Philippines 2008"};
            sl.addAll(Arrays.asList(tl));
            properties = new HashMap();
            properties.put("LINE SPACING", "1.5");
            
            results = DocumentPropertyChecker.checkPropertiesOfParagraphs(docx1.getParagraphs(), sl, properties);
            System.out.println("3. " + results.toString());
            
            //4
            tl = new String[]{"2008-2011"};
            sl.addAll(Arrays.asList(tl));
            results = DocumentPropertyChecker.checkIfStringExistsInParagraphs(docx1.getParagraphs(), sl);
            System.out.println("4. " + results.toString());
            
            //5
            tl = new String[]{"St. Peter's University",
                              "2011 – Present",
                              "Teaches English and Spanish to students ages 15 and up",
                              "Creates course materials, including exams, quizzes and visual aids used by all teachers throughout the organization",
                              "Initiates programs focused in improving grammar and active listening, writing and speaking skills of students"};
            sl.addAll(Arrays.asList(tl));
            properties = new HashMap();
            properties.put("NUMBERING FORMAT", "bullet"); 
            
            results = DocumentPropertyChecker.checkPropertiesOfParagraphs(docx1.getParagraphs(), sl, properties);
            System.out.println("5. " + results.toString());
            
            //6
            tl = new String[]{"Black Pen Movement \u00AE"};
            sl.addAll(Arrays.asList(tl));
            results = DocumentPropertyChecker.checkIfStringExistsInParagraphs(docx1.getParagraphs(), sl);
            System.out.println("6. " + results.toString());
            
            //7
            properties = new HashMap();
            properties.put("MARGIN TOP", "2");
            properties.put("MARGIN BOTTOM", "2");
            properties.put("MARGIN LEFT", "2");
            properties.put("MARGIN RIGHT", "2");
            
            results = DocumentPropertyChecker.checkPropertiesOfDocument(docx1, properties);
            System.out.println("7. " + results.toString());
            
            //8
            properties = new HashMap();
            properties.put("ALIGN", "both");
            
            results = DocumentPropertyChecker.checkPropertiesOfAllParagraphs(docx1.getParagraphs(), properties);
            System.out.println("8. " + results.toString());
            
        } catch (IOException ex) {
            Logger.getLogger(ApachePOITest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
