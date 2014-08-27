/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package apachepoitest;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Make this class read from XML files that contain the formatted "question"
 * @author Noel
 */
public class DocumentPropertyChecker {
    // Make better documentation later
    // Pass a list of paragraphs with the said propeties?
    // Pass a list of runs in a paragraph
    // check if run in passed text
    // double hash map
    // exists + properties
    public static Map<String, HashMap> checkIfStringsExistInParagraph(XWPFParagraph p, List<String> sl) {
        Map<String, HashMap> results = new HashMap();
        for (String s: sl) {
            results.put(s, new HashMap());
            results.get(s).put("EXISTS", p.getParagraphText().contains(s));
        }
        return results;
    }
    public static Boolean checkIfRunHasProperty(XWPFRun r, String property, String value) {
        try {
            switch (property) {
                case "COLOR":
                    return r.getColor().equals(value);
                case "FONT FAMILY":
                    return r.getFontFamily().equalsIgnoreCase(value);
                case "FONT SIZE":
                    return r.getFontSize() == Integer.parseInt(value);
                case "BOLD":
                    return r.isBold() == Boolean.getBoolean(value);
                case "ITALIC":
                    return r.isItalic() == Boolean.getBoolean(value);
                case "STRIKETHROUGH":
                    return r.isStrike() == Boolean.getBoolean(value);
                default:
                    System.out.println("Property " + property +  " does not exist!");
                    return false;
            }
        }
        catch (NullPointerException e) {
            return false;
        }
    }
    //Checking the runs, count if all instances contain the said formating
    /*Returns a map of strings with a map of properties with booleans as checks*/
    public static Map<String, HashMap> checkPropertiesofParagraphRuns(XWPFParagraph p, ArrayList<String> sl, Map<String, String> properties) {
        List<XWPFRun> rl = p.getRuns();
        Map<String, HashMap> results;
        
        //Check first if elements in sl are in p
        results = checkIfStringsExistInParagraph(p, sl);
        
        //Initialize counts to 0
        for (String s : sl) {
            for (String property : properties.keySet()) {
                results.get(s).put( property, 0);
            }
        }
        //For each existing string, 
        for (XWPFRun r : rl) {
            for (String s : sl) {
                //Skip string if it does't exist
                if (results.get(s).get("EXISTS").equals(true)) {
                    //For each property, check if it applies to the run
                    for (String property : properties.keySet()) {
                        if (checkIfRunHasProperty(r, property, properties.get(property)))
                        {
                            results.get(s).put( property, (int) results.get(s).get(property) + 1);
                        }
                    }
                }
            }
        }
        //Transform results to score
        for (String s : sl) {
            for (String property : properties.keySet()) {
                results.get(s).put( property, Integer.toString((int) results.get(s).get(property)) + "/" + Integer.toString(rl.size()));
            }
        }
        return results;
    }
    // check for strings that span whole paragraphs
    public static Map<String, HashMap> checkRunPropertiesOfParagraphs(List<XWPFParagraph> pl, ArrayList<String> sl, Map<String, String> properties) {
        Map<String, HashMap> results = new HashMap<>(), 
                             tempMap = new HashMap<>();
        ArrayList tempList;
        String removeString = "";
        
        // Initialize results, strings which were not found in the document are left as EXISTS : false
        for (String s : sl) {
            results.put(s, new HashMap<>());
            results.get(s).put("EXISTS", false);
        }
        
        for (XWPFParagraph p : pl) {
            for (String s : sl) {
                tempMap = null;
                //Will fail on typos, but pass on extra elements before or after string of interest
                //Need to change for typo toleration and exactness?
                if (p.getParagraphText().contains(s))
                {
                    tempList = new ArrayList();
                    tempList.add(s);
                    tempMap = checkPropertiesofParagraphRuns(p, tempList, properties);
                    results.put(s, tempMap.get(s));
                    removeString = s;
                    break;
                }
            }
            //Remove string if it has been evaluated
            if (tempMap != null) {
                sl.remove(removeString);
            }
        }
        return results;
    }
}
