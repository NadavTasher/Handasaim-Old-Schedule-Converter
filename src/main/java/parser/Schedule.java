/*
 * Copyright (c) 2019 Nadav Tasher
 * https://github.com/NadavTasher/HandasaimScheduler
 * https://github.com/NadavTasher/HandasaimWeb
 */

package parser;

import appcore.components.Classroom;
import appcore.components.Subject;
import okhttp3.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.regex.Pattern;

public class Schedule extends JSONObject {

    private static String[] DAYS = {
            "ראשון",
            "שני",
            "שלישי",
            "רביעי",
            "חמישי",
            "שישי",
            "שבת"
    };

    private static String[][] TRIMMERS = {
            {", ", " · "},
            {",", " · "},
            {"מתמטיקה", "מתמט'"},
            {"טכניונית", "טכ'"},
            {"מעבדה", "מע'"}
    };


    public Schedule(String page) {
        // Add ringing times
        put("schedule", new int[]{465, 510, 555, 615, 660, 730, 775, 830, 875, 930, 975, 1020, 1065});
        // Initialize sheet
        Sheet sheet = initializeSheet(page);
        // Initialize messages
        put("messages", parseMessages(sheet));
        // Initialize day
        put("day", parseDay(sheet));
        // Initialize grades
        put("grades", parseGrades(sheet));
        // Initialize teachers
        put("teachers", parseTeachers(get("grades")));

    }

    public String export() {
        return toString();
    }

    private void addError(String error) {
        // Read array from structure
        JSONArray errors = optJSONArray("errors");
        // Initialize if null
        if (errors == null) errors = new JSONArray();
        // Push 'error' to array
        errors.put(error);
        // Write array to structure
        put("errors", errors);
    }

    private ArrayList<String> parseMessages(Sheet sheet) {
        ArrayList<String> messages = new ArrayList<>();
        try {
            // Check type of sheet
            if (sheet.getWorkbook() instanceof HSSFWorkbook) {
                // Get messages list
                List<HSSFShape> shapes = ((HSSFSheet) sheet).createDrawingPatriarch().getChildren();
                // Loop through shapes
                for (HSSFShape shape : shapes) {
                    if (shape instanceof HSSFTextbox) {
                        // Add to list
                        messages.add(((HSSFTextbox) shape).getString().getString());
                    }
                }
            } else {
                // Get messages list
                List<XSSFShape> shapes = ((XSSFSheet) sheet).createDrawingPatriarch().getShapes();
                // Loop through shapes
                for (XSSFShape shape : shapes) {
                    if (shape instanceof XSSFSimpleShape) {
                        // Add to list
                        messages.add(((XSSFSimpleShape) shape).getText());
                    }
                }
            }
        } catch (Exception ignored) {
            addError("Failed reading messages");
        }
        return messages;
    }

    private Sheet initializeSheet(String page) {
        // Fetch link from page
        String link = initializeLink(page);
        // Verify that a link has been found
        if (link != null) {
            try {
                // Connect to Excel file
                OkHttpClient client = null;
                // Check for connection protocol
                if (link.startsWith("https")) {
                    // Create a client for HTTPS SSL enabled requests
                    client = new OkHttpClient.Builder().connectionSpecs(Collections.singletonList(ConnectionSpec.MODERN_TLS)).build();
                } else {
                    // Create a client for standard HTTP requests
                    client = new OkHttpClient();
                }
                // Fetch file from link
                Response response = client.newCall(new Request.Builder().url(link).get().build()).execute();
                // Extract InputStream from response
                InputStream stream = response.body() != null ? response.body().byteStream() : null;
                // Verify stream integrity... kinda
                if (stream != null) {
                    Workbook workbook = null;
                    // Create a workbook of right type by examining file type
                    if (link.endsWith(".xls")) {
                        workbook = new HSSFWorkbook(new POIFSFileSystem(stream));
                    } else if (link.endsWith(".xlsx")) {
                        workbook = new XSSFWorkbook(stream);
                    }
                    // Verify that the file was actually an Excel file, won't proceed if type detected isn't supported
                    if (workbook != null) {
                        // Loop through sheets
                        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                            Sheet current = workbook.getSheetAt(s);
                            // Check for minimum rows
                            if (current.getLastRowNum() > 1) {
                                // Return sheet
                                return current;
                            }
                        }
                    }
                } else {
                    addError("Null Excel response body");
                }
            } catch (Exception ignored) {
            }
            return null;
        } else {
            addError("Schedule link not found");
        }
        return null;
    }

    private String initializeLink(String link) {
        // Treat link as schedule page
        try {
            // Connect to page, 7.5 second timeout.
            Document document = Jsoup.connect(link).timeout(7500).get();
            // Look for 'a' tags
            Elements elements = document.select("a");
            // Loop on 'a' tags
            for (Element element : elements) {
                // Pull 'href' attribute from 'a' tag
                String href = element.attr("href");
                // Check 'href' attribute file format against known excel file types, and verify that it is indeed a schedule Excel (other Excel files might exist on the homepage.).
                if ((href.endsWith(".xls") || href.endsWith(".xlsx") && Pattern.compile("^(.(|.)-.(|.)\\..+)$").matcher(href).find())) {
                    // Return 'href' attribute
                    return href;
                }
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private String subject(String untrimmed) {
        // Loop through replacements to shorten names
        for (String[] replacement : TRIMMERS) {
            if (replacement.length == 2)
                untrimmed = untrimmed.replaceAll(replacement[0], replacement[1]);
        }
        return untrimmed;
    }

    private JSONObject parseGrades(Sheet sheet) {
        JSONObject grades = new JSONObject();
        // If the cell after day name is empty, first row is 1, else 0
        int firstRow = parseCell(sheet, 1, 0).isEmpty() ? 1 : 0;
        // First column is always after the hour column
        int firstColumn = 1;
        int lastRow = sheet.getLastRowNum();
        int lastColumn = sheet.getRow(firstRow).getLastCellNum();
        // Loop through columns
        for (int c = firstColumn; c < lastColumn; c++) {
            // Create grade structure
            JSONObject grade = new JSONObject();
            // Parse minimal grade name
            String name = parseCell(sheet, c, firstRow).split(" ")[0];
            // Put parsed grade number (7-12)
            grade.put("grade", parseGrade(name));
            // Create subjects structure
            JSONObject subjects = new JSONObject();
            // Loop through rows, first row is the one after the title
            for (int r = firstRow + 1; r < lastRow; r++) {
                // Get cell value
                String text = parseCell(sheet, c, r);
                // Check if cell is not empty
                if (!text.isEmpty()) {
                    // Create subject and teachers structure
                    JSONObject subject = new JSONObject();
                    JSONArray teachers = new JSONArray();
                    // Split cell to rows
                    String[] rows = text.split("(|\r)(\n)");
                    // Put trimmed subject name in subject
                    subject.put("name", subject(rows[0]));
                    // Check if cell has more then one row
                    if (rows.length > 1) {
                        // Loop through last row divided by commas and add to teachers
                        for (String teacher : rows[rows.length - 1].split(",")) teachers.put(teacher);
                    }
                    // Put teachers in subject
                    subject.put("teachers", teachers);
                    // Put subject in subjects as hour number (0-13+), for easy scanning
                    subjects.put(String.valueOf(r - (firstRow + 1)), subject);
                }
            }
            // Put subjects in grade
            grade.put("subjects", subjects);
            // Put grade in grades
            grades.put(name, grade);
        }
    }

    private int parseGrade(String name) {
        // Parse grade from name
        if (name.startsWith("ז")) return 7;
        if (name.startsWith("ח")) return 8;
        if (name.startsWith("ט")) return 9;
        if (name.startsWith("יב")) return 12;
        if (name.startsWith("יא")) return 11;
        if (name.startsWith("י")) return 10;
        return 0;
    }

    private String parseCell(Sheet sheet, int x, int y) {
        return parseCell(sheet.getRow(y).getCell(x));
    }

    private String parseCell(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    return String.valueOf((int) cell.getNumericCellValue());
            }
        }
        return "";
    }

    private int parseDay(Sheet sheet) {
        // Get cell value
        String day = parseCell(sheet, 0, 0);
        // Loop through days and compare until match found and return the number of the day (1-7, on error 0)
        for (int d = 0; d < DAYS.length; d++) {
            if (DAYS[d].equals(day)) return d + 1;
        }
        return 0;
    }
}
