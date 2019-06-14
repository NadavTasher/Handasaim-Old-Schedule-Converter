/*
 * Copyright (c) 2019 Nadav Tasher
 * https://github.com/NadavTasher/HandasaimScheduler
 * https://github.com/NadavTasher/HandasaimWeb
 */

package parser;

import okhttp3.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Pattern;

public class Schedule extends JSONObject {

    private Sheet sheet = null;

    public Schedule(String page) {
        // Add ringing times
        put("schedule", new int[]{465, 510, 555, 615, 660, 730, 775, 830, 875, 930, 975, 1020, 1065});
        // Initialize sheet
        sheet = initializeSheet(page);
        // Initialize messages
        put("messages", initializeMessages());
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

    private ArrayList<String> initializeMessages(Sheet sheet) {
        ArrayList<String> messages = new ArrayList<>();
        try {
            // Check type of sheet
//            ArrayList<Shape>
            if (sheet.getWorkbook() instanceof HSSFWorkbook) {
                sheet.createDrawingPatriarch();
                HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
                List<HSSFShape> shapes = patriarch.getChildren();
                for (int s = 0; s < shapes.size(); s++) {
                    if (shapes.get(s) instanceof HSSFTextbox) {
                        try {
                            HSSFShape mShape = shapes.get(s);
                            if (mShape != null) {
                                HSSFTextbox mTextShape = (HSSFTextbox) mShape;
                                HSSFRichTextString mString = mTextShape.getString();
                                if (mString != null) {
                                    messages.add(mString.getString());
                                }
                            }
                        } catch (NullPointerException ignored) {
                        }
                    }
                }
            } else {
                XSSFSheet convertedSheet = (XSSFSheet) sheet;
                XSSFDrawing drawing = convertedSheet.createDrawingPatriarch();
                List<XSSFShape> shapes = drawing.getShapes();
                for (int s = 0; s < shapes.size(); s++) {
                    if (shapes.get(s) instanceof XSSFSimpleShape) {
                        try {
                            XSSFSimpleShape mShape = (XSSFSimpleShape) shapes.get(s);
                            if (mShape != null) {
                                if (mShape.getText() != null) {
                                    String mString = mShape.getText();
                                    if (mString != null) {
                                        messages.add(mString);
                                    }
                                }
                            }
                        } catch (NullPointerException ignored) {
                        }
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
                if ((href.endsWith(".xls") || href.endsWith(".xlsx") && Pattern.compile("^(.(|.)-.(|.)\\..+)$").matcher(href).find())
                {
                    // Return 'href' attribute
                    return href;
                }
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private static int getReadingRow(Sheet sheet) {
        Cell secondCell = sheet.getRow(0).getCell(1);
        if (!readCell(secondCell).isEmpty()) {
            return 0;
        } else {
            return 1;
        }
    }

    private static int getReadingColumn(Sheet sheet) {
        return 1;
    }

    private void parseMessages(appcore.components.Schedule.Builder builder, Sheet sheet) {
        try {
            if (sheet.getWorkbook() instanceof HSSFWorkbook) {
                HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
                List<HSSFShape> shapes = patriarch.getChildren();
                for (int s = 0; s < shapes.size(); s++) {
                    if (shapes.get(s) instanceof HSSFTextbox) {
                        try {
                            HSSFShape mShape = shapes.get(s);
                            if (mShape != null) {
                                HSSFTextbox mTextShape = (HSSFTextbox) mShape;
                                HSSFRichTextString mString = mTextShape.getString();
                                if (mString != null) {
                                    builder.addMessage(mString.getString());
                                }
                            }
                        } catch (NullPointerException ignored) {
                        }
                    }
                }
            } else {
                XSSFSheet convertedSheet = (XSSFSheet) sheet;
                XSSFDrawing drawing = convertedSheet.createDrawingPatriarch();
                List<XSSFShape> shapes = drawing.getShapes();
                for (int s = 0; s < shapes.size(); s++) {
                    if (shapes.get(s) instanceof XSSFSimpleShape) {
                        try {
                            XSSFSimpleShape mShape = (XSSFSimpleShape) shapes.get(s);
                            if (mShape != null) {
                                if (mShape.getText() != null) {
                                    String mString = mShape.getText();
                                    if (mString != null) {
                                        builder.addMessage(mString);
                                    }
                                }
                            }
                        } catch (NullPointerException ignored) {
                        }
                    }
                }
            }
        } catch (Exception e) {
            builder.addMessage("Failed: Reading Messages");
        }
    }

    private static Sheet getSheet(File f) {
        try {
            if (f.toString().endsWith(".xls")) {
                POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(f));
                Workbook workBook = new HSSFWorkbook(fileSystem);
                Sheet foundSheet = null;
                for (int s = 0; s < workBook.getNumberOfSheets() && foundSheet == null; s++) {
                    Sheet current = workBook.getSheetAt(s);
                    if (current.getLastRowNum() - 1 > 0) {
                        foundSheet = current;
                    }
                }
                return foundSheet;
            } else {
                XSSFWorkbook workBook = new XSSFWorkbook(new FileInputStream(f));
                Sheet foundSheet = null;
                for (int s = 0; s < workBook.getNumberOfSheets() && foundSheet == null; s++) {
                    Sheet current = workBook.getSheetAt(s);
                    if (current.getLastRowNum() - 1 > 0) {
                        foundSheet = current;
                    }
                }
                return foundSheet;
            }
        } catch (IOException ignored) {
            return null;
        }
    }

    private static String readCell(Cell cell) {
        if (cell != null) {
            try {
                switch (cell.getCellType()) {
                    case STRING:
                        return cell.getStringCellValue();
                    case NUMERIC:
                        return String.valueOf((int) cell.getNumericCellValue());
                    case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue());
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return "";
    }

}
