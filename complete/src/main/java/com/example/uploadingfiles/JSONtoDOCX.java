package com.example.uploadingfiles;

import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.web.multipart.MultipartFile;
import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.rendering.ImageType;
//import org.apache.pdfbox.rendering.PDFRenderer;


import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Path;
import java.util.List;
import java.util.Scanner;

public class JSONtoDOCX {
    private XWPFDocument doc;
    private JSONObject JSON;
    private XWPFParagraph titles;

    private MultipartFile file;

    private File output;

    public JSONtoDOCX(MultipartFile file) {
        this.file = file;
    }

    public void main() throws Exception {

        //Scanner scanner = askForFile();
        //File json = file.getResource().getFile();
        String fileName = file.getOriginalFilename();
        System.out.println(fileName);
        File json = Path.of("upload-dir", fileName).toFile();
        //File json = new File("resources/example-json-to-docx-v1.json");
        Scanner scanner = new Scanner(json);
        StringBuilder full = new StringBuilder();
        while (scanner.hasNext()) {
            full.append(scanner.next()).append(" ");
        }
        JSON = new JSONObject(full.toString()); // complete/src/main/resources
        doc = new XWPFDocument(new FileInputStream(Path.of("src", "main", "resources", "FirstPage2.0.docx").toFile()));

        XWPFDocument template = new XWPFDocument(new FileInputStream(Path.of("src", "main", "resources", "Template.docx").toFile()));
        XWPFStyles newStyles = doc.createStyles();
        newStyles.setStyles(template.getStyle());

        FrontPage();
        if (JSON.getBoolean("TableOfContent")) {
            TableOfContent();
        }
        HeaderFooter();
        Sections();
        //JSONObject attachments = JSON.getJSONObject("attachments");
//        if(attachments != null)
//            Attachments(attachments);


        FileOutputStream out = new FileOutputStream(Path.of("upload-dir", JSON.getString("fileName")).toFile());
        doc.write(out);

        System.out.println("works");
    }

    private Scanner askForFile() {
        try {
            Scanner input = new Scanner(System.in);
            System.out.println("Please put the file location here:");
            String location = input.next();
            File output = new File(location);
            return new Scanner(output);
        } catch (Exception e) {
            System.out.println("File not found, please try again");
            return askForFile();
        }
    }

    private void FrontPage() {
        JSONObject frontPage = JSON.getJSONObject("frontPage");

        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.equals("TITLE")) {
                        text = text.replace("TITLE", frontPage.getString("title"));
                        r.setText(text, 0);
                    } else if (text != null && text.equals("TITLE2")) {
                        text = text.replace("TITLE2", frontPage.getString("titleSub"));
                        r.setText(text, 0);
                    } else if (text != null && text.equals("DATE1")) {
                        text = text.replace("DATE1", frontPage.getString("valuationDate"));
                        r.setText(text, 0);
                    } else if (text != null && text.equals("DATE2")) {
                        text = text.replace("DATE2", frontPage.getString("reportDate"));
                        r.setText(text, 0);
                    }
                }
            }
        }
    }

    private void TableOfContent() {
        doc.createTOC();
        addCustomHeadingStyle(doc, "heading 1", 1);
        addCustomHeadingStyle(doc, "heading 2", 2);
        addCustomHeadingStyle(doc, "heading 3", 3);
        titles = doc.createParagraph();
        CTP ctP = titles.getCTP();
        CTSimpleField toc = ctP.addNewFldSimple();
        toc.setInstr("TOC \\h");
        toc.setDirty(true);
        XWPFRun run = titles.createRun();
        run.addBreak(BreakType.PAGE);
    }

    private void HeaderFooter() throws IOException, InvalidFormatException {
        doc.createHeader(HeaderFooterType.FIRST);


        //header en footer
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        if (headerFooterPolicy == null) headerFooterPolicy = doc.createHeaderFooterPolicy();
        //header
        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph headerParagraph = header.createParagraph();
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = headerParagraph.createRun();
        JSONObject JSONHeader = JSON.getJSONObject("documentHeader");
        //picture
        String url = JSONHeader.getString("logo");
        saveFileFromUrl(url,"src\\main\\resources\\header.jpg");
        File image = Path.of("src/main/resources/header.jpg").toFile();
        FileInputStream imageData = new FileInputStream(image);
        int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
        String imageFileName = image.getName();
        int width = 600;
        int height = 152/4;
        run.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));
        run.addBreak();
        run.setText(JSONHeader.getString("headerText"));
        //footer
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph footerParagraph = footer.createParagraph();
        footerParagraph.setAlignment(ParagraphAlignment.CENTER);
        run = footerParagraph.createRun();
        JSONObject JSONFooter = JSON.getJSONObject("documentFooter");
        run.setText(JSONFooter.getString("footerText"));
        run.addBreak();
        boolean pageNumbers = JSONFooter.getBoolean("pageNumbers");
        if (pageNumbers) {
            run.setText("Page ");
            footerParagraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
            run = footerParagraph.createRun();
            run.setText(" of ");
            footerParagraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");
        }

    }

    private void Sections() {
        JSONArray sectionsArray = JSON.getJSONArray("sections");
        titles = doc.createParagraph();
        XWPFParagraph texts;


        for (int i = 0; i < sectionsArray.length(); i++) {
            JSONObject section = (JSONObject) sectionsArray.get(i);
            titles = doc.createParagraph();

            XWPFRun title = titles.createRun();

            titles.setStyle("heading 1");
            setParagraphShading(titles, "005B82");
            title.setBold(true);
            title.setFontSize(18);
            title.setColor("FFFFFF");
            if(i != 0)
                title.setText(i + ". " + section.getString("title"));
            else
                title.setText(section.getString("title"));
            try {
                String introText = section.getString("introText");
                texts = doc.createParagraph();
                XWPFRun introTextRun = texts.createRun();
                introTextRun.setText(introText);
            } catch (Exception e) {
                //no introText;
            }

            JSONArray subsections = null;

            try {
                subsections = section.getJSONArray("subSections");
                JSONObject check = (JSONObject) subsections.get(0);
                check.getString("title");
            } catch (JSONException e) {
                try {
                    JSONArray bullets = section.getJSONArray("bullets");
                    CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
                    cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));

                    CTLvl cTLvl = cTAbstractNum.addNewLvl();
                    cTLvl.setIlvl(BigInteger.valueOf(0)); // set indent level 0
                    cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
                    cTLvl.addNewLvlText().setVal("•");


                    XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
                    XWPFNumbering numbering = doc.createNumbering();
                    BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
                    BigInteger numID = numbering.addNum(abstractNumID);

                    XWPFParagraph text;

                    for (Object bullet : bullets) {
                        String temp = (String) bullet;
                        text = doc.createParagraph();
                        text.setNumID(numID);
                        XWPFRun subIntroText = text.createRun();
                        subIntroText.setText(temp);
                    }


                } catch (JSONException ignored) {
                    //no bullets
                }
            }

            if(subsections != null)
                SubSections(subsections, i);


        }

    }

    private void SubSections(JSONArray subsections, int i) {
        XWPFParagraph subtitle;
        XWPFParagraph subtext;

        for (int j = 0; j < subsections.length(); j++) {
            JSONObject subsection = subsections.getJSONObject(j);


            //title stuff
            try {
                String title = subsection.getString("title");
                subtitle = doc.createParagraph();
                XWPFRun subTitleRun = subtitle.createRun();
                setParagraphShading(subtitle, "005B82");
                subTitleRun.setBold(true);
                subTitleRun.setFontSize(14);
                subTitleRun.setColor("FFFFFF");

                if(i != 0)
                    subTitleRun.setText(i + "." + (j+1) + ". " + title);
                else
                    subTitleRun.setText(title);

                subtitle.setStyle("heading 2");
            } catch (JSONException ignored) {

            }
            //intro text stuff
            try {
                String introText = subsection.getString("introText");
                subtext = doc.createParagraph();
                XWPFRun subIntroText = subtext.createRun();
                addlongTextToRun(introText, subIntroText);
                //subIntroText.setText(introText);
            } catch (JSONException ignored) {
            }
            //bulleted list
            try {
                JSONArray bullets = subsection.getJSONArray("bullets");
                CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
                cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));

                CTLvl cTLvl = cTAbstractNum.addNewLvl();
                cTLvl.setIlvl(BigInteger.valueOf(0)); // set indent level 0
                cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
                cTLvl.addNewLvlText().setVal("•");


                XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
                XWPFNumbering numbering = doc.createNumbering();
                BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
                BigInteger numID = numbering.addNum(abstractNumID);

                for (Object bullet : bullets) {
                    String temp = (String) bullet;
                    subtext = doc.createParagraph();
                    subtext.setNumID(numID);
                    XWPFRun subIntroText = subtext.createRun();
                    subIntroText.setText(temp);
                }


            } catch (JSONException e) {
                //no bullets
            }
            //table
            try {
                JSONObject tableJSON = subsection.getJSONObject("table");
                XWPFTable table = doc.createTable();

                XWPFStyles styles = doc.getStyles();
                XWPFStyle style = styles.getStyleWithName("Table Grid");
                table.setStyleID(style.getStyleId());
                String check = "";
                table.setBottomBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setLeftBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setRightBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setTopBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setInsideVBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setInsideHBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                if(i == 0) table.setWidth("50%");


                try{
                    JSONArray headers = tableJSON.getJSONArray("dataHeaders");
                    check = headers.getString(0);
                    XWPFTableRow tableRowOne = table.getRow(0);
                    tableRowOne.getCell(0).setText(headers.getString(0));
                    for (int k = 1; k < headers.length(); k++) {
                        tableRowOne.addNewTableCell().setText(headers.getString(k));
                    }
                }
                catch (JSONException e){
                    //no reader rows
                }

                JSONArray data = tableJSON.getJSONArray("dataRows");
                XWPFTableRow tableRowOne = table.getRow(0);
                tableRowOne.setHeight((int)(1440/3));
                tableRowOne.getCell(0).setText(data.getJSONArray(0).getString(0));
                for (int k = 1; k < data.getJSONArray(0).length(); k++) {
                    tableRowOne.addNewTableCell().setText(data.getJSONArray(0).getString(k));
                }
                for (int k = 1; k < data.length(); k++) {
                    JSONArray row = data.getJSONArray(k);
                    XWPFTableRow nextRow = table.createRow();


                    for (int l = 0; l < row.length(); l++) {
                        XWPFTableCell cell = nextRow.getCell(l);
                        if(l == 0) cell.setWidth("30%");
                        cell.setText(row.getString(l));
                        //if(l != row.length() -1 && check.equals("")) nextRow.addNewTableCell();
                    }
                    nextRow.setHeight((int)(1440/3));
                    nextRow.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.EXACT);
                    System.out.println(nextRow.getHeight());
                }
                data = tableJSON.getJSONArray("dataTotals");
                for (int k = 0; k < data.length(); k++) {
                    JSONArray row = data.getJSONArray(k);
                    XWPFTableRow nextRow = table.createRow();
                    for (int l = 0; l < row.length(); l++) {

                        nextRow.getCell(l).setText(row.getString(l));
                        if(l != row.length() -1 && check.equals("")) nextRow.addNewTableCell();
                    }
                }
            } catch (JSONException ignored) {

            }
            catch (Exception e){
                e.printStackTrace();
            }

            //image
            try {
                JSONObject image = subsection.getJSONObject("image");
                XWPFParagraph imageParagraph = doc.createParagraph();
                XWPFRun imageRun = imageParagraph.createRun();
                String alignment = image.getString("align");
                switch (alignment) {
                    case "C" -> imageParagraph.setAlignment(ParagraphAlignment.CENTER);
                    case "R" -> imageParagraph.setAlignment(ParagraphAlignment.RIGHT);
                    case "L" -> imageParagraph.setAlignment(ParagraphAlignment.LEFT);
                }
                String url = image.getString("data");
//                try{
//                    //saveFileFromUrl(url, "src\\main\\resources\\temp.png");
//                }
//                //catch (MalformedURLException ignored){}
                File imageFile = Path.of("src\\main\\resources\\temp.png").toFile();
                FileInputStream imageData = new FileInputStream(imageFile);
                int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
                String imageFileName = imageFile.getName();
                int width = image.getInt("maxWidth");
                int height = image.getInt("maxHeight");
                imageRun.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));

            } catch (JSONException ignored) {
            }
            catch (Exception e){
                e.printStackTrace();
            }

            try{
                JSONArray nested = subsection.getJSONArray("subSections");
                nestedSubSection(nested);
            }
            catch (JSONException ignored){}

        }

    }

    private void nestedSubSection(JSONArray subsections){
        XWPFParagraph subtitle;
        XWPFParagraph subtext;

        for (int j = 0; j < subsections.length(); j++) {
            JSONObject subsection = subsections.getJSONObject(j);
            //title stuff
            try {
                String title = subsection.getString("title");
                subtitle = doc.createParagraph();
                XWPFRun subTitleRun = subtitle.createRun();
                subTitleRun.setBold(true);
                subTitleRun.setFontSize(11);
                subTitleRun.setText(title);
                subtitle.setStyle("heading 3");
            } catch (JSONException ignored) {

            }
            //intro text stuff
            try {
                String introText = subsection.getString("introText");
                subtext = doc.createParagraph();
                XWPFRun subIntroText = subtext.createRun();
                addlongTextToRun(introText, subIntroText);
                //subIntroText.setText(introText);
            } catch (JSONException ignored) {
            }
            //bulleted list
            try {
                JSONArray bullets = subsection.getJSONArray("bullets");
                CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
                cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));

                CTLvl cTLvl = cTAbstractNum.addNewLvl();
                cTLvl.setIlvl(BigInteger.valueOf(0)); // set indent level 0
                cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
                cTLvl.addNewLvlText().setVal("•");

                XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
                XWPFNumbering numbering = doc.createNumbering();
                BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
                BigInteger numID = numbering.addNum(abstractNumID);

                for (Object bullet : bullets) {
                    String temp = (String) bullet;
                    subtext = doc.createParagraph();
                    subtext.setNumID(numID);
                    XWPFRun subIntroText = subtext.createRun();
                    subIntroText.setText(temp);
                }


            } catch (JSONException e) {
                //no bullets
            }
            //table
            try {
                JSONObject tableJSON = subsection.getJSONObject("table");
                XWPFTable table = doc.createTable();



                XWPFStyles styles = doc.getStyles();
                XWPFStyle style = styles.getStyleWithName("Table Grid");
                table.setStyleID(style.getStyleId());
                String check = "";
                table.setBottomBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setLeftBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setRightBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setTopBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setInsideVBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);
                table.setInsideHBorder(XWPFTable.XWPFBorderType.NONE,0,0,null);

                try{
                    JSONArray headers = tableJSON.getJSONArray("dataHeaders");
                    check = headers.getString(0);
                    XWPFTableRow tableRowOne = table.getRow(0);
                    tableRowOne.getCell(0).setText(headers.getString(0));
                    for (int k = 1; k < headers.length(); k++) {
                        tableRowOne.addNewTableCell().setText(headers.getString(k));
                    }
                }
                catch (JSONException e){
                    //no reader rows
                }

                JSONArray data = tableJSON.getJSONArray("dataRows");
                XWPFTableRow tableRowOne = table.getRow(0);
                tableRowOne.getCell(0).setText(data.getJSONArray(0).getString(0));
                for (int k = 1; k < data.getJSONArray(0).length(); k++) {
                    tableRowOne.addNewTableCell().setText(data.getJSONArray(0).getString(k));
                }
                for (int k = 1; k < data.length(); k++) {
                    JSONArray row = data.getJSONArray(k);
                    XWPFTableRow nextRow = table.createRow();
                    for (int l = 0; l < row.length(); l++) {

                        nextRow.getCell(l).setText(row.getString(l));
                        //if(l != row.length() -1 && check.equals("")) nextRow.addNewTableCell();
                    }
                }
                data = tableJSON.getJSONArray("dataTotals");
                for (int k = 0; k < data.length(); k++) {
                    JSONArray row = data.getJSONArray(k);
                    XWPFTableRow nextRow = table.createRow();
                    for (int l = 0; l < row.length(); l++) {

                        nextRow.getCell(l).setText(row.getString(l));
                        if(l != row.length() -1 && check.equals("")) nextRow.addNewTableCell();
                    }
                }
            } catch (JSONException ignored) {

            }
            catch (Exception e){
                e.printStackTrace();
            }

            //image
            try {
                JSONObject image = subsection.getJSONObject("image");
                XWPFParagraph imageParagraph = doc.createParagraph();
                XWPFRun imageRun = imageParagraph.createRun();
                String alignment = image.getString("align");
                switch (alignment) {
                    case "C" -> imageParagraph.setAlignment(ParagraphAlignment.CENTER);
                    case "R" -> imageParagraph.setAlignment(ParagraphAlignment.RIGHT);
                    case "L" -> imageParagraph.setAlignment(ParagraphAlignment.LEFT);
                }
                String url = image.getString("data");
//                try{
//                    //saveFileFromUrl(url, "src\\main\\resources\\temp.png");
//                }
//                //catch (MalformedURLException ignored){}
                File imageFile = Path.of("src\\main\\resources\\temp.png").toFile();
                FileInputStream imageData = new FileInputStream(imageFile);
                int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
                String imageFileName = imageFile.getName();
                int width = image.getInt("maxWidth");
                int height = image.getInt("maxHeight");
                imageRun.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));

            } catch (JSONException ignored) {
            }
            catch (Exception e){
                e.printStackTrace();
            }

        }
    }

    private void Attachments(JSONObject attachments) throws InvalidFormatException, IOException {

        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addBreak(BreakType.PAGE);
        run.setText(attachments.getString("title"));
        if(attachments.getBoolean("toc")){
            attachmentsTOC(attachments);
        }
        //saveFileFromUrl("https://www.africau.edu/images/default/sample.pdf", "src\\main\\resources\\sample.pdf");

        JSONArray items = attachments.getJSONArray("items");
        for (int i = 0; i < items.length(); i++) {
            JSONObject item = items.getJSONObject(i);
            paragraph = doc.createParagraph();
            run = paragraph.createRun();
            run.setText(item.getString("title"));

            JSONArray files = item.getJSONArray("files");
            for (int j = 0; j < files.length(); j++) {
                String url = files.getString(j);
                //saveFileFromUrl(url, "src\\main\\resources\\temp.pdf");
                int pageNumber = 0;
                String fileName = "";
                try {
                    //String sourceDir = "C:\\Users\\limmi\\Documents\\OOPP\\gs-uploading-files\\complete\\src\\main\\resources\\sample.pdf";
                    //String destinationDir = "C:\\Users\\limmi\\OneDrive\\Documenten\\JSON_to_DOCX\\complete\\src\\main\\PDF_images\\"; // converted images from pdf document are saved here

                    File sourceFile = Path.of("src\\main\\resources\\sample.pdf").toFile();
                    File destinationFile = Path.of("src\\main\\PDF_images").toFile();

                    //file found checker
                    try(InputStream stream = new FileInputStream(sourceFile)){
                        System.out.println("File found!");
                    }
                    catch (FileNotFoundException e){
                        e.printStackTrace();
                    }

                    if (!destinationFile.exists()) {
                        destinationFile.mkdir();
                        System.out.println("Folder Created -> "+ destinationFile.getAbsolutePath());
                    }
                    if (sourceFile.exists()) {
                        System.out.println("Images copied to Folder Location: "+ destinationFile.getAbsolutePath());
                        PDDocument document = PDDocument.load(sourceFile);
                        PDFRenderer pdfRenderer = new PDFRenderer(document);

                        pageNumber = document.getNumberOfPages();
                        System.out.println("Total files to be converting -> "+ pageNumber);

                        fileName = sourceFile.getName().replace(".pdf", "");
                        String fileExtension= "png";
                        int dpi = 100;

                        for (int k = 0; k < pageNumber; ++k) {
                            File outPutFile = new File(destinationFile.getAbsolutePath() + "/" +  fileName +"_"+ (k+1) +"."+ fileExtension);
                            BufferedImage bImage = pdfRenderer.renderImageWithDPI(k, dpi, ImageType.RGB);
                            ImageIO.write(bImage, fileExtension, outPutFile);
                        }

                        document.close();
                        System.out.println("Converted Images are saved at -> "+ destinationFile.getAbsolutePath());
                    } else {
                        System.err.println(sourceFile.getName() +" File not exists");
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }

                for (int k = 1; k <= pageNumber; k++) {
                    File pdf = Path.of("src", "main", "PDF_images", fileName + "_" + k + ".png").toFile();
                    InputStream stream = new FileInputStream(pdf);
                    String imageFileName = pdf.getName();
                    run.addPicture(stream, XWPFDocument.PICTURE_TYPE_PNG, imageFileName, Units.toEMU(450), Units.toEMU(600));

                }
            }
            if(i != items.length() - 1)
                run.addBreak(BreakType.PAGE);
        }

    }

    private void attachmentsTOC(JSONObject attachments) {
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Table of content for the appencides:");

        JSONArray items = attachments.getJSONArray("items");
        for (int i = 0; i < items.length(); i++) {
            JSONObject item = items.getJSONObject(i);
            paragraph = doc.createParagraph();
            run = paragraph.createRun();
            run.setText("   - " + item.getString("title"));

        }

    }

    private void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onOffNull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onOffNull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onOffNull);

        // style defines a heading of the given level
        CTPPrGeneral ppr = CTPPrGeneral.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }

    public static void saveFileFromUrl(String URL, String destinationFile) throws IOException {
        URL url = new URL(URL);
        InputStream is = url.openStream();
        OutputStream os = new FileOutputStream(Path.of(destinationFile).toFile());

        byte[] b = new byte[2048];
        int length;

        while ((length = is.read(b)) != -1) {
            os.write(b, 0, length);
        }

        is.close();
        os.close();
    }
    private static void setParagraphShading(XWPFParagraph paragraph, String rgb) {
        if (paragraph.getCTP().getPPr() == null) paragraph.getCTP().addNewPPr();
        if (paragraph.getCTP().getPPr().getShd() != null) paragraph.getCTP().getPPr().unsetShd();
        paragraph.getCTP().getPPr().addNewShd();
        paragraph.getCTP().getPPr().getShd().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd.CLEAR);
        paragraph.getCTP().getPPr().getShd().setColor("auto");
        paragraph.getCTP().getPPr().getShd().setFill(rgb);
    }

    private void addlongTextToRun(String data, XWPFRun run){
        if (data.contains("\n")) {
            String[] lines = data.split("\n");
            run.setText(lines[0], 0); // set first line into XWPFRun
            for(int i=1;i<lines.length;i++){
                // add break and insert new text
                run.addBreak();
                run.setText(lines[i]);
            }
        } else {
            run.setText(data, 0);
        }
    }
    private void addLongTextToCell(String data, XWPFTableCell cell){
        if (data.contains("\n")) {
            String[] lines = data.split("\n");
            cell.setText(lines[0]); // set first line into XWPFRun
            for(int i=1;i<lines.length;i++){
                // add break and insert new text
                //cell.set;
                cell.setText(lines[i]);
            }
        } else {
            cell.setText(data);
        }
    }


}



