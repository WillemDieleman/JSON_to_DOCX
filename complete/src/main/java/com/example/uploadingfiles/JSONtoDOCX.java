package com.example.uploadingfiles;

import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
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

    private int widthInt = 0;
    private int tableHeigth = 0;

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

    private void HeaderFooter() throws Exception {
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
        BufferedImage BI = ImageIO.read(image);
        FileInputStream imageData = new FileInputStream(image);
        int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
        String imageFileName = image.getName();
        int width = BI.getWidth()/4;
        int height = BI.getHeight()/4;
        run.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));
        CTDrawing drawing = run.getCTR().getDrawingArray(0);
        CTGraphicalObject graphicalobject = drawing.getInlineArray(0).getGraphic();
        CTAnchor anchor = getAnchorWithGraphic(graphicalobject, "header.jpeg",
                Units.toEMU(width), Units.toEMU(height),
                Units.toEMU(-60), Units.toEMU(-40));
        drawing.setAnchorArray(new CTAnchor[]{anchor});
        drawing.removeInline(0);
        run.addBreak();
        run.setText(JSONHeader.getString("headerText"));
        //footer
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFTable footerTable = footer.createTable(1,1);
        footerTable.setStyleID(doc.getStyles().getStyleWithName("Table Grid").getStyleId());
        XWPFTableRow row = footerTable.getRow(0);
        row.setHeight(1440);
        XWPFTableCell cell = row.getCell(0);
        footerTable.setWidth("132%");

        footerTable.setTableAlignment(TableRowAlign.CENTER);
        cell.setColor("005B82");
        cell.addParagraph(doc.createParagraph());

        XWPFParagraph footerParagraph = cell.getParagraphs().get(0);
        run = footerParagraph.createRun();
        footerParagraph.setAlignment(ParagraphAlignment.CENTER);
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

        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();
        pageMar.setFooter(BigInteger.valueOf(0)); //28.4 pt * 20 = 568 = 28.4 pt footer from bottom

        footerTable.setBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        footerTable.setRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        footerTable.setLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        footerTable.setTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        footerTable.setInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        footerTable.setInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);

    }

    private void Sections() {
        JSONArray sectionsArray = JSON.getJSONArray("sections");
        titles = doc.createParagraph();
        XWPFParagraph texts = null;



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

            doc.createParagraph().createRun().addBreak(BreakType.PAGE);
        }


    }

    private void SubSections(JSONArray subsections, int i) {
        XWPFParagraph subtitle;
        XWPFParagraph subtext;

        boolean hasHeaders = false;
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
                table(tableJSON);


            }
            catch (JSONException ignored){}
            catch (Exception e){
                e.printStackTrace();
            }

            //image
            try {
                JSONObject image = subsection.getJSONObject("image");
                image(image);
            }
            catch (JSONException ignored) {}
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
                table(tableJSON);
            } catch (JSONException ignored) {

            }
            catch (Exception e){
                e.printStackTrace();
            }

            //image
            try {
                JSONObject image = subsection.getJSONObject("image");
                image(image);

            } catch (JSONException ignored) {
            }
            catch (Exception e){
                e.printStackTrace();
            }

        }
    }

    private void table(JSONObject tableJSON){
        boolean hasHeaders = false;
        tableHeigth = 0;
        XWPFTable table = doc.createTable();
        XWPFStyles styles = doc.getStyles();
        XWPFStyle style = styles.getStyleWithName("Table Grid");
        table.setStyleID(style.getStyleId());
        String check = "";
        String width = tableJSON.getString("width");
        widthInt = Integer.parseInt(width.replace("%", ""));
        table.setWidth(width);
        table.setBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        table.setRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        table.setLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        table.setTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        table.setInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);
        table.setInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, null);

        try{
            JSONArray headers = tableJSON.getJSONArray("dataHeaders");
            check = headers.getString(0);
            hasHeaders = true;
            XWPFTableRow tableRowOne = table.getRow(0);
            XWPFTableCell cell = tableRowOne.getCell(0);
            cell.setColor("005B82");
            cell.setText(check);
            for (int k = 1; k < headers.length(); k++) {
                cell = tableRowOne.addNewTableCell();
                cell.setColor("005B82");
                cell.setText(headers.getString(k));
            }

            tableHeigth += tableRowOne.getHeight();

        }
        catch (JSONException e){
            //no reader rows
        }

        JSONArray data = tableJSON.getJSONArray("dataRows");
        XWPFTableRow tableRowOne;
        if(hasHeaders){
            table.createRow();
            tableRowOne = table.getRow(1);
            for (int k = 0; k < data.getJSONArray(0).length(); k++) {
                tableRowOne.getCell(k).setText(data.getJSONArray(0).getString(k));
            }
        }
        else{
            tableRowOne = table.getRow(0);
            for (int k = 1; k < data.getJSONArray(0).length(); k++) {
                tableRowOne.addNewTableCell().setText(data.getJSONArray(0).getString(k));
            }
            tableRowOne.getCell(0).setText(data.getJSONArray(0).getString(0));
        }
        tableRowOne.setHeight((int)(1440/3));

        tableHeigth += tableRowOne.getHeight();


        for (int k = 1; k < data.length(); k++) {
            JSONArray row = data.getJSONArray(k);
            XWPFTableRow nextRow = table.createRow();
            for (int l = 0; l < row.length(); l++) {
                XWPFTableCell cell = nextRow.getCell(l);
                cell.setWidth(widthInt/row.length() + "%");
                XWPFRun run = cell.getParagraphs().get(0).createRun();
                addlongTextToRun(row.getString(l), run);

            }
            nextRow.setHeight((int)(1440/3));
            nextRow.getCtRow().getTrPr().getTrHeightArray(0).setHRule(STHeightRule.AT_LEAST);
            tableHeigth += nextRow.getHeight();
        }
        data = tableJSON.getJSONArray("dataTotals");
        for (int k = 0; k < data.length(); k++) {
            JSONArray row = data.getJSONArray(k);
            XWPFTableRow nextRow = table.createRow();
            for (int l = 0; l < row.length(); l++) {
                nextRow.getCell(l).setText(row.getString(l));

            }
            tableHeigth += nextRow.getHeight();
        }

    }

    private void image(JSONObject image) throws Exception {
        XWPFParagraph imageParagraph = doc.createParagraph();
        XWPFRun imageRun = imageParagraph.createRun();
        String alignment = image.getString("align");
        switch (alignment) {
            case "C" -> imageParagraph.setAlignment(ParagraphAlignment.CENTER);
            case "R" -> imageParagraph.setAlignment(ParagraphAlignment.RIGHT);
            case "L" -> imageParagraph.setAlignment(ParagraphAlignment.LEFT);
        }
        String url = image.getString("data");
        try{
            saveFileFromUrl(url, "src\\main\\resources\\temp.png");
        }
        catch (IOException ignored){}
        File imageFile = Path.of("src\\main\\resources\\temp.png").toFile();

        String widthString = image.getString("width");
        int factor = Integer.parseInt(widthString.replace("%", ""));
        FileInputStream imageData = new FileInputStream(imageFile);
        BufferedImage bi = ImageIO.read(imageData);
        int imageType = XWPFDocument.PICTURE_TYPE_PNG;
        String imageFileName = imageFile.getName();
        int cm = Units.EMU_PER_CENTIMETER;
        double full = cm * (21.59 - 5);
        double width = full * factor / 100;
        double height = (double)bi.getHeight()/ (double)bi.getWidth() * width;

        imageRun.addPicture(new FileInputStream(imageFile), imageType, imageFileName, (int)width, (int)height);
        if(widthInt != 0 && widthInt + factor <= 100){
            CTDrawing drawing = imageRun.getCTR().getDrawingArray(0);
            CTGraphicalObject graphicalobject = drawing.getInlineArray(0).getGraphic();
            CTAnchor anchor = getAnchorWithGraphic(graphicalobject, "header.jpeg",
                    (int)width, (int)height,
                    (int)full * widthInt/ 100, (int) -(5*cm));
            drawing.setAnchorArray(new CTAnchor[]{anchor});
            drawing.removeInline(0);
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
                    }
                    catch (FileNotFoundException e){
                        e.printStackTrace();
                    }

                    if (!destinationFile.exists()) {
                        destinationFile.mkdir();
                    }
                    if (sourceFile.exists()) {
                        PDDocument document = PDDocument.load(sourceFile);
                        PDFRenderer pdfRenderer = new PDFRenderer(document);

                        pageNumber = document.getNumberOfPages();

                        fileName = sourceFile.getName().replace(".pdf", "");
                        String fileExtension= "png";
                        int dpi = 100;

                        for (int k = 0; k < pageNumber; ++k) {
                            File outPutFile = new File(destinationFile.getAbsolutePath() + "/" +  fileName +"_"+ (k+1) +"."+ fileExtension);
                            BufferedImage bImage = pdfRenderer.renderImageWithDPI(k, dpi, ImageType.RGB);
                            ImageIO.write(bImage, fileExtension, outPutFile);
                        }

                        document.close();
                    } else {
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
    private CTAnchor getAnchorWithGraphic(CTGraphicalObject graphicalobject, String drawingDescr, int width, int height, int left, int top) throws Exception {

        String anchorXML =
                "<wp:anchor xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                        +"simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"1\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
                        +"<wp:simplePos x=\"0\" y=\"0\"/>"
                        +"<wp:positionH relativeFrom=\"column\"><wp:posOffset>"+left+"</wp:posOffset></wp:positionH>"
                        +"<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>"+top+"</wp:posOffset></wp:positionV>"
                        +"<wp:extent cx=\""+width+"\" cy=\""+height+"\"/>"
                        +"<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
                        +"<wp:wrapTight wrapText=\"bothSides\">"
                        +"<wp:wrapPolygon edited=\"0\">"
                        +"<wp:start x=\"0\" y=\"0\"/>"
                        +"<wp:lineTo x=\"0\" y=\"21600\"/>" //Square polygon 21600 x 21600 leads to wrap points in fully width x height
                        +"<wp:lineTo x=\"21600\" y=\"21600\"/>"// Why? I don't know. Try & error ;-).
                        +"<wp:lineTo x=\"21600\" y=\"0\"/>"
                        +"<wp:lineTo x=\"0\" y=\"0\"/>"
                        +"</wp:wrapPolygon>"
                        +"</wp:wrapTight>"
                        +"<wp:docPr id=\"1\" name=\"Drawing 0\" descr=\""+drawingDescr+"\"/><wp:cNvGraphicFramePr/>"
                        +"</wp:anchor>";

        CTDrawing drawing = CTDrawing.Factory.parse(anchorXML);
        CTAnchor anchor = drawing.getAnchorArray(0);
        anchor.setGraphic(graphicalobject);
        return anchor;
    }


}



