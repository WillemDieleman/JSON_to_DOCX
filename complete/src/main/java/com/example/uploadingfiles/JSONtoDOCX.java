package com.example.uploadingfiles;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;

import org.apache.xmlbeans.XmlException;
import org.json.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigInteger;
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

        saveImage("https://cdn.pixabay.com/photo/2015/08/23/09/22/banner-902589__340.jpg","src\\main\\resources\\header.jpg");
        File image = Path.of("src/main/resources/header.jpg").toFile();
        FileInputStream imageData = new FileInputStream(image);
        int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
        String imageFileName = image.getName();
        int width = 450;
        int height = 50;
        run.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));
        //text
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
            title.setBold(true);
            title.setFontSize(18);
            title.setText((String) section.get("title"));
            try {
                String introText = section.getString("introText");
                texts = doc.createParagraph();
                XWPFRun introTextRun = texts.createRun();
                introTextRun.setText(introText);
            } catch (Exception e) {
                //no introText;
            }

            JSONArray subsections = section.getJSONArray("subSections");

            try {
                JSONObject check = (JSONObject) subsections.get(0);
                check.getString("title");
            } catch (JSONException ignored) {
            }

            SubSections(subsections);


        }

    }

    private void SubSections(JSONArray subsections) {
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
                subTitleRun.setText(title);
                subtitle.setStyle("heading 2");
            } catch (JSONException ignored) {

            }
            //intro text stuff
            try {
                String introText = subsection.getString("introText");
                subtext = doc.createParagraph();
                XWPFRun subIntroText = subtext.createRun();
                subIntroText.addTab();
                subIntroText.setText(introText);
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
                XWPFStyle style = styles.getStyleWithName("Grid Table 4");
                table.setStyleID(style.getStyleId());

                JSONArray headers = tableJSON.getJSONArray("headerRows");
                XWPFTableRow tableRowOne = table.getRow(0);
                tableRowOne.getCell(0).setText(headers.getString(0));
                for (int k = 1; k < headers.length(); k++) {
                    tableRowOne.addNewTableCell().setText(headers.getString(k));
                }
                JSONArray data = tableJSON.getJSONArray("dataRows");
                for (int k = 0; k < data.length(); k++) {
                    JSONArray row = data.getJSONArray(k);
                    XWPFTableRow nextRow = table.createRow();
                    for (int l = 0; l < row.length(); l++) {
                        nextRow.getCell(l).setText(row.getString(l));
                    }
                }
            } catch (JSONException ignored) {

            }
            catch (Exception e){
                e.printStackTrace();
            }

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

                saveImage("https://www.agilesparks.com/wp-content/uploads/2022/08/Java_logo_icon.png", "src\\main\\resources\\JAVA.jpeg");
                File imageFile = Path.of("src\\main\\resources\\JAVA.jpeg").toFile();
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

    public static void saveImage(String imageUrl, String destinationFile) throws IOException {
        URL url = new URL(imageUrl);
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


}



