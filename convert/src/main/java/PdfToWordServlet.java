import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import javax.imageio.ImageIO;
import java.util.List;
import java.util.stream.Collectors;

@WebServlet("/convertPdfToWord")
@MultipartConfig(
        fileSizeThreshold = 1024 * 1024, // 1MB
        maxFileSize = 1024 * 1024 * 5,   // 5MB
        maxRequestSize = 1024 * 1024 * 10 // 10MB
)
public class PdfToWordServlet extends HttpServlet {
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        // Lấy file PDF từ request
        InputStream pdfInputStream = request.getPart("pdfFile").getInputStream();

        // Tạo tài liệu Word mới
        try (XWPFDocument wordDocument = new XWPFDocument();
             PDDocument pdfDocument = PDDocument.load(pdfInputStream)) {

            PDFTextStripper pdfStripper = new PDFTextStripper();

            // Lặp qua từng trang trong PDF
            int numberOfPages = pdfDocument.getNumberOfPages();
            for (int i = 0; i < numberOfPages; i++) {
                pdfStripper.setStartPage(i + 1);
                pdfStripper.setEndPage(i + 1);

                // Trích xuất văn bản từ trang hiện tại
                String pageText = pdfStripper.getText(pdfDocument);

                // Thêm văn bản vào tài liệu Word
                if (!pageText.trim().isEmpty()) {
                    XWPFParagraph paragraph = wordDocument.createParagraph();
                    paragraph.setSpacingBetween(1.5);
                    paragraph.setAlignment(ParagraphAlignment.BOTH);
                    XWPFRun run = paragraph.createRun();
                    run.setText(pageText.trim());
                }

                // Trích xuất hình ảnh từ trang hiện tại
                PDPage page = pdfDocument.getPage(i);
                PDResources resources = page.getResources();

             // Trích xuất hình ảnh từ trang hiện tại
                for (COSName cosName : resources.getXObjectNames()) {
                    if (resources.isImageXObject(cosName)) {
                        PDImageXObject imageObject = (PDImageXObject) resources.getXObject(cosName);
                        BufferedImage image = imageObject.getImage();

                        // Chuyển BufferedImage thành byte array
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        ImageIO.write(image, "png", baos);
                        byte[] imageBytes = baos.toByteArray();

                        // Thêm hình ảnh vào tài liệu Word
                        XWPFParagraph imageParagraph = wordDocument.createParagraph();
                        XWPFRun imageRun = imageParagraph.createRun();
                        imageRun.addPicture(
                                new ByteArrayInputStream(imageBytes),
                                XWPFDocument.PICTURE_TYPE_PNG,
                                "image.png",
                                Units.toEMU(image.getWidth()),
                                Units.toEMU(image.getHeight())
                        );
                        imageParagraph.setSpacingBetween(1.5);
                    }
                }


            }

            // Thiết lập phản hồi HTTP
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setHeader("Content-Disposition", "attachment; filename=\"converted.docx\"");

            // Ghi tài liệu Word vào output stream của phản hồi
            try (ServletOutputStream out = response.getOutputStream()) {
                wordDocument.write(out);
                out.flush();
            }
        } catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
}
