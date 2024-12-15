import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.xwpf.usermodel.*;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.Rectangle;
import java.io.*;

@WebServlet("/convertPdfToWord")
@MultipartConfig(
        fileSizeThreshold = 1024 * 1024 * 10, // 10MB
        maxFileSize = 1024 * 1024 * 10,       // 10MB
        maxRequestSize = 1024 * 1024 * 20     // 20MB
)
public class PdfToWordServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        // Lấy file PDF từ request
        InputStream pdfInputStream = request.getPart("pdfFile").getInputStream();

        if (pdfInputStream == null || pdfInputStream.available() == 0) {
            response.sendError(HttpServletResponse.SC_BAD_REQUEST, "File PDF không hợp lệ hoặc trống.");
            return;
        }

        // Tạo tài liệu Word mới
        try (XWPFDocument wordDocument = new XWPFDocument();
             PDDocument pdfDocument = PDDocument.load(pdfInputStream)) {

            // Lặp qua từng trang trong PDF và trích xuất văn bản theo vùng
            int numberOfPages = pdfDocument.getNumberOfPages();
            for (int i = 0; i < numberOfPages; i++) {
                PDPage page = pdfDocument.getPage(i);
                extractTextByArea(page, wordDocument);
            }

            // Thiết lập phản hồi HTTP
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setHeader("Content-Disposition", "attachment; filename=\"converted.docx\"");

            // Ghi tài liệu Word vào output stream của phản hồi
            try (ServletOutputStream out = response.getOutputStream()) {
                wordDocument.write(out);
                out.flush();
            }
        } catch (IOException e) {
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Lỗi khi xử lý file PDF.");
            e.printStackTrace();
        }
    }

    // Phương thức trích xuất văn bản từ trang PDF theo vùng
    private void extractTextByArea(PDPage page, XWPFDocument wordDocument) throws IOException {
        PDFTextStripperByArea stripper = new PDFTextStripperByArea();

        // Định nghĩa một vùng bao phủ toàn bộ trang
        Rectangle rect = new Rectangle(0, 0, (int) page.getMediaBox().getWidth(), (int) page.getMediaBox().getHeight());
        stripper.addRegion("fullPage", rect);
        
        // Áp dụng stripper trên trang
        stripper.extractRegions(page);
        String pageText = stripper.getTextForRegion("fullPage");

        if (!pageText.trim().isEmpty()) {
            XWPFParagraph paragraph = wordDocument.createParagraph();
            paragraph.setSpacingBetween(1.5);
            paragraph.setAlignment(ParagraphAlignment.BOTH);
            XWPFRun run = paragraph.createRun();
            run.setText(pageText.trim());
        }
    }
}
