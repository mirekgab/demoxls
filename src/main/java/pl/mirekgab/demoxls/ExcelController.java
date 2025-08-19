package pl.mirekgab.demoxls;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Controller
public class ExcelController {

    @GetMapping("/")
    public String index() {
        return "upload";
    }

    @PostMapping("/upload")
    public String uploadExcel(@RequestParam("file") MultipartFile file, Model model) {
        List<String> headers = new ArrayList<>();
        List<List<String>> rows = new ArrayList<>();

        try (InputStream is = file.getInputStream()) {
            Workbook workbook = WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(0);

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            boolean isHeader = true;
            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    // If it's a formula cell, evaluate it first
                    String cellValue = formatter.formatCellValue(cell, evaluator);
                    rowData.add(cellValue);
                }
                if (isHeader) {
                    headers = rowData;
                    isHeader = false;
                } else {
                    rows.add(rowData);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
            model.addAttribute("message", "Error reading file: " + e.getMessage());
            return "upload";
        }

        model.addAttribute("headers", headers);
        model.addAttribute("rows", rows);
        return "table";
    }

}

