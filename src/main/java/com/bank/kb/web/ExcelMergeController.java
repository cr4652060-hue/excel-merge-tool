package com.bank.kb.web;

import com.bank.kb.service.excel.ExcelMergeResult;
import com.bank.kb.service.excel.ExcelMergeService;
import com.bank.kb.service.excel.ExcelTemplateInfo;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelMergeController {

    private final ExcelMergeService excelMergeService;

    @PostMapping("/template")
    public ExcelTemplateInfo uploadTemplate(@RequestParam("file") MultipartFile file) {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("模板文件不能为空。");
        }
        return excelMergeService.analyzeTemplate(file);
    }

    @PostMapping("/merge")
    public ExcelMergeResult mergeFiles(@RequestParam("files") List<MultipartFile> files) {
        return excelMergeService.mergeFiles(files);
    }

    @GetMapping("/export")
    public ResponseEntity<byte[]> exportMerged() {
        byte[] bytes = excelMergeService.exportMerged();
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"merged_result.xlsx\"")
                .body(bytes);
    }
}