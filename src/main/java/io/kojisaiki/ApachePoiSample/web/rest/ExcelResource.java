package io.kojisaiki.ApachePoiSample.web.rest;

import io.kojisaiki.ApachePoiSample.service.ExcelService;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@RestController
@RequestMapping("/api/excel")
public class ExcelResource {

    private ExcelService excelService;

    public ExcelResource(
            ExcelService excelService
    ) {
        this.excelService = excelService;
    }

    @GetMapping("/simple")
    public ResponseEntity<Resource> getSimpleExcel() throws Exception {
        String generatedFilePath = this.excelService.generateSimpleExcel();

        File generatedFile = new File(generatedFilePath);
        Path absolutePath = Paths.get(generatedFile.getAbsolutePath());

        ByteArrayResource resource = new ByteArrayResource(Files.readAllBytes(absolutePath));

        return ResponseEntity.ok()
                .header("Content-disposition", "attachment;filename=" + generatedFile.getName())
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .contentLength(generatedFile.length())
                .body(resource);
    }
}
