package io.kojisaiki.ApachePoiSample.web.rest;

import io.kojisaiki.ApachePoiSample.service.ExcelService;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

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
    public String getSimpleExcel() throws Exception {
        String ret;

        ret = this.excelService.generateSimpleExcel();

        return ret;
    }
}
