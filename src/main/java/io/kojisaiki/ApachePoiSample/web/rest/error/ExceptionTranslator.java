package io.kojisaiki.ApachePoiSample.web.rest.error;

import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;

@ControllerAdvice
public class ExceptionTranslator {

    @ExceptionHandler
    public String handleException(Exception e) {
        return "exception! : " + e.getMessage();
    }
}
