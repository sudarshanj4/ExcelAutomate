package com.example.excel_automate.dtos;

import java.util.List;

public class RequestDto {
    private List<String> language_name;
    private String engine_name;
    private String version;
    private String filePathUrl;
    private String Destination_filePathUrl;

    public List<String> getLanguage_name() {
        return language_name;
    }

    public void setLanguage_name(List<String> language_name) {
        this.language_name = language_name;
    }

    public String getEngine_name() {
        return engine_name;
    }

    public void setEngine_name(String engine_name) {
        this.engine_name = engine_name;
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public String getFilePathUrl() {
        return filePathUrl;
    }

    public void setFilePathUrl(String filePathUrl) {
        this.filePathUrl = filePathUrl;
    }

    public String getDestination_filePathUrl() {
        return Destination_filePathUrl;
    }

    public void setDestination_filePathUrl(String destination_filePathUrl) {
        Destination_filePathUrl = destination_filePathUrl;
    }
}
