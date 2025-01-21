package com.example.excel_automate.dtos;

public class RequestDto {
    private String language_name;
    private String enginee_name;
    private String version;
    private String filePathUrl;
    private String Destination_filePathUrl;

    public String getLanguage_name() {
        return language_name;
    }

    public void setLanguage_name(String language_name) {
        this.language_name = language_name;
    }

    public String getEnginee_name() {
        return enginee_name;
    }

    public void setEnginee_name(String enginee_name) {
        this.enginee_name = enginee_name;
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
