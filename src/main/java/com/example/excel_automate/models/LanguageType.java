package com.example.excel_automate.models;

import java.util.ArrayList;
import java.util.List;

public class LanguageType {

    private final String ReferenceText = "ReferenceText";
    private final String English = "English (United States)";
    private final String German = "German (Germany)";
    private final String Spanish = "Spanish (Spain, International Sort)";
    private final String Portuguese = "Portuguese (Portugal)";
    private final String French = "French (France)";
    private final String Italian = "Italian (Italy)";
    private final String Russian = "Russian (Russia)";
    private final String Japanese = "Japanese (Japan)";
    private final String ChineseSimplified = "Chinese (Simplified, China)";
    private final String ChineseTraditional = "Chinese (Traditional, Taiwan)";
    private final String Arabic = "Arabic (Oman)";

    public List<String> addLanguagesBasedOnCondition(String language_name) {

        // Use a mutable list (ArrayList)
        List<String> commonLanguages = new ArrayList<>();
        commonLanguages.add("Key");
        commonLanguages.add(ReferenceText);
        commonLanguages.add(English);
        commonLanguages.add(German);
        commonLanguages.add(Spanish);
        commonLanguages.add(Portuguese);

        // Add languages based on the condition
        if (language_name.equals("Standard")) {
            // Return the list as is for "Standard"
            return commonLanguages;
        } else if (language_name.equals(French)) {
            commonLanguages.add(French);
        } else if (language_name.equals(Italian)) {
            commonLanguages.add(Italian);
        } else if (language_name.equals(Russian)) {
            commonLanguages.add(Russian);
        } else if (language_name.equals(Japanese)) {
            commonLanguages.add(Japanese);
        } else if (language_name.equals(ChineseSimplified)) {
            commonLanguages.add(ChineseSimplified);
        } else if (language_name.equals(ChineseTraditional)) {
            commonLanguages.add(ChineseTraditional);
        } else if (language_name.equals(Arabic)) {
            commonLanguages.add(Arabic);
        }

        return commonLanguages;
    }

    // Getters for the language names
    public String getReferenceText() {
        return ReferenceText;
    }

    public String getEnglish() {
        return English;
    }

    public String getGerman() {
        return German;
    }

    public String getSpanish() {
        return Spanish;
    }

    public String getPortuguese() {
        return Portuguese;
    }

    public String getFrench() {
        return French;
    }

    public String getItalian() {
        return Italian;
    }

    public String getRussian() {
        return Russian;
    }

    public String getJapanese() {
        return Japanese;
    }

    public String getChineseSimplified() {
        return ChineseSimplified;
    }

    public String getChineseTraditional() {
        return ChineseTraditional;
    }

    public String getArabic() {
        return Arabic;
    }
}
