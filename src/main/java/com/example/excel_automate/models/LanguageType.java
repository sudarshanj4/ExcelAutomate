package com.example.excel_automate.models;

import java.util.ArrayList;
import java.util.List;

public class LanguageType {


    private final String ReferenceText = "Reference Text";
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


    public List<String> addLanguagesBasedOnCondition(String language_name){

        List<String> commonLanguages = List.of(ReferenceText, English, German, Spanish, Portuguese);
        List<String> list = new ArrayList<String>();
        if(language_name.equals("Standard")){
            return commonLanguages;
        }else if(language_name.equals(French)){
            list.add(French);
        }
        else if(language_name.equals(Italian)){
            list.add(Italian);
        }
        else if (language_name.equals(Russian)) {
            list.add(Russian);
        }
        else if (language_name.equals(Japanese)) {
            list.add(Japanese);
        }
        else if (language_name.equals(ChineseSimplified)) {
            list.add(ChineseSimplified);
        }
        else if (language_name.equals(ChineseTraditional)) {
            list.add(ChineseTraditional);
        }
        else if (language_name.equals(Arabic)) {
            list.add(Arabic);
        }


        return list;
    }
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
