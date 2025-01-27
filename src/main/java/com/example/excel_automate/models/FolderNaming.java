package com.example.excel_automate.models;

public class FolderNaming {
 LanguageType languageType;
    public String folderName(String language){
        if(language.equals("English (United States)")){
            return "_Stand";
        }else if(language.equals("French (France)")){
            return "_Fr";
        }else if(language.equals("Italian (Italy)")){
            return "_It";
        }else if(language.equals("Russian (Russia)")){
            return "_Ru";
        }else if(language.equals("Japanese (Japan)")){
            return "_Ja";
        }else if(language.equals("Chinese (Simplified, China)")){
            return "_ZhCn";
        }else if(language.equals("Chinese (Traditional, Taiwan)")){
            return "_ZhTw";
        }else if(language.equals("Arabic (Oman)")){
            return "_Ar";
        }
        return null;
    }
}
