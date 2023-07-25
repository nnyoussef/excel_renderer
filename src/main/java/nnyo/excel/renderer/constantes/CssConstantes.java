package nnyo.excel.renderer.constantes;

import org.apache.poi.ss.usermodel.BorderStyle;

import java.util.HashMap;
import java.util.Map;

public class CssConstantes {

    public static final String BACKGROUND = "background";

    public static final String BORDER = "border";

    public static final String BORDER_DASH = "border-";

    public static final String COLOR = "color";

    public static final String TEXT_ALIGN = "text-align";

    public static final String RGB_WHITE = "rgb(255,255,255)";

    public static final String RGB_BLACK = "rgb(255,255,255)";

    public static final String HEX_WHITE = "FFFFFF";

    public static final String BORDER_NONE = "none";

    public static final String LEFT = "left";

    public static final String HEX_LAVENDER_RED = "B0054B";

    public static final int SECTION_TITLE_FONT_SIZE = 12;

    public static final String HEX_GREEN = "2C876E";

    public static final int REPORT_TITLE_FONT_SIZE = 10;

    public static final String TOP = "top";

    public static final String BOTTOM = "bottom";

    public static final String ONE_PX = "1px";

    public static final String TWO_PX = "2px";

    public static final String SOLID = "solid";

    public static final String NORMAL_GRAY_BORDER_COLOR_HEX = "7F7F7F";

    public static final String HASHBANG = "#";

    public static final String DASH = "-";

    public static final String TEXT_ALIGN_CENTER = "center";

    public static final String TEXT_ALIGN_RIGHT = "right";

    public static final String TEXT_ALIGN_LEFT = "left";

    public static final String FONT_WEIGHT = "font-weight";

    public static final String BOLD = "bold";

    public static final Map<String, String> DEFAULT_CSS_PROPERTIES = new HashMap<>() {{
        put(BACKGROUND, RGB_WHITE);
        put(BORDER, BORDER_NONE);
        put(COLOR, RGB_BLACK);
        put(TEXT_ALIGN, TEXT_ALIGN_LEFT);
    }};

    public static final Map<String, BorderStyle> CSS_BORDER_VALUE_TO_BORDER_STYLE_MAP = new HashMap<>() {{
        put("1px solid", BorderStyle.THIN);
        put("solid 1px", BorderStyle.THIN);
        put("2px solid", BorderStyle.MEDIUM);
        put("solid 2px", BorderStyle.MEDIUM);

        put("1px dashed", BorderStyle.DASHED);
        put("dashed 1px", BorderStyle.DASHED);
        put("2px dashed", BorderStyle.MEDIUM_DASHED);
        put("dashed 2px", BorderStyle.MEDIUM_DASHED);
    }};

    private CssConstantes() throws InstantiationException {
        throw new InstantiationException("Class of Constants not Instantiable");
    }
}
